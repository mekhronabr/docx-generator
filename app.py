"""Streamlit: Excel -> несколько общих DOCX, готовых к печати."""
from collections import deque
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile, ZIP_DEFLATED
import hashlib
import math
import re

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
from jinja2 import Environment, StrictUndefined
from lxml import etree

BASE = Path(__file__).resolve().parent
TEMPLATE = BASE / 'template.docx'
DEFAULT_BATCH_SIZE = 1000
MAX_WORKERS = 2
MAX_READY_BYTES = 64 * 1024 * 1024
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
MAX_ROWS = 100_000
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
COLUMNS = ['ФИО Студента', 'Курс', 'Группа', 'Специальность',
           'Факультет', 'ФИО Куратора', 'Телефон Куратора', 'ФИО Зам. декана',
           'Телефон Зам. декана', 'ФИО Декана', 'Телефон Декана']
ALIASES = dict(zip(
    ['ному насаби донишҷӯ', 'курс', 'гурӯҳ', 'ихтисос', 'куратори гурӯҳ',
     'рақами телефонии куратори гурӯҳ', 'муовини декан оид ба тарбия',
     'рақами телефонии муовини декан оид ба тарбия', 'декани факултет',
     'рақами телефонии декани факултет', 'факултет'],
    ['ФИО Студента', 'Курс', 'Группа', 'Специальность', 'ФИО Куратора',
     'Телефон Куратора', 'ФИО Зам. декана', 'Телефон Зам. декана',
     'ФИО Декана', 'Телефон Декана', 'Факультет']))


def clean(value):
    if value is None or pd.isna(value):
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return re.sub(r'\s+', ' ', str(value)).strip()


def specialty_code(group):
    """Код из группы: 1.53.01.05тб → 1.53.01.05, 344 → 344."""
    match = re.fullmatch(r'(\d+(?:[.-]\d+)*)\s*[^\W\d_]*', clean(group))
    return match.group(1) if match else ''


def normalize_name(value):
    """Одинаковый регистр ФИО, включая двойные фамилии и инициалы."""
    return clean(value).title()


def normalize_label(value):
    """Обычный регистр названий с сохранением известных сокращений."""
    acronyms = {'ит', 'it', 'ict', 'эвм', 'cad', 'cam', 'чпу'}
    text = clean(value).lower()
    text = text[:1].upper() + text[1:]
    return re.sub(r'[^\W\d_]+',
                  lambda m: m.group().upper() if m.group().casefold() in acronyms else m.group(),
                  text)


def short_name(value):
    parts = normalize_name(value).split()
    return parts[0] + (' ' + ''.join(p[0] + '.' for p in parts[1:]) if len(parts) > 1 else '') if parts else ''


def person(name, phone):
    name, phone = normalize_name(name), clean(phone)
    return f'{name} ({phone})' if phone else name


def context(row):
    return dict(student_name=normalize_name(row['ФИО Студента']).upper(), course=clean(row['Курс']),
                group=clean(row['Группа']).lower(), group_numbers=specialty_code(row['Группа']),
                spec=normalize_label(row['Специальность']), faculty=normalize_label(row['Факультет']),
                curator_info=person(row['ФИО Куратора'], row['Телефон Куратора']),
                zam_dekan_info=person(row['ФИО Зам. декана'], row['Телефон Зам. декана']),
                dekan_info=person(row['ФИО Декана'], row['Телефон Декана']),
                dekan_short=short_name(row['ФИО Декана']),
                student_short=short_name(row['ФИО Студента']),
                curator_short=short_name(row['ФИО Куратора']))


def read_excel(upload):
    # Проверяем размеры распакованного XLSX до загрузки DataFrame.
    with ZipFile(upload) as archive:
        if sum(i.file_size for i in archive.infolist()) > 100 * 1024 * 1024:
            raise ValueError('Распакованный Excel превышает 100 МБ. Разделите исходную таблицу.')
    upload.seek(0)
    df = pd.read_excel(upload, dtype=str, keep_default_na=False, engine='openpyxl', nrows=MAX_ROWS + 1)
    if len(df) > MAX_ROWS:
        raise ValueError(f'Допустимо до {MAX_ROWS:,} строк первого листа. Разделите Excel.')
    column_names = {c.casefold(): c for c in COLUMNS}
    column_names.update({c.casefold(): canonical for c, canonical in ALIASES.items()})
    df.columns = [column_names.get(clean(c).casefold(), clean(c)) for c in df.columns]
    if df.columns.duplicated().any():
        raise ValueError('После переименования обнаружены повторяющиеся колонки.')
    missing = [c for c in COLUMNS if c not in df]
    if missing:
        raise ValueError('Отсутствуют колонки: ' + ', '.join(missing))
    df = df[COLUMNS].map(clean)
    df = df.loc[df.ne('').any(axis=1)]
    if len(df) > MAX_ROWS:
        raise ValueError(f'Допустимо до {MAX_ROWS:,} строк за загрузку. Разделите Excel.')
    if df.empty:
        raise ValueError('В таблице нет студентов.')
    required = [c for c in COLUMNS if not c.startswith('Телефон')]
    bad = df[required].eq('').any(axis=1)
    if bad.any():
        labels = {value: key for key, value in ALIASES.items()}
        details = []
        for i in df.index[bad][:20]:
            fields = ', '.join(labels.get(c, c) for c in required if df.at[i, c] == '')
            details.append(f'строка {int(i) + 2}: {fields}')
        raise ValueError('Заполните обязательные поля — ' + '; '.join(details) + '.')
    if df.map(lambda x: bool(re.search(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', x))).any().any():
        raise ValueError('В Excel обнаружены недопустимые управляющие символы.')
    return df.reset_index(drop=True)


def check_template(template):
    """Сборщик рассчитан на повторение одного статического шаблона."""
    with ZipFile(template) as z:
        root = etree.fromstring(z.read('word/document.xml'))
        if len(root.findall('.//' + W + 'sectPr')) != 1:
            raise ValueError('Шаблон должен содержать ровно один раздел Word.')
        unsupported = ['drawing', 'pict', 'bookmarkStart', 'footnoteReference',
                       'endnoteReference', 'commentRangeStart', 'hyperlink', 'sdt']
        if any(root.find('.//' + W + tag) is not None for tag in unsupported):
            raise ValueError('В шаблоне появились неподдерживаемые объекты. Используйте комплектный template.docx.')
        for name in z.namelist():
            if re.match(r'word/(header|footer)\d+\.xml$', name):
                raise ValueError('Используйте шаблон без колонтитулов.')


def render_student(template_bytes, row, path):
    # У каждого задания свой DocxTemplate, никаких общих изменяемых объектов.
    doc = DocxTemplate(BytesIO(template_bytes))
    doc.render(context(row), jinja_env=Environment(undefined=StrictUndefined), autoescape=True)
    apply_print_layout(doc.docx.element)
    path.write_bytes(etree.tostring(doc.docx.element, encoding='UTF-8', xml_declaration=True))
    return path


def apply_print_layout(root):
    """Reserve room for filled names, specialties and all three signature tables."""
    section = root.find(W + 'body').find(W + 'sectPr')
    size = section.find(W + 'pgSz')
    size.set(W + 'w', '11906')  # A4, portrait
    size.set(W + 'h', '16838')
    margins = section.find(W + 'pgMar')
    for edge, value in [('top', '850'), ('bottom', '850'),
                        ('left', '1021'), ('right', '1021')]:
        margins.set(W + edge, value)
    # Floating signature tables let the following acknowledgement wrap beside
    # them. Keep them in document order so labels stay above their signatures.
    for positioning in root.iter(W + 'tblpPr'):
        positioning.getparent().remove(positioning)
    body = root.find(W + 'body')
    tables = body.findall(W + 'tbl')
    if len(tables) == 3:
        for paragraph in list(body.findall(W + 'p')):
            text = ''.join(paragraph.itertext(W + 't'))
            if text.startswith('Огоҳномаро гирифтам'):
                paragraph.addprevious(tables[0])
            elif text.startswith('Бо мазмуни огоҳинома'):
                paragraph.addprevious(tables[1])
    for paragraph in root.iter(W + 'p'):
        props = paragraph.find(W + 'pPr')
        if props is None:
            props = etree.Element(W + 'pPr')
            paragraph.insert(0, props)
        spacing = props.find(W + 'spacing')
        if spacing is None:
            spacing = etree.SubElement(props, W + 'spacing')
        for key, value in [('before', '0'), ('after', '0'),
                           ('line', '240'), ('lineRule', 'auto'),
                           ('beforeAutospacing', '0'), ('afterAutospacing', '0')]:
            spacing.set(W + key, value)
        # Empty separators must not consume full body-text lines.
        if next(paragraph.iter(W + 't'), None) is None:
            spacing.set(W + 'line', '40')
            spacing.set(W + 'lineRule', 'exact')


def rendering_template(template_bytes):
    """Шрифты нужны Word при печати, но не Jinja при подстановке текста.
    Рабочая копия без их бинарных данных снижает RAM каждого потока.
    Итоговый DOCX всё равно получает ВСЕ ресурсы оригинала.
    """
    buffer = BytesIO()
    with ZipFile(BytesIO(template_bytes)) as source, ZipFile(buffer, 'w', ZIP_DEFLATED) as target:
        for item in source.infolist():
            name = item.filename
            if name.startswith('word/fonts/') or name == 'word/_rels/fontTable.xml.rels':
                continue
            data = source.read(name)
            if name == 'word/fontTable.xml':
                root = etree.fromstring(data)
                for element in list(root.iter()):
                    if etree.QName(element).localname.startswith('embed'):
                        element.getparent().remove(element)
                data = etree.tostring(root)
            elif name == '[Content_Types].xml':
                root = etree.fromstring(data)
                for element in list(root):
                    if element.get('PartName', '').startswith('/word/fonts/'):
                        root.remove(element)
                data = etree.tostring(root)
            target.writestr(item, data)
    return buffer.getvalue()


def generate_part(rows, template_bytes, output, progress=lambda n, total: None):
    """Параллельный рендер + потоковая сборка XML: RAM не растёт с числом страниц.

    Все документы происходят из одного шаблона, вставки только текстовые.
    Поэтому styles, numbering, relationships и встроенные шрифты общие.
    Это НЕ универсальный склеиватель произвольных Word-документов.
    """
    check_template(BytesIO(template_bytes))
    render_bytes = rendering_template(template_bytes)
    count = len(rows)
    if not count:
        raise ValueError('Пустая часть.')
    with TemporaryDirectory(prefix='ogoh-') as temp, ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        pending = deque()
        source = iter(enumerate(rows))
        def submit_one():
            item = next(source, None)
            if item is not None:
                i, row = item
                pending.append(pool.submit(render_student, render_bytes, row, Path(temp) / f'{i}.xml'))
        for _ in range(MAX_WORKERS * 2):
            submit_one()
        try:
            with ZipFile(output, 'w', compression=ZIP_DEFLATED, compresslevel=6) as result:
                # Общие ресурсы берём один раз из неизменного шаблона.
                first = pending[0].result()
                root = etree.parse(str(first)).getroot()
                with ZipFile(BytesIO(template_bytes)) as z:
                    for item in z.infolist():
                        if item.filename != 'word/document.xml':
                            result.writestr(item, z.read(item.filename))
                section = root.find(W + 'body').find(W + 'sectPr')
                with result.open('word/document.xml', 'w', force_zip64=True) as stream:
                    with etree.xmlfile(stream, encoding='UTF-8') as xml:
                        xml.write_declaration(standalone=True)
                        with xml.element(root.tag, attrib=root.attrib, nsmap=root.nsmap):
                            with xml.element(W + 'body'):
                                done = 0
                                while pending:
                                    path = pending.popleft().result()  # порядок исходного Excel
                                    body = etree.parse(str(path)).getroot().find(W + 'body')
                                    if done:
                                        p = body.find(W + 'p')
                                        if p is None:
                                            raise ValueError('Шаблон должен начинаться с абзаца.')
                                        props = p.find(W + 'pPr')
                                        if props is None:
                                            props = etree.Element(W + 'pPr'); p.insert(0, props)
                                        existing = props.find(W + 'pageBreakBefore')
                                        if existing is not None:
                                            props.remove(existing)
                                        etree.SubElement(props, W + 'pageBreakBefore')
                                    for child in body:
                                        if child.tag != W + 'sectPr':
                                            xml.write(child)
                                    path.unlink()
                                    done += 1
                                    progress(done, count)
                                    submit_one()
                                xml.write(section)
        except BaseException:
            for future in pending:
                future.cancel()
            raise


def main():
    st.set_page_config(page_title='Генератор Огоҳинома', page_icon='📄')
    st.title('Макрос')
    template_kind = st.selectbox('Шаблон документа', ['Огоҳинома', 'Талабот'])
    if template_kind == 'Талабот':
        # Streamlit reruns this entrypoint on deploy, but imported modules may
        # stay in sys.modules. Reload the stateless renderer with its template.
        import importlib
        import talabot
        importlib.reload(talabot).main()
        return
    from examples_ui import show_examples
    show_examples('ogohinoma', TEMPLATE)
    st.write('1. Скачайте пустую таблицу с колонками на таджикском языке.')
    excel_template = BASE / 'example.xlsx'
    if excel_template.exists():
        st.download_button('Скачать пустую таблицу Excel', excel_template.read_bytes(),
                           file_name='Ҷадвали_холӣ.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           key='download_excel_template', on_click='ignore')
    else:
        st.warning('Пустая таблица временно недоступна.')
    st.write('2. Заполните таблицу со второй строки и сохраните её. Названия колонок не меняйте.')
    st.caption('Телефоны можно оставить пустыми. Перед вводом телефонов и кодов задайте ячейкам текстовый формат.')
    st.write('3. Загрузите заполненный файл ниже и нажмите «Сгенерировать документы».')
    upload = st.file_uploader('Загрузите заполненную таблицу (.xlsx)', type=['xlsx'])
    size = int(st.number_input('Студентов в одном документе', min_value=1, max_value=1000,
                               value=DEFAULT_BATCH_SIZE, step=100))
    if not TEMPLATE.exists():
        st.error('Рядом с app.py отсутствует template.docx.'); return
    if upload is None:
        return
    if upload.size > MAX_UPLOAD_BYTES:
        st.error('Размер Excel должен быть не более 20 МБ.'); return
    digest = hashlib.sha256(upload.getvalue() + TEMPLATE.read_bytes() + Path(__file__).read_bytes()).hexdigest()
    identity = (digest, size)
    if st.session_state.get('identity') != identity:
        st.session_state.update(identity=identity, results=[], next_row=0)
    if 'results' not in st.session_state:
        st.session_state.update(results=[], next_row=0)
    try:
        df = read_excel(BytesIO(upload.getvalue()))
    except Exception as error:
        st.error(f'Не удалось прочитать таблицу: {error}'); return
    st.info(f'Студентов: {len(df)}. Документов: {math.ceil(len(df) / size)}.')
    start = st.session_state.next_row
    run = st.button('Сгенерировать документы', type='primary', disabled=start > 0)
    if start > 0:
        if st.button('Начать заново (удалить готовые результаты)'):
            st.session_state.update(results=[], next_row=0); st.rerun()
    for item in st.session_state.results:
        download(item)
    if 0 < start < len(df):
        st.warning('Скачайте готовые части перед продолжением: кнопка ниже удалит их из сеанса.')
        if st.button('Части скачаны — продолжить'):
            st.session_state.results = []
            run = True
    if run:
        bar = st.progress(start / len(df))
        status = st.empty()
        template_bytes = TEMPLATE.read_bytes()
        try:
            for offset in range(start, len(df), size):
                end = min(offset + size, len(df))
                part = offset // size + 1
                def update(done, total):
                    if done == total or done % 10 == 0:
                        bar.progress((offset + done) / len(df))
                        status.text(f'Часть {part}: готово {done} из {total}. Всего {offset + done} из {len(df)}.')
                with TemporaryDirectory(prefix='ogoh-part-') as temp:
                    path = Path(temp) / f'Огоҳинома_Часть_{part:03d}.docx'
                    generate_part(df.iloc[offset:end].to_dict('records'), template_bytes, path, update)
                    if path.stat().st_size > MAX_READY_BYTES:
                        raise ValueError('Часть превысила 64 МБ. Уменьшите количество студентов в документе.')
                    item = dict(name=path.name, data=path.read_bytes(), first=offset + 1, last=end)
                st.session_state.results.append(item)
                st.session_state.next_row = end
                download(item)
                # Максимум 64 МБ сохранённых результатов плюс одна часть (до 64 МБ).
                if sum(len(x['data']) for x in st.session_state.results) >= MAX_READY_BYTES:
                    break
            if st.session_state.next_row == len(df):
                status.success('Все документы готовы. Скачайте части для печати.')
            else:
                st.warning('Достигнут лимит хранения. Скачайте части, затем нажмите кнопку ниже.')
                if st.button('Показать продолжение'):
                    st.rerun()
        except Exception as error:
            st.error(f'Генерация остановлена: {error}. Готовые части сохранены. '
                     'При повторе незавершённая часть создаётся заново.')
            if st.button('Показать восстановление'):
                st.rerun()


def download(item):
    st.download_button(f"Скачать {item['name']} · студенты {item['first']}–{item['last']}",
                       item['data'], file_name=item['name'],
                       mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                       key=f"download-{item['first']}", on_click='ignore')


if __name__ == '__main__':
    main()

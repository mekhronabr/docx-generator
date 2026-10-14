"""Fill the supplied Talabot form without rebuilding its layout or ZIP parts."""
from copy import deepcopy
from datetime import date
from io import BytesIO
from pathlib import Path
import hashlib
import re
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree
import pandas as pd
import streamlit as st

TEMPLATE = Path(__file__).with_name('template_talabot.docx')
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def lines(text):
    return [line.strip() for line in text.splitlines() if line.strip()]


def set_text(paragraph, text, font_size=None):
    """Reuse existing run formatting, including formatting on empty paragraphs."""
    nodes = list(paragraph.iter(W + 't'))
    if nodes:
        nodes[0].text = text
        for node in nodes[1:]:
            node.text = ''
    else:
        run = etree.SubElement(paragraph, W + 'r')
        props = paragraph.find(W + 'pPr/' + W + 'rPr')
        if props is not None:
            run.append(deepcopy(props))
        nodes = [etree.SubElement(run, W + 't')]
        nodes[0].text = text
    nodes[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    if font_size is not None:
        run = nodes[0].getparent()
        props = run.find(W + 'rPr')
        if props is None:
            props = etree.Element(W + 'rPr')
            run.insert(0, props)
        for tag in ('sz', 'szCs'):
            size = props.find(W + tag)
            if size is None:
                size = etree.SubElement(props, W + tag)
            size.set(W + 'val', str(font_size * 2))


def generate(data, template_bytes=None):
    students, absent = data['students'], data.get('absent', [])
    if not 1 <= len(students) <= 40:
        raise ValueError('В основном списке должно быть от 1 до 40 студентов.')
    if len(absent) > 10:
        raise ValueError('В списке отсутствующих допустимо до 10 студентов.')
    for field in ('form_code', 'group', 'date', 'curator', 'deputy'):
        if not data.get(field, '').strip():
            raise ValueError('Заполните код факультета, группу, дату и ФИО ответственных.')
    values = students + absent + [data[k] for k in ('form_code', 'group', 'date', 'curator', 'deputy')]
    if any(re.search(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', value) for value in values):
        raise ValueError('В тексте обнаружены недопустимые управляющие символы.')
    output = BytesIO()
    with ZipFile(BytesIO(template_bytes or TEMPLATE.read_bytes())) as source:
        root = etree.fromstring(source.read('word/document.xml'))
        paragraphs = list(root.iter(W + 'p'))
        set_text(paragraphs[0], data['form_code'])
        tables = list(root.iter(W + 'tbl'))
        align_main_table(tables[2])
        match_numbering(tables[2], tables[3])
        for table in (tables[2], tables[3]):
            equalize_student_rows(table)
        for table, offset in ((tables[2], 0), (tables[3], 20), (tables[5], 0), (tables[6], 5)):
            names = students if len(table.findall(W + 'tr')) == 21 else absent
            for index, row in enumerate(table.findall(W + 'tr')[1:], offset):
                cell = row.findall(W + 'tc')[1]
                set_text(cell.find(W + 'p'), names[index] if index < len(names) else '', font_size=10)
        for paragraph in paragraphs:
            text = ''.join(n.text or '' for n in paragraph.iter(W + 't'))
            if 'Ба ҳайъати асосии' in text:
                replace_text(paragraph, '_______________', data['group'])
                replace_text(paragraph, '_______', data['date'])
        for row, name in zip(tables[7].findall(W + 'tr'), (data['curator'], data['deputy'])):
            set_text(row.findall(W + 'tc')[2].find(W + 'p'), name, font_size=10)
        with ZipFile(output, 'w', ZIP_DEFLATED) as target:
            for item in source.infolist():
                contents = etree.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True) if item.filename == 'word/document.xml' else source.read(item.filename)
                target.writestr(item, contents)
    return output.getvalue()


def match_numbering(left, right):
    """Use the right table's explicit number formatting for numbers 1–20."""
    reference = right.findall(W + 'tr')[1].findall(W + 'tc')[0].find(W + 'p')
    for number, row in enumerate(left.findall(W + 'tr')[1:], 1):
        cell = row.findall(W + 'tc')[0]
        original = cell.find(W + 'p')
        paragraph = deepcopy(reference)
        # The left number column is narrower; the right column's indent would wrap.
        indent = paragraph.find(W + 'pPr/' + W + 'ind')
        if indent is not None:
            indent.set(W + 'left', '0')
        for attribute in list(paragraph.attrib):
            if etree.QName(attribute).localname in ('paraId', 'textId'):
                del paragraph.attrib[attribute]
        set_text(paragraph, f'{number}.')
        cell.replace(original, paragraph)


def equalize_student_rows(table):
    """Keep empty and filled student rows equally tall, with room for two lines."""
    for row in table.findall(W + 'tr')[1:]:
        props = row.find(W + 'trPr')
        if props is None:
            props = etree.Element(W + 'trPr')
            row.insert(0, props)
        height = props.find(W + 'trHeight')
        if height is None:
            height = etree.SubElement(props, W + 'trHeight')
        height.set(W + 'val', '560')
        height.set(W + 'hRule', 'atLeast')


def align_main_table(table):
    """Use one grid for the header and body; the source header skips a grid slot."""
    widths = (704, 2961, 1505)
    grid = table.find(W + 'tblGrid')
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        etree.SubElement(grid, W + 'gridCol').set(W + 'w', str(width))
    props = table.find(W + 'tblPr')
    props.find(W + 'tblW').set(W + 'w', str(sum(widths)))
    layout = props.find(W + 'tblLayout')
    if layout is None:
        layout = etree.SubElement(props, W + 'tblLayout')
    layout.set(W + 'type', 'fixed')
    for row in table.findall(W + 'tr'):
        row_props = row.find(W + 'trPr')
        if row_props is not None:
            for tag in ('gridBefore', 'gridAfter', 'wBefore', 'wAfter'):
                node = row_props.find(W + tag)
                if node is not None:
                    row_props.remove(node)
        for cell, width in zip(row.findall(W + 'tc'), widths):
            cell_props = cell.find(W + 'tcPr')
            span = cell_props.find(W + 'gridSpan')
            if span is not None:
                cell_props.remove(span)
            cell_props.find(W + 'tcW').set(W + 'w', str(width))


def replace_text(paragraph, old, new):
    """Replace a slot even when Word has split it across multiple text runs."""
    nodes = list(paragraph.iter(W + 't'))
    text = ''.join(n.text or '' for n in nodes)
    start = text.index(old)
    end = start + len(old)
    offset = 0
    for node in nodes:
        value = node.text or ''
        stop = offset + len(value)
        if offset < end and stop > start:
            node.text = value[:max(0, start - offset)] + (new if offset <= start < stop else '') + value[max(0, end - offset):]
            node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        offset = stop


def main():
    st.write('Лист ознакомления группы с требованиями. ФИО вводятся по одному на строку или загружаются из Excel.')
    if not TEMPLATE.exists():
        st.error('Отсутствует template_talabot.docx рядом с app.py.')
        return
    excel_example = BytesIO()
    pd.DataFrame(columns=['Ному насаби донишҷӯ', 'Гурӯҳ', 'Код факультета', 'Дата', 'ФИО Куратора', 'ФИО Заместителя', 'Отсутствует']).to_excel(excel_example, index=False)
    st.download_button('Скачать пустую таблицу Excel для Талабот', excel_example.getvalue(), file_name='Талабот_ҷадвали_холӣ.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', on_click='ignore')
    mode = st.radio('Режим генерации', ['Одна группа', 'Несколько групп'], horizontal=True)
    if mode == 'Несколько групп':
        batch_main()
        return
    upload = st.file_uploader('Excel со списком студентов (.xlsx)', type=['xlsx'], key='talabot_excel')
    imported = ''
    if upload is not None:
        try:
            if upload.size > 20 * 1024 * 1024:
                raise ValueError('Размер Excel должен быть не более 20 МБ.')
            with ZipFile(BytesIO(upload.getvalue())) as archive:
                if sum(item.file_size for item in archive.infolist()) > 100 * 1024 * 1024:
                    raise ValueError('Распакованный Excel превышает 100 МБ.')
            df = pd.read_excel(BytesIO(upload.getvalue()), dtype=str, keep_default_na=False, nrows=100001)
            if len(df) > 100000:
                raise ValueError('Допустимо до 100 000 строк. Разделите Excel.')
            if not len(df.columns):
                raise ValueError('В Excel нет колонок.')
            columns = list(df.columns)
            preferred = next((i for i, c in enumerate(columns) if str(c).strip().casefold() in ('фио студента', 'ному насаби донишҷӯ', 'фио', 'ному насаб')), 0)
            column = st.selectbox('Колонка с ФИО студентов', columns, index=preferred, key='talabot_name_column')
            group_columns = [c for c in df.columns if str(c).strip().casefold() in ('группа', 'гурӯҳ', 'гурух')]
            if group_columns:
                group_column = group_columns[0]
                selected = st.selectbox('Группа из Excel', list(dict.fromkeys(df[group_column])), key='talabot_excel_group')
                df = df.loc[df[group_column] == selected]
            imported = '\n'.join(str(v).strip() for v in df[column] if str(v).strip())
        except Exception as error:
            st.error(f'Не удалось прочитать Excel: {error}')
            return
    with st.form('talabot_form'):
        code = st.text_input('Код факультета', value='ФЭ_11')
        group = st.text_input('Группа', value='1_53.01.05тб')
        when = st.date_input('Дата ознакомления', value=date(2025, 10, 10))
        curator = st.text_input('Фамилия и имя — куратори гурӯҳи академӣ')
        deputy = st.text_input('Фамилия и имя — муовини директор оид ба тарбия')
        students = st.text_area('Фамилии и имена студентов (до 40)', value=imported, height=300)
        absent = st.text_area('Отсутствующие (до 10, по одному на строку)', height=120)
        submitted = st.form_submit_button('Сгенерировать документ', type='primary')
    if submitted:
        st.session_state.pop('talabot_result', None)
        try:
            data = dict(form_code=code.strip(), group=group.strip(), date=when.strftime('%d.%m.%Y'), curator=curator.strip(), deputy=deputy.strip(), students=lines(students), absent=lines(absent))
            st.session_state.talabot_result = generate(data)
        except Exception as error:
            st.error(f'Не удалось создать документ: {error}')
    if st.session_state.get('talabot_result'):
        st.download_button('Скачать заполненный Талабот', st.session_state.talabot_result, file_name='Талабот.docx', mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document', on_click='ignore')


BATCH_FIELDS = ['Группа', 'Код факультета', 'Дата', 'ФИО Куратора', 'ФИО Заместителя']


def read_batch_excel(contents):
    if len(contents) > 20 * 1024 * 1024:
        raise ValueError('Размер Excel должен быть не более 20 МБ.')
    with ZipFile(BytesIO(contents)) as archive:
        if sum(item.file_size for item in archive.infolist()) > 100 * 1024 * 1024:
            raise ValueError('Распакованный Excel превышает 100 МБ.')
    df = pd.read_excel(BytesIO(contents), dtype=str, keep_default_na=False, nrows=100001)
    if len(df) > 100000:
        raise ValueError('Допустимо до 100 000 строк. Разделите Excel.')
    aliases = {
        'ному насаби донишҷӯ': 'ФИО Студента', 'фио': 'ФИО Студента',
        'ному насаб': 'ФИО Студента', 'гурӯҳ': 'Группа', 'гурух': 'Группа',
        'куратори гурӯҳ': 'ФИО Куратора',
        'фио зам. декана': 'ФИО Заместителя',
        'муовини декан оид ба тарбия': 'ФИО Заместителя',
        'муовини директор оид ба тарбия': 'ФИО Заместителя',
    }
    aliases.update({c.casefold(): c for c in ['ФИО Студента', *BATCH_FIELDS, 'Отсутствует']})
    df.columns = [aliases.get(str(c).strip().casefold(), str(c).strip()) for c in df.columns]
    if df.columns.duplicated().any():
        raise ValueError('В Excel повторяются названия колонок.')
    for column in ('ФИО Студента', 'Группа'):
        if column not in df:
            raise ValueError(f'В Excel отсутствует колонка «{column}».')
    df = df.map(lambda value: str(value).strip())
    df = df.loc[df.ne('').any(axis=1)].copy()
    if df.empty:
        raise ValueError('В Excel нет студентов.')
    bad = df[['ФИО Студента', 'Группа']].eq('').any(axis=1)
    if bad.any():
        raise ValueError('Заполните ФИО и группу в строках: ' + ', '.join(str(i + 2) for i in df.index[bad][:20]))
    if 'Отсутствует' in df:
        allowed = {'', 'да', 'нет', '1', '0', 'yes', 'no', 'true', 'false', '+', '-'}
        if not df['Отсутствует'].str.casefold().isin(allowed).all():
            raise ValueError('В колонке «Отсутствует» укажите «да», «нет» или оставьте пусто.')
    return df


def batch_settings(df):
    settings = []
    for group, rows in df.groupby('Группа', sort=False):
        item = {'Группа': group}
        for field in BATCH_FIELDS[1:]:
            values = list(dict.fromkeys(value for value in rows[field] if value)) if field in rows else []
            if len(values) > 1:
                raise ValueError(f'Группа {group}: разные значения в колонке «{field}». Укажите одно значение для группы.')
            value = values[0] if values else ''
            if field == 'Дата' and value:
                value = pd.to_datetime(value, dayfirst=True, errors='raise').strftime('%d.%m.%Y')
            item[field] = value
        settings.append(item)
    return pd.DataFrame(settings, columns=BATCH_FIELDS)


def generate_batch(df, settings, progress=lambda done, total: None):
    """One separate DOCX per group, numbered ZIP entries avoid name collisions."""
    if len(settings) != len(df['Группа'].unique()) or settings['Группа'].duplicated().any() or set(settings['Группа']) != set(df['Группа']):
        raise ValueError('Настройки должны содержать каждую группу ровно один раз.')
    archive = BytesIO()
    template_bytes = TEMPLATE.read_bytes()
    with ZipFile(archive, 'w', ZIP_DEFLATED) as output:
        for index, item in enumerate(settings.to_dict('records'), 1):
            group = item['Группа']
            rows = df.loc[df['Группа'] == group]
            missing = [field for field in BATCH_FIELDS if pd.isna(item[field]) or not str(item[field]).strip()]
            if missing:
                raise ValueError(f'Группа {group}: заполните ' + ', '.join(missing))
            when = pd.to_datetime(str(item['Дата']), dayfirst=True, errors='raise').strftime('%d.%m.%Y')
            absent = rows.loc[rows['Отсутствует'].str.casefold().isin({'да', '1', 'yes', 'true', '+'}), 'ФИО Студента'].tolist() if 'Отсутствует' in rows else []
            data = dict(form_code=str(item['Код факультета']).strip(), group=group, date=when, curator=str(item['ФИО Куратора']).strip(), deputy=str(item['ФИО Заместителя']).strip(), students=rows['ФИО Студента'].tolist(), absent=absent)
            try:
                document = generate(data, template_bytes)
            except ValueError as error:
                raise ValueError(f'Группа {group}: {error}') from error
            name = re.sub(r'[^\w.-]', '_', group)[:80]
            output.writestr(f'Талабот_{index:03d}_{name}.docx', document)
            if archive.tell() > 64 * 1024 * 1024:
                raise ValueError('Архив превышает 64 МБ. Разделите Excel на несколько файлов.')
            progress(index, len(settings))
    return archive.getvalue()


def batch_main():
    st.write('Загрузите Excel со всеми группами. Одна строка — один студент; колонка группы обязательна.')
    st.caption('Данные ответственных можно заполнить в Excel или в настройках ниже. «Отсутствует»: да/нет; отсутствующий остаётся и в основном списке.')
    upload = st.file_uploader('Excel для нескольких групп (.xlsx)', type=['xlsx'], key='talabot_batch_excel')
    if upload is None:
        return
    try:
        df = read_batch_excel(upload.getvalue())
        settings = batch_settings(df)
    except Exception as error:
        st.error(f'Не удалось прочитать Excel: {error}')
        return
    selected = st.multiselect('Группы для генерации', settings['Группа'].tolist(), default=settings['Группа'].tolist())
    if not selected:
        st.info('Выберите хотя бы одну группу.')
        return
    df = df.loc[df['Группа'].isin(selected)]
    settings = settings.loc[settings['Группа'].isin(selected)]
    st.info(f'Студентов: {len(df)}. Документов: {len(settings)}.')
    with st.form('talabot_batch_form'):
        st.write('Настройки каждого документа')
        edited = st.data_editor(settings, hide_index=True, disabled=['Группа'], num_rows='fixed', key='talabot_batch_settings')
        submitted = st.form_submit_button('Сгенерировать все документы', type='primary')
    identity = hashlib.sha256(upload.getvalue() + edited.to_csv(index=False).encode()).hexdigest()
    if st.session_state.get('talabot_batch_identity') != identity:
        st.session_state.pop('talabot_batch_result', None)
    if submitted:
        st.session_state.pop('talabot_batch_result', None)
        bar = st.progress(0.0)
        try:
            st.session_state.talabot_batch_result = generate_batch(df, edited, lambda done, total: bar.progress(done / total))
            st.session_state.talabot_batch_identity = identity
            st.success(f'Готово документов: {len(edited)}.')
        except Exception as error:
            st.error(f'Не удалось создать документы: {error}')
    if st.session_state.get('talabot_batch_result'):
        st.download_button('Скачать все документы ZIP', st.session_state.talabot_batch_result, file_name='Талабот_все_группы.zip', mime='application/zip', on_click='ignore')

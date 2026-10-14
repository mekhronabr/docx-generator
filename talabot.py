"""Fill the supplied Talabot form without rebuilding its layout or ZIP parts."""
from copy import deepcopy
from datetime import date
from io import BytesIO
from pathlib import Path
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
            raise ValueError('Заполните код бланка, группу, дату и ФИО ответственных.')
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
    pd.DataFrame(columns=['Ному насаби донишҷӯ', 'Гурӯҳ']).to_excel(excel_example, index=False)
    st.download_button('Скачать пустую таблицу Excel для Талабот', excel_example.getvalue(), file_name='Талабот_ҷадвали_холӣ.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', on_click='ignore')
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
        code = st.text_input('Код бланка', value='ФЭ_11')
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
        st.download_button('Скачать заполненный Талабот', st.session_state.talabot_result, file_name='Талабот_саҳифаи_2.docx', mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document', on_click='ignore')

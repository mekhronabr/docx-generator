"""Adapt the supplied Word reference in place, preserving its package and layout.
Run with bundled document Python: python qa/build_talabot_template.py reference.docx
"""
from copy import deepcopy
from pathlib import Path
import sys
from zipfile import ZipFile
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
BASE = Path(__file__).resolve().parents[1]


def wrap_run(parent, run, tag):
    control = etree.Element(W + 'sdt')
    props = etree.SubElement(control, W + 'sdtPr')
    etree.SubElement(props, W + 'tag').set(W + 'val', tag)
    content = etree.SubElement(control, W + 'sdtContent')
    parent.replace(run, control)
    content.append(run)


def wrap_paragraph(paragraph, tag, clear=False):
    runs = paragraph.findall(W + 'r')
    run = deepcopy(runs[0]) if runs else etree.Element(W + 'r')
    for child in list(run):
        if child.tag != W + 'rPr': run.remove(child)
    node = etree.SubElement(run, W + 't')
    node.text = '' if clear else ''.join(paragraph.itertext())
    for child in list(paragraph):
        if child.tag != W + 'pPr': paragraph.remove(child)
    paragraph.append(run)
    wrap_run(paragraph, run, tag)


def build(reference):
    with ZipFile(reference) as source:
        root = etree.fromstring(source.read('word/document.xml'))
        tables = list(root.iter(W + 'tbl'))
        wrap_paragraph(tables[0].find(W + 'tr').find(W + 'tc').find(W + 'p'), 'form_code')
        for table_index, prefix, offset in [(1, 'student', 0), (2, 'student', 20), (4, 'absent', 0), (5, 'absent', 5)]:
            table = tables[table_index]
            header = table.findall(W + 'tr')[0].findall(W + 'tc')[1].find(W + 'p')
            nodes = list(header.iter(W + 't'))
            nodes[0].text = 'Ному насаб'
            for node in nodes[1:]: node.text = ''
            for number, row in enumerate(table.findall(W + 'tr')[1:], offset + 1):
                if prefix == 'student' and offset == 0:
                    paragraph = row.findall(W + 'tc')[0].find(W + 'p')
                    props = paragraph.find(W + 'pPr')
                    for tag in ('numPr',):
                        node = props.find(W + tag) if props is not None else None
                        if node is not None: props.remove(node)
                    if props is not None:
                        etree.SubElement(etree.SubElement(props, W + 'numPr'), W + 'numId').set(W + 'val', '0')
                    run = etree.SubElement(paragraph, W + 'r')
                    font = props.find(W + 'rPr') if props is not None else None
                    if font is not None: run.append(deepcopy(font))
                    etree.SubElement(run, W + 't').text = f'{number}.'
                if prefix == 'student':
                    props = row.find(W + 'trPr')
                    if props is None:
                        props = etree.Element(W + 'trPr'); row.insert(0, props)
                    height = props.find(W + 'trHeight')
                    if height is None: height = etree.SubElement(props, W + 'trHeight')
                    height.set(W + 'val', '450')
                    height.set(W + 'hRule', 'atLeast')
                cell = row.findall(W + 'tc')[1]
                wrap_paragraph(cell.find(W + 'p'), f'{prefix}_{number}', True)
                for paragraph in cell.findall(W + 'p')[1:]:
                    # Empty extra paragraphs are source page furniture; preserve them.
                    if ''.join(paragraph.itertext()):
                        for node in paragraph.iter(W + 't'): node.text = ''
        for paragraph in root.iter(W + 'p'):
            if 'Ба ҳайъати' in ''.join(paragraph.itertext()):
                for run in list(paragraph.findall(W + 'r')):
                    text = ''.join(run.itertext())
                    if text in ('______________', '________'):
                        tag = 'group' if len(text) == 14 else 'date'
                        props = run.find(W + 'rPr')
                        if props is None:
                            props = etree.Element(W + 'rPr'); run.insert(0, props)
                        etree.SubElement(props, W + 'u').set(W + 'val', 'single')
                        if tag == 'group':
                            spacer = etree.Element(W + 'r')
                            etree.SubElement(spacer, W + 't').text = ' '
                            spacer[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                            paragraph.insert(paragraph.index(run), spacer)
                        else:
                            spacer = etree.Element(W + 'r')
                            etree.SubElement(spacer, W + 't').text = ' '
                            spacer[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                            paragraph.insert(paragraph.index(run), spacer)
                        wrap_run(paragraph, run, tag)
        for row, tag in zip(tables[3].findall(W + 'tr')[1:], ('curator', 'deputy')):
            paragraph = row.findall(W + 'tc')[2].findall(W + 'p')[-1]
            wrap_paragraph(paragraph, tag)
        with ZipFile(BASE / 'template_talabot.docx', 'w') as target:
            for item in source.infolist():
                data = etree.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True) if item.filename == 'word/document.xml' else source.read(item.filename)
                target.writestr(item, data)

if __name__ == '__main__':
    build(Path(sys.argv[1]))

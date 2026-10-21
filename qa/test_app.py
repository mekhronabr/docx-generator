import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
from lxml import etree
from tempfile import TemporaryDirectory


def run():
    assert app.short_name('Иванов Иван Иванович') == 'Иванов И.И.'
    assert app.short_name('Иванов Иван') == 'Иванов И.'
    assert app.normalize_name('  шАМОЛОВ   тУПАЛАНГ ') == 'Шамолов Тупаланг'
    assert app.normalize_name('ИВАНОВ-ПЕТРОВ ИВАН') == 'Иванов-Петров Иван'
    assert app.short_name('дЕКАНОВ фАКУЛТЕТ') == 'Деканов Ф.'
    assert app.normalize_label('фАКУЛЬТЕТИ ИНЖЕНЕРИЯ') == 'Факультети инженерия'
    assert app.normalize_label('САМОЛЕТОСТРОЕНИЕ') == 'Самолетостроение'
    assert app.normalize_label('ПРОГРАММИРОВАНИЕ ЭВМ') == 'Программирование ЭВМ'
    for empty in [None, float('nan'), '', '  ', pd.NA]:
        assert app.person('Иванов Иван', empty) == 'Иванов Иван'
    assert app.person('Иванов Иван', '123') == 'Иванов Иван (123)'
    record = dict(zip(app.COLUMNS, [
        '  аКРАМчонЗОДА   фИРДАВСЧОН ', '1', '1.53.01.05ТБ',
        'ҳАРАКАТДИҲАНДАҲОИ АВТОМАТИИ БАРҚӢ', 'фАКУЛТЕТИ ЭНЕРГЕТИКА',
        'нУРИДДИНОВА зАРИНА дАВРОНОВНА', '92-884-00-85',
        'рАҲМАТУЛЛОЕВА мАРҲАБО мАҲМУДОВНА', '92-717-10-55',
        'мӮМИНОВА шОҲЗОДАХОН нАЗИП', '92-915-10-02']))
    stream = BytesIO()
    pd.DataFrame([record]).rename(columns=str.swapcase).to_excel(stream, index=False)
    stream.seek(0)
    df = app.read_excel(stream)
    assert len(df) == 1
    assert app.context(df.iloc[0])['group_numbers'] == '1.53.01.05'
    ctx = app.context(df.iloc[0])
    assert ctx['student_name'] == 'АКРАМЧОНЗОДА ФИРДАВСЧОН'
    assert ctx['group'] == '1.53.01.05тб'
    assert ctx['faculty'] == 'Факултети энергетика'
    assert ctx['curator_info'] == 'Нуриддинова Зарина Давроновна (92-884-00-85)'
    assert ctx['dekan_short'] == 'Мӯминова Ш.Н.'
    for group, expected in [('344', '344'), ('344аб', '344'), ('1.53.01.05тб', '1.53.01.05'), (' 01.02тб ', '01.02'), ('12AB', '12')]:
        assert app.specialty_code(group) == expected
    rows = df.to_dict('records') * 3
    rows = [dict(x) for x in rows]
    for i, row in enumerate(rows):
        row['ФИО Студента'] = f'Тестовый{i+1} Иван Иванович'
    rows[1]['Телефон Куратора'] = ''
    rows[2]['Специальность'] = 'Энергетика & <автоматика>'
    path = Path('/private/tmp/ogoh-test.docx')
    app.generate_part(rows, app.TEMPLATE.read_bytes(), path)
    with ZipFile(path) as z:
        root = etree.fromstring(z.read('word/document.xml'))
        text = ''.join(root.itertext())
        assert '{{' not in text and '()' not in text
        assert text.index('ТЕСТОВЫЙ1 ИВАН') < text.index('ТЕСТОВЫЙ2 ИВАН') < text.index('ТЕСТОВЫЙ3 ИВАН')
        assert 'Энергетика & <автоматика>' in text
        assert len(root.findall('.//' + app.W + 'sectPr')) == 1
        assert len(root.findall('.//' + app.W + 'pageBreakBefore')) == 2
        size = root.find('.//' + app.W + 'pgSz')
        assert (size.get(app.W + 'w'), size.get(app.W + 'h')) == ('11906', '16838')
        assert not root.findall('.//' + app.W + 'tblpPr')
        assert len(root.findall('.//' + app.W + 'tbl')) == 9
        assert z.testzip() is None
        for r in root.iter(app.W + 'r'):
            value = ''.join(r.itertext(app.W + 't'))
            if value in ['1', '1.53.01.05тб', ctx['curator_info']]:
                props = r.find(app.W + 'rPr')
                assert props.find(app.W + 'b') is not None
                assert props.find(app.W + 'i').get(app.W + 'val') == '1'
                assert props.find(app.W + 'u').get(app.W + 'val') == 'single'
    # Ошибки обязательных полей должны обнаруживаться до генерации.
    df.loc[0,'ФИО Студента'] = ''
    stream = BytesIO(); df.to_excel(stream, index=False); stream.seek(0)
    try:
        app.read_excel(stream)
    except ValueError as error:
        assert 'строка 2: ному насаби донишҷӯ' in str(error)
    else:
        raise AssertionError('Missing name accepted')
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(app.BASE / 'app.py')).run(timeout=30)
    assert not at.exception
    print('PASS: formatting, original Excel, validation, merge order, page breaks, XML escaping, UI startup')

if __name__ == '__main__':
    run()

"""Run with: venv/bin/python -m unittest discover -s qa -p 'test_talabot_batch.py'."""
from io import BytesIO
import unittest
from zipfile import ZipFile

import pandas as pd
from lxml import etree
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

from talabot import W, batch_settings, display_date, generate_batch, read_batch_excel


class BatchTests(unittest.TestCase):
    def source(self):
        rows = [
            ['Алиев Али', 'Группа/1', 'ФЭ_11', '10.10.2025', 'Куратор Один', 'Заместитель Один', 'да'],
            ['Петров Пётр', 'Группа/1', '', '', '', '', 'нет'],
            ['Иванов Иван & Иванович', 'Группа:1', 'ФЭ_12', '11.10.2025', 'Куратор Два', 'Заместитель Два', ''],
        ]
        df = pd.DataFrame(rows, columns=['Ному насаби донишҷӯ', 'Гурӯҳ', 'Код факультета', 'Дата', 'ФИО Куратора', 'ФИО Заместителя', 'Отсутствует'])
        contents = BytesIO()
        df.to_excel(contents, index=False)
        return read_batch_excel(contents.getvalue())

    def test_distinct_documents_and_absences(self):
        df = self.source()
        with ZipFile(BytesIO(generate_batch(df, batch_settings(df)))) as document:
            root = etree.fromstring(document.read('word/document.xml'))
            pages = [[]]
            for child in root.find(W + 'body'):
                if child.find('.//' + W + 'br[@' + W + 'type="page"]') is not None:
                    pages.append([])
                else:
                    pages[-1].append(''.join(child.itertext()))
            self.assertEqual(len(pages), 2)
            texts = [''.join(page) for page in pages]
            self.assertEqual(len(list(root.iter(W + 'sectPr'))), 1)
            self.assertEqual(len(list(root.iter(W + 'sdt'))), 0)
            self.assertNotIn('Ному насаб ва номи падар', ''.join(texts))
            for table in root.iter(W + 'tbl'):
                if len(table.findall(W + 'tr')) == 21:
                    nums = [''.join(row.findall(W + 'tc')[0].itertext()) for row in table.findall(W + 'tr')[1:]]
                    self.assertIn(nums, [[f'{i}.' for i in range(1, 21)], [f'{i}.' for i in range(21, 41)]])
            self.assertIn('Куратор Один', texts[0])
            self.assertNotIn('Куратор Два', texts[0])
            self.assertEqual(texts[0].count('Алиев Али'), 2)
            self.assertIn('ФЭ_12', texts[1])
            self.assertIn('11.10.2025', texts[1])
            self.assertIn('Иванов Иван & Иванович', texts[1])
            self.assertNotIn('Петров Пётр', texts[1])

    def test_rich_text_case_and_underline(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(['ФИО Студента', 'Группа', 'Код факультета', 'Дата', 'ФИО Куратора', 'ФИО Заместителя'])
        sheet.append(['Фамилия Имя', 'ГР_01', 'ФЭ_11', '10.10.2025', 'Куратор Один', 'Заместитель Один'])
        sheet['A2'] = CellRichText([TextBlock(InlineFont(b=True, u='single'), 'ФАМИЛИЯ'), ' имя_с_подчеркиванием'])
        sheet['B2'].font = Font(b=True, i=True, u='double')
        source = BytesIO()
        workbook.save(source)
        df = read_batch_excel(source.getvalue())
        with ZipFile(BytesIO(generate_batch(df, batch_settings(df)))) as document:
            root = etree.fromstring(document.read('word/document.xml'))
            runs = [r for r in root.iter(W + 'r') if ''.join(r.itertext()) == 'ФАМИЛИЯ']
            self.assertEqual(len(runs), 1)
            self.assertEqual(runs[0].find(W + 'rPr/' + W + 'b').get(W + 'val'), '1')
            self.assertEqual(runs[0].find(W + 'rPr/' + W + 'u').get(W + 'val'), 'single')
            self.assertIn('имя_с_подчеркиванием', ''.join(root.itertext()))
            group_run = next(r for r in root.iter(W + 'r') if ''.join(r.itertext()) == 'ГР_01')
            self.assertEqual(group_run.find(W + 'rPr/' + W + 'u').get(W + 'val'), 'double')

    def test_conflicting_group_metadata(self):
        df = self.source()
        df.loc[1, 'ФИО Куратора'] = 'Другой куратор'
        with self.assertRaisesRegex(ValueError, 'разные значения'):
            batch_settings(df)

    def test_capacity_and_required_fields(self):
        df = self.source().iloc[[0]].copy()
        settings = batch_settings(df)
        with self.assertRaisesRegex(ValueError, '40 студентов'):
            generate_batch(pd.concat([df] * 41, ignore_index=True), settings)
        settings.loc[0, 'Код факультета'] = ''
        with self.assertRaisesRegex(ValueError, 'Код факультета'):
            generate_batch(df, settings)

    def test_excel_iso_dates_keep_day_and_month(self):
        self.assertEqual(display_date('2025-10-11 00:00:00'), '11.10.2025')
        self.assertEqual(display_date('11.10.2025'), '11.10.2025')


if __name__ == '__main__':
    unittest.main()

"""Run with: venv/bin/python -m unittest discover -s qa -p 'test_talabot_batch.py'."""
from io import BytesIO
import unittest
from zipfile import ZipFile

import pandas as pd
from lxml import etree

from talabot import W, batch_settings, generate_batch, read_batch_excel


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
        with ZipFile(BytesIO(generate_batch(df, batch_settings(df)))) as archive:
            self.assertEqual(len(archive.namelist()), 2)
            texts = []
            for name in archive.namelist():
                with ZipFile(BytesIO(archive.read(name))) as document:
                    root = etree.fromstring(document.read('word/document.xml'))
                    texts.append(''.join(root.itertext()))
                    tables = list(root.iter(W + 'tbl'))
                    nums = [''.join(row.findall(W + 'tc')[0].itertext()) for row in tables[2].findall(W + 'tr')[1:]]
                    self.assertEqual(nums, [f'{i}.' for i in range(1, 21)])
            self.assertIn('Куратор Один', texts[0])
            self.assertNotIn('Куратор Два', texts[0])
            self.assertEqual(texts[0].count('Алиев Али'), 2)
            self.assertIn('ФЭ_12', texts[1])
            self.assertIn('11.10.2025', texts[1])
            self.assertIn('Иванов Иван & Иванович', texts[1])
            self.assertNotIn('Петров Пётр', texts[1])

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


if __name__ == '__main__':
    unittest.main()

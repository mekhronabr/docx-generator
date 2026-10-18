from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import pandas as pd

import batch_jobs as jobs


class JobTests(unittest.TestCase):
    def test_resume_and_release_processed_rows(self):
        with TemporaryDirectory() as folder, patch.object(jobs, 'ROOT', Path(folder)):
            df = pd.DataFrame({'Группа': [f'G{i}' for i in range(100)]})
            df.attrs['formats'] = {(i, 'Группа'): [str(i)] for i in range(100)}
            token = jobs.create('source.xlsx', df, df.copy())
            state = jobs.load(token)
            state.update(batch_size=20, running=True)
            jobs.save(token, state)
            calls = []

            def generate(part, settings):
                calls.extend(settings['Группа'])
                self.assertEqual(len(part), 20)
                self.assertEqual(len(part.attrs['formats']), 20)
                return ','.join(settings['Группа']).encode()

            for count in range(1, 6):
                jobs.process_next(token, generate)
                restored = jobs.load(token)
                self.assertEqual(restored['done'], count * 20)
                self.assertEqual(len(restored['pending']), 100 - count * 20)
                self.assertEqual(len(restored['pending'].attrs['formats']), 100 - count * 20)
                self.assertEqual(len(restored['results']), count)
                self.assertTrue(all((jobs.directory(token) / name).is_file() for name in restored['results']))
            self.assertEqual(calls, df['Группа'].tolist())
            self.assertFalse(restored['running'])

    def test_failure_keeps_queue_and_completed_result(self):
        with TemporaryDirectory() as folder, patch.object(jobs, 'ROOT', Path(folder)):
            df = pd.DataFrame({'Группа': ['A', 'B']})
            token = jobs.create('source.xlsx', df, df.copy())
            state = jobs.load(token)
            state.update(batch_size=1, running=True)
            jobs.save(token, state)
            jobs.process_next(token, lambda *args: b'first')
            def fail(*args):
                raise ValueError('failure')
            jobs.process_next(token, fail)
            state = jobs.load(token)
            self.assertEqual(state['done'], 1)
            self.assertEqual(state['pending']['Группа'].tolist(), ['B'])
            self.assertEqual(len(state['results']), 1)
            self.assertEqual(state['error'], 'failure')
            state['running'] = True
            jobs.save(token, state)
            jobs.process_next(token, lambda *args: b'second')
            self.assertEqual(jobs.load(token)['done'], 2)

    def test_invalid_tokens(self):
        with self.assertRaises(ValueError):
            jobs.directory('../other')

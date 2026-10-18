"""Disk checkpoints for resumable Talabot jobs; no process-wide data cache."""
from contextlib import contextmanager
import fcntl
import os
from pathlib import Path
import pickle
import re
import secrets
import tempfile

ROOT = Path(__file__).with_name('output') / 'batch_jobs'


def directory(token):
    if not re.fullmatch(r'[a-f0-9]{48}', token):
        raise ValueError('Некорректная ссылка задания.')
    return ROOT / token


def save(token, state):
    folder = directory(token)
    folder.mkdir(parents=True, exist_ok=True, mode=0o700)
    fd, name = tempfile.mkstemp(dir=folder)
    try:
        with os.fdopen(fd, 'wb') as stream:
            pickle.dump(state, stream, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(name, folder / 'state.pkl')
    finally:
        if os.path.exists(name):
            os.unlink(name)


def load(token):
    with (directory(token) / 'state.pkl').open('rb') as stream:
        return pickle.load(stream)


def create(filename, df, settings):
    token = secrets.token_hex(24)
    save(token, dict(filename=filename, pending=df, settings=settings,
                     total=len(settings), done=0, results=[], running=False, error=''))
    return token


@contextmanager
def locked(token):
    with (directory(token) / 'lock').open('a') as stream:
        try:
            fcntl.flock(stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            yield False
            return
        try:
            yield True
        finally:
            fcntl.flock(stream, fcntl.LOCK_UN)


def process_next(token, generate):
    """Commit output and queue checkpoint together; retry safely after interruption."""
    with locked(token) as acquired:
        if not acquired:
            return
        state = load(token)
        if not state['running'] or state['pending'].empty:
            return
        settings = state['settings'].iloc[state['done']:state['done'] + state['batch_size']]
        pending = state['pending']
        selected = pending['Группа'].isin(settings['Группа'])
        part = pending.loc[selected].copy()
        part.attrs['formats'] = {k: v for k, v in pending.attrs.get('formats', {}).items() if k[0] in part.index}
        try:
            contents = generate(part, settings)
            end = state['done'] + len(settings)
            filename = f'Талабот_{state["done"] + 1:03d}-{end:03d}.docx'
            output = directory(token) / filename
            temporary = output.with_suffix('.tmp')
            temporary.write_bytes(contents)
            os.replace(temporary, output)
            remaining = pending.loc[~selected].copy()
            remaining.attrs['formats'] = {k: v for k, v in pending.attrs.get('formats', {}).items() if k[0] in remaining.index}
            state.update(pending=remaining, done=end, running=end < state['total'], error='')
            state['results'].append(filename)
        except Exception as error:
            state.update(running=False, error=str(error))
        save(token, state)

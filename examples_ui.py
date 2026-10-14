"""Downloadable templates, filled workbooks and results from those workbooks."""
from pathlib import Path
import pandas as pd
import streamlit as st

EXAMPLES = Path(__file__).with_name('examples')


def show_examples(kind, template):
    with st.expander('Шаблон и наглядные примеры'):
        st.caption('Скачайте шаблон, заполненный Excel и Word, созданный из этого Excel. Данные примеров учебные.')
        items = [('Скачать используемый шаблон Word', template, 'docx'),
                 ('Скачать пример заполнения Excel', EXAMPLES / f'{kind}.xlsx', 'xlsx'),
                 ('Скачать пример результата Word', EXAMPLES / f'{kind}.docx', 'docx')]
        for index, (label, path, extension) in enumerate(items):
            if path.exists():
                st.download_button(label, path.read_bytes(), file_name=path.name, mime=f'application/vnd.openxmlformats-officedocument.{"spreadsheetml.sheet" if extension == "xlsx" else "wordprocessingml.document"}', key=f'example_{kind}_{index}', on_click='ignore')
        workbook = EXAMPLES / f'{kind}.xlsx'
        if workbook.exists():
            preview = pd.read_excel(workbook, nrows=5, keep_default_na=False)
            if 'Дата' in preview:
                preview['Дата'] = pd.to_datetime(preview['Дата']).dt.strftime('%d.%m.%Y')
            st.caption('Пример заполнения · первые пять строк')
            st.dataframe(preview, hide_index=True)
        previews = sorted(EXAMPLES.glob(f'{kind}_page_*.png'))
        for index, path in enumerate(previews, 1):
            st.image(str(path), caption=f'Пример результата · страница {index}')

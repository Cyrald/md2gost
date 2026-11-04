"""
Общие TypedDict-описания параметров задачи конвертации.
Используются и слоем GUI, и рабочим процессом конвейера.
"""
from typing import Optional, TypedDict


class ConversionParams(TypedDict):
    """
    Параметры, которые GUI собирает до выбора пути сохранения.

    Пути хранятся строками, а не ``pathlib.Path``: словарь сериализуется
    в JSON и передаётся рабочему процессу конвейера через stdin, а ``Path``
    в JSON не сериализуется. Принимающая сторона сама оборачивает значение,
    например ``Path(params["md_path"])``.
    """
    md_path: str
    bib_path: Optional[str]
    title_name: Optional[str]
    style_name: Optional[str]
    use_toc: bool
    number_headings: bool

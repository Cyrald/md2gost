"""
Сборка командной строки pandoc и её запуск в подпроцессе.
"""
import os
import shlex
import subprocess
import logging
from pathlib import Path
from typing import List, Optional

logger = logging.getLogger(__name__)

# В Windows скрывает консольные окна pandoc и его фильтров, когда GUI
# запущен через pythonw. На других платформах флага нет — подставляется 0.
_NO_WINDOW_FLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0)

# Аргументы, которые программа формирует сама. В файле pandoc_args профиля
# они запрещены: повтор сломал бы порядок фильтров или подменил вход/выход.
_MANAGED_LONG_OPTIONS = (
    "--output", "--to", "--write", "--from", "--read",
    "--filter", "--lua-filter", "--metadata-file", "--reference-doc",
    "--citeproc", "--bibliography", "--csl",
    "--toc", "--table-of-contents", "--toc-depth", "--resource-path",
)
_MANAGED_SHORT_OPTIONS = ("-o", "-t", "-w", "-f", "-r", "-F", "-L", "-C")


def _is_managed_option(arg: str) -> bool:
    if arg.startswith("--"):
        return any(arg == opt or arg.startswith(opt + "=") for opt in _MANAGED_LONG_OPTIONS)
    return any(arg.startswith(opt) for opt in _MANAGED_SHORT_OPTIONS)


def read_pandoc_args(path: Path) -> List[str]:
    """
    Читает файл дополнительных аргументов pandoc из профиля ГОСТ.

    Аргументы разделяются пробелами и переводами строк, кавычки работают
    как в shell, строки с «#» в начале — комментарии. Аргументы, которыми
    управляет сама программа, отклоняются с ValueError.
    """
    args: List[str] = []
    for line_no, raw_line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        try:
            parts = shlex.split(line)
        except ValueError as exc:
            raise ValueError(f"{path.name}, строка {line_no}: {exc}") from exc
        for arg in parts:
            if _is_managed_option(arg):
                raise ValueError(
                    f"{path.name}, строка {line_no}: аргумент '{arg}' задаётся "
                    "программой и не может быть указан в pandoc_args."
                )
        args.extend(parts)
    return args


class PandocRunner:
    """
    Формирует и выполняет команду pandoc.

    Наличие обязательных файлов проверяется до запуска, чтобы пользователь
    получил понятное сообщение с путём, а не ошибку, потерянную в stderr pandoc.

    Отмена выполняется снаружи — завершением всей группы процессов
    рабочего процесса конвейера, поэтому здесь она не обрабатывается.
    """

    def __init__(
        self,
        pandoc_path: Path,
        crossref_path: Optional[Path] = None,
        metadata_file_path: Optional[Path] = None,
        lua_filter_path: Optional[Path] = None,
        pandoc_args_path: Optional[Path] = None,
    ) -> None:
        """
        Проверяет и сохраняет пути к программам.

        pandoc-crossref запускается и без YAML-конфига (со своими настройками
        по умолчанию), а вот конфиг без pandoc-crossref бессмыслен — ValueError.
        Отсутствие любого обязательного файла — FileNotFoundError.
        Lua-фильтр необязателен: если путь указан, но файла нет, фильтр
        отключается с предупреждением. Файл pandoc_args профиля читается
        сразу, чтобы ошибка в нём была видна до запуска pandoc.
        """
        if metadata_file_path is not None and crossref_path is None:
            raise ValueError("crossref config указан без pandoc-crossref.")

        required_files = [("pandoc", pandoc_path)]
        if crossref_path is not None:
            required_files.append(("pandoc-crossref", crossref_path))
        if metadata_file_path is not None:
            required_files.append(("crossref config", metadata_file_path))
        if pandoc_args_path is not None:
            required_files.append(("файл аргументов pandoc", pandoc_args_path))

        for label, path in required_files:
            if not path.exists():
                raise FileNotFoundError(f"Не найден {label}: {path}")

        self.pandoc_path        = pandoc_path
        self.crossref_path      = crossref_path
        self.metadata_file_path = metadata_file_path
        self.extra_args = read_pandoc_args(pandoc_args_path) if pandoc_args_path else []

        if lua_filter_path and lua_filter_path.exists():
            self.lua_filter_path = lua_filter_path
        else:
            self.lua_filter_path = None
            if lua_filter_path:
                logger.warning("Lua-фильтр не найден: %s", lua_filter_path)

    def run_conversion(
        self,
        input_md: Path,
        output_body: Path,
        style_path: Optional[Path] = None,
        bib_path: Optional[Path] = None,
        csl_path: Optional[Path] = None,
        use_toc: bool = False,
        original_resource_dir: Optional[Path] = None,
    ) -> None:
        """
        Конвертирует Markdown в .docx средствами pandoc.

        pandoc применяет фильтры в порядке аргументов, поэтому порядок здесь
        значим: Lua-фильтр профиля, затем pandoc-crossref и только после них
        --citeproc — иначе citeproc принял бы ссылки @fig:/@tbl: за цитаты.

        pandoc запускается из папки предобработанного файла; в --resource-path
        дополнительно добавляется папка исходного .md, чтобы находились
        изображения, указанные относительными путями.

        Вывод читается через communicate(), который сам вычитывает оба канала
        и исключает взаимоблокировку при переполнении буфера.

        При ненулевом коде возврата pandoc выбрасывает RuntimeError.
        """
        args = [str(self.pandoc_path), str(input_md), "-o", str(output_body),
                "--to=docx", *self.extra_args]

        if self.lua_filter_path:
            args.extend(["--lua-filter", str(self.lua_filter_path)])

        if self.crossref_path:
            args.extend(["--filter", str(self.crossref_path)])
            if self.metadata_file_path:
                args.extend(["--metadata-file", str(self.metadata_file_path)])

        if style_path:
            args.extend(["--reference-doc", str(style_path)])
        if bib_path:
            args.extend(["--citeproc", "--bibliography", str(bib_path)])
            if csl_path:
                args.extend(["--csl", str(csl_path)])
        if use_toc:
            args.extend(["--toc", "--toc-depth=3"])

        cwd_dir      = str(input_md.parent)
        resource_dir = cwd_dir
        if original_resource_dir is not None:
            orig = str(original_resource_dir)
            if orig != cwd_dir:
                resource_dir = cwd_dir + os.pathsep + orig
        args.extend(["--resource-path", resource_dir])

        logger.debug("Pandoc command: %s", shlex.join(args))

        process = subprocess.Popen(
            args,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            # pandoc всегда пишет в UTF-8, а в Windows кодировка каналов по
            # умолчанию — системная кодовая страница, что дало бы UnicodeDecodeError.
            encoding="utf-8",
            env=os.environ,
            cwd=cwd_dir,
            creationflags=_NO_WINDOW_FLAGS,
        )
        stdout, stderr = process.communicate()

        if process.returncode != 0:
            parts = [p.strip() for p in (stderr, stdout) if p and p.strip()]
            error_msg = "\n".join(parts)
            logger.error("Pandoc завершился с кодом %d: %s", process.returncode, error_msg)
            raise RuntimeError(f"Ошибка генерации тела документа:\n{error_msg}")

        if stderr.strip():
            logger.warning("Pandoc (предупреждения): %s", stderr.strip())
        if stdout.strip():
            logger.debug("Pandoc (stdout): %s", stdout.strip())

        logger.info("Тело документа успешно сгенерировано.")

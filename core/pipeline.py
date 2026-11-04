"""
Управление полным конвейером конвертации Markdown → DOCX.
"""
import codecs
import locale
import re
import logging
import shutil
import os
import sys
from pathlib import Path
from typing import Optional
from .config import TaskConfig
from .merger import attach_title_page
from .pandoc_runner import PandocRunner
from tools.markdown_mask import mask_markdown_code
from tools.preprocess import preprocess_markdown, _number_headings

logger = logging.getLogger(__name__)

# Префиксы ссылок pandoc-crossref (@fig:…, @tbl:… и т. д.): такие ссылки
# не являются библиографическими цитатами.
_CROSSREF_KEY_RE = re.compile(r"(?:fig|tbl|eq|sec|lst):")

class UserFacingError(RuntimeError):
    """
    Ожидаемая ошибка с сообщением, написанным для пользователя.
    Рабочий процесс выводит её в журнал без трассировки.
    """


def _normalize_newlines(text: str) -> str:
    return text.replace('\r\n', '\n').replace('\r', '\n')


def _system_encodings() -> list[str]:
    """
    Возвращает устаревшие (не UTF-8) кодировки, настроенные в ОС, от более
    специфичной к общей: в Windows — кодовые страницы ANSI и OEM, иначе —
    кодировку локали. Варианты UTF-8 и повторы исключаются, так как UTF-8
    всегда пробуется раньше.
    """
    candidates: list[str] = []
    if sys.platform == "win32":
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            candidates += [f"cp{kernel32.GetACP()}", f"cp{kernel32.GetOEMCP()}"]
        except (AttributeError, OSError):
            pass
    getencoding = getattr(locale, "getencoding", None)
    if getencoding is not None:
        candidates.append(getencoding())
    candidates.append(locale.getpreferredencoding(False))

    result: list[str] = []
    seen = {"utf-8", "utf-8-sig"}
    for name in candidates:
        try:
            canonical = codecs.lookup(name).name
        except (LookupError, TypeError):
            continue
        if canonical not in seen:
            seen.add(canonical)
            result.append(canonical)
    return result


def _decode_text(raw: bytes, file_name: str) -> str:
    """
    Декодирует байты: сначала как UTF-8 (BOM удаляется), затем по очереди
    в каждой системной кодировке из _system_encodings(). Кодировка не
    угадывается — берётся первая, в которой декодирование прошло без ошибок.

    Если не подошла ни одна, выбрасывает ValueError с понятным пользователю
    сообщением вместо малоинформативного UnicodeDecodeError.
    """
    for encoding in ("utf-8", "utf-8-sig"):
        try:
            return raw.decode(encoding).lstrip("\ufeff")
        except UnicodeDecodeError:
            pass

    system = _system_encodings()
    for encoding in system:
        try:
            text = raw.decode(encoding)
        except UnicodeDecodeError:
            continue
        logger.warning(
            "Файл '%s' не в UTF-8; прочитан в системной кодировке %s. "
            "Если текст отображается неверно, пересохраните файл в UTF-8.",
            file_name, encoding,
        )
        return text

    tried = ", ".join(["UTF-8", *system])
    raise ValueError(
        f"Не удалось прочитать файл '{file_name}': он не соответствует ни одной "
        f"из кодировок ({tried}). Пересохраните файл в кодировке UTF-8."
    )


def _detect_and_read(path: Path) -> str:
    """Читает текстовый файл в подходящей кодировке и приводит переводы строк к \\n."""
    return _normalize_newlines(_decode_text(path.read_bytes(), path.name))


class PipelineOrchestrator:
    """
    Выполняет конвертацию по шагам:

      1. Предобработка Markdown: блоки <diagram> рендерятся в PNG, а во
         временную папку записывается .md, где на их месте стоят обычные
         ссылки на изображения.
      2. При включённой опции — нумерация заголовков.
      3. Подготовка .bib: при необходимости перекодирование в UTF-8.
      4. Если подключена библиография — предупреждения об отсутствии цитат
         и блока {#refs}.
      5. pandoc превращает подготовленный Markdown в body.docx.
      6. Если выбран титульный лист, merger.attach_title_page добавляет его
         перед телом. Основным документом остаётся тело (настройки, тема,
         стили), а титульный лист импортируется изолированным блоком,
         не попадающим в оглавление.
      7. Готовый файл копируется по пути сохранения.

    Пути к pandoc и pandoc-crossref определяет вызывающий код и передаёт сюда,
    чтобы их поиск был сосредоточен в одном месте.

    Временную папку создаёт и удаляет EventHandler; рабочий процесс получает
    её через переменную окружения MD2GOST_TMPDIR и передаёт сюда как work_dir.

    Отмена выполняется снаружи — принудительным завершением рабочего процесса
    вместе с дочерними, поэтому здесь она не обрабатывается.
    """

    def __init__(
        self,
        config: TaskConfig,
        pandoc_path: Path,
        crossref_path: Optional[Path],
        work_dir: Path,
        diagram_timeout: int = 30,
    ) -> None:
        self.config          = config
        self.diagram_timeout = diagram_timeout
        self.body_path       = work_dir / "body.docx"

        self.runner = PandocRunner(
            pandoc_path,
            crossref_path,
            config.crossref_config_path,
            lua_filter_path=config.lua_filter_path,
            pandoc_args_path=config.pandoc_args_path,
        )

    def execute(self) -> bool:
        """
        Выполняет шаги конвейера по порядку.
        Возвращает True при успехе; при любой ошибке выбрасывает исключение,
        чтобы рабочий процесс записал его в журнал и завершился с кодом 1.
        """
        logger.info("Запуск пайплайна конвертации...")
        work_dir = self.body_path.parent

        logger.info("Пре-обработка Markdown (рендеринг диаграмм)...")
        md_input = self._preprocess_md(work_dir)

        if self.config.number_headings:
            text = md_input.read_text(encoding="utf-8")
            md_input.write_text(_number_headings(text), encoding="utf-8")

        bib_input = self._prepare_bib(work_dir)

        if bib_input is not None:
            self._patch_bib_markers(md_input)

        logger.info("Генерация тела документа через pandoc...")
        self.runner.run_conversion(
            input_md=md_input,
            output_body=self.body_path,
            style_path=self.config.style_path,
            bib_path=bib_input,
            csl_path=self.config.csl_path,
            use_toc=self.config.use_toc,
            original_resource_dir=self.config.md_path.parent,
        )

        self._write_output(work_dir)

        logger.info("Конвертация завершена. Файл сохранён: %s", self.config.output_path)
        return True

    def _preprocess_md(self, work_dir: Path) -> Path:
        """Рендерит диаграммы и сохраняет подготовленный Markdown во временную папку в UTF-8."""
        source    = _detect_and_read(self.config.md_path)
        processed = preprocess_markdown(
            source,
            work_dir,
            diagram_timeout=self.diagram_timeout,
        )
        out_path = work_dir / "input_processed.md"
        out_path.write_text(processed, encoding="utf-8")
        return out_path

    def _prepare_bib(self, work_dir: Path) -> Optional[Path]:
        """
        Возвращает путь к .bib в кодировке UTF-8, которую требует pandoc.
        Файл, уже записанный в UTF-8, используется на месте; иначе во временной
        папке создаётся перекодированная копия.
        """
        bib_path = self.config.bib_path
        if bib_path is None:
            return None
        if not bib_path.exists():
            raise ValueError(f"Файл библиографии не найден: '{bib_path}'.")

        raw = bib_path.read_bytes()

        try:
            raw.decode("utf-8")
            return bib_path
        except UnicodeDecodeError:
            pass

        utf8_path = work_dir / "bibliography.bib"
        utf8_path.write_text(_decode_text(raw, bib_path.name), encoding="utf-8")
        return utf8_path

    def _patch_bib_markers(self, md_path: Path) -> None:
        """
        Проверяет документ с подключённой библиографией и предупреждает
        об отсутствии цитат и блока {#refs}. Сам файл не изменяется.
        """
        text          = md_path.read_text(encoding="utf-8")
        analysis_text = _mask_markdown_code(text)

        # Предупреждаем, если .bib подключён, но в тексте нет ни одной цитаты.
        # nocite автоматически не добавляется: включать ли все источники,
        # решает пользователь.
        has_citation = any(
            not _CROSSREF_KEY_RE.match(m.group(1))
            for m in re.finditer(r"(?<![\w])-?@([\w][\w:./-]*)", analysis_text)
        )
        # Распознаваемые формы: nocite: '@*'  |  nocite: |\n  @*  |  nocite: "@*, @key"
        has_nocite_all = re.search(
            r"(?m)^[ \t]*nocite[ \t]*:[ \t]*"
            r"(?:[|>][-+0-9]*[ \t]*\n(?:[ \t]*\n|[ \t]+[^\n]*\n)*?[ \t]+)?"
            r"[^\n]*@\*",
            analysis_text,
        )
        if not has_citation and not has_nocite_all:
            logger.warning(
                "Файл библиографии подключён, но в документе нет ни одной цитаты "
                "вида [@ключ]. Все источники из .bib будут отсутствовать в списке "
                "литературы. Добавьте цитаты в текст или укажите nocite: '@*' "
                "в YAML front-matter вручную, чтобы вывести все источники."
            )

        # Ищем div-блок ::: {#refs}. После идентификатора должен идти пробел,
        # класс или закрывающая скобка; опережающая проверка оставляет эту
        # скобку для завершающего \}.
        has_refs_block = re.search(
            r"(?m)^[ \t]*:{3,}[ \t]*\{[^}\n]*#refs(?=[\s.}])[^}\n]*\}[ \t]*$",
            analysis_text,
        )
        if not has_refs_block:
            logger.warning(
                "В документе не найден блок {#refs}. Pandoc поместит список "
                "литературы в самый конец документа без заголовка. Добавьте в "
                "нужное место раздел, например:\n"
                "# Список использованных источников {-}\n\n::: {#refs}\n:::"
            )

    def _write_output(self, work_dir: Path) -> None:
        """
        Собирает итоговый DOCX целиком во временной папке и копирует в папку
        пользователя только этот один файл.

        Ошибки доступа превращаются в понятные сообщения: если файл уже
        существовал, скорее всего, он открыт в Word; если нет — нет прав
        на запись в папку.
        """
        output_path = self.config.output_path
        output_dir = output_path.parent
        if not output_dir.is_dir():
            raise UserFacingError(f"Папка для сохранения не существует: '{output_dir}'.")

        title_path = self.config.title_path
        if title_path and title_path.exists():
            logger.info("Склейка титульного листа и тела документа...")
            final_path = work_dir / "final.docx"
            attach_title_page(
                title=title_path,
                document=self.body_path,
                output=final_path,
            )
        else:
            final_path = self.body_path

        if not final_path.is_file():
            raise RuntimeError("Сборка DOCX не создала итоговый файл.")

        output_existed = output_path.exists()
        try:
            shutil.copyfile(final_path, output_path)
        except PermissionError as exc:
            if output_existed:
                raise UserFacingError(
                    f"Нет доступа к файлу '{output_path.name}'. Возможно, он открыт "
                    "в Word или другой программе — закройте его и повторите конвертацию."
                ) from exc
            raise UserFacingError(
                f"Нет доступа к папке '{output_dir}': недостаточно прав на запись. "
                "Выберите другую папку для сохранения."
            ) from exc
        except OSError:
            # Прерванное копирование (например, при нехватке места) оставляет
            # обрезанный файл, который Word не откроет, — удаляем его.
            output_path.unlink(missing_ok=True)
            raise


def _mask_markdown_code(text: str) -> str:
    """
    Готовит текст к поиску цитат и {#refs}: блоки кода, встроенный код
    и HTML-комментарии заменяются пробелами (см. tools.markdown_mask).
    """
    return mask_markdown_code(text, mask_html_comments=True)

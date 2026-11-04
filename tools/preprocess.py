"""
Препроцессор Markdown перед передачей документа в pandoc.

1. Находит блоки <diagram>, рендерит их в PNG во временной папке и заменяет
   каждый блок обычной ссылкой на изображение Markdown. Каждая диаграмма
   рендерится в отдельном подпроцессе (tools/_diagram_worker.py).
2. Нумерует заголовки H1–H3 (_number_headings; вызывается конвейером
   отдельно, если нумерация включена).

Блоки <diagram> внутри блоков кода, встроенного кода и HTML-комментариев
не трогаются: примеры синтаксиса в документе не исполняются и не заменяются.

Поддерживаемые типы диаграмм
----------------------------
matplotlib
    Тело блока выполняется как код Python через exec(). ``plt``
    (matplotlib.pyplot) доступен сразу, импортировать его не нужно.
    Другие модули (numpy, scipy, pandas и т. д.) импортируются явно.

    Атрибуты:
        type       Обязательный. Имя рендерера.
        label      Необязательный. Метка для перекрёстных ссылок вида ``fig:имя``.
        caption    Необязательный. Подпись рисунка (допускается Markdown).
        width      Необязательный. Например, ``80%``, ``12cm``, ``400px``.
        height     Необязательный. В том же формате, что и width.
        dpi        Разрешение в точках на дюйм, от 1 до 1200. По умолчанию 150.

    Пример:

        <diagram type="matplotlib" label="fig:growth" caption="Скорость роста" width="80%" dpi="200">
        import numpy as np
        x = np.linspace(0, 2 * np.pi, 200)
        plt.plot(x, np.sin(x))
        plt.xlabel("x")
        plt.ylabel("sin(x)")
        </diagram>

Добавление нового рендерера
---------------------------
1. Добавьте функцию рендеринга в ``tools/_diagram_worker.py`` и внесите её
   в словарь ``_RENDERERS`` этого модуля.
2. Добавьте здесь обёртку с декоратором ``@_register("мой_тип")``.
Других файлов менять не нужно.
"""

import re
import sys
import html
import json
import hashlib
import logging
import subprocess
from pathlib import Path

from tools.markdown_mask import mask_markdown_code

logger = logging.getLogger(__name__)

# В Windows скрывает консольные окна дочерних процессов, когда GUI запущен
# через pythonw. На других платформах флага нет — подставляется 0.
_NO_WINDOW_FLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0)

# Нумерация заголовков. Нумеруются только уровни 1–3.
_ATX_HEADING_RE = re.compile(
    r"^(?P<indent> {0,3})(?P<hashes>#{1,3})[ \t]+(?P<title>\S.*?)[ \t]*$"
)
_SETEXT_UNDERLINE_RE = re.compile(r"^ {0,3}(?P<marker>=+|-+)[ \t]*$")
# Строки, которые не могут быть текстом setext-заголовка: списки, цитаты,
# таблицы, ATX-заголовки, HTML и код с отступом.
_NOT_SETEXT_TEXT_RE = re.compile(
    r"^(?: {4,}|\t| {0,3}(?:[>#|<]|[-+*](?:[ \t]|$)|\d+[.)](?:[ \t]|$)))"
)
_NUMBER_PREFIX_RE = re.compile(r"^(?P<number>\d+(?:\.\d+)*)(?:[.)])?\s+")
_HEADING_ATTR_RE = re.compile(r"[ \t]*\{(?P<inner>[^{}]*)\}[ \t]*$")
_ATTR_TOKEN_RE = re.compile(r"""[^\s"'=]+(?:=(?:"[^"]*"|'[^']*'|[^\s"']+))?""")

# Номер, уже написанный автором, принимается, только если каждая его часть
# похожа на номер раздела. Большие значения (например, «2024 год») считаются
# частью названия.
_MAX_SECTION_NUMBER = 99


def _split_heading_attrs(title: str) -> tuple[str, list[str] | None]:
    """
    Отделяет завершающий блок атрибутов pandoc ``{#id .class key=value -}``.

    Возвращает (текст_без_атрибутов, токены). Токены равны None, если
    корректного блока атрибутов нет: завершающие ``{...}``, не являющиеся
    списком атрибутов, — обычный текст.
    """
    match = _HEADING_ATTR_RE.search(title)
    if not match:
        return title, None
    tokens = _ATTR_TOKEN_RE.findall(match.group("inner"))
    if not tokens or not all(
        token == "-" or token[0] in "#." or "=" in token for token in tokens
    ):
        return title, None
    return title[:match.start()], tokens


class _HeadingNumberer:
    """Хранит счётчики H1–H3 и решает, какой номер получит каждый заголовок."""

    def __init__(self) -> None:
        self.counters = [0, 0, 0]
        # True на уровне i — текущий раздел этого уровня не нумеруется,
        # и его подзаголовки тоже остаются без номера.
        self.unnumbered_context = [False, False, False]

    def number(self, level: int, title: str) -> str | None:
        """
        Возвращает номер, который нужно поставить перед *title*, или None,
        если заголовок остаётся как есть.
        """
        text, attr_tokens = _split_heading_attrs(title)
        tokens = attr_tokens or []
        # Заголовок с собственной меткой pandoc-crossref (приложения:
        # {label="А"}) нумеруется ею, каким бы ни было значение метки.
        has_own_label = any(token.startswith("label=") for token in tokens)
        is_unnumbered = (
            "-" in tokens or ".unnumbered" in tokens or has_own_label
        )

        # Заголовок этого уровня открывает новый контекст нумерации; более
        # глубокие уровни наследуют, нумеруется ли текущий заголовок.
        for i in range(level - 1, 3):
            self.unnumbered_context[i] = False

        if is_unnumbered:
            for i in range(level - 1, 3):
                self.unnumbered_context[i] = True
            return None

        # Номер, уже написанный автором, сохраняется, а счётчики
        # синхронизируются с ним.
        existing = self._existing_number(level, text)
        if existing is not None:
            for i in range(level):
                self.counters[i] = existing[i]
            for i in range(level, 3):
                self.counters[i] = 0
            return None

        # Подзаголовок под ненумеруемым или отсутствующим родителем
        # не должен получить бессмысленный номер вроде 0.0.1.
        if any(self.unnumbered_context[:level - 1]) or any(
            self.counters[i] == 0 for i in range(level - 1)
        ):
            for i in range(level - 1, 3):
                self.unnumbered_context[i] = True
            return None

        self.counters[level - 1] += 1
        for i in range(level, 3):
            self.counters[i] = 0
        return ".".join(str(self.counters[i]) for i in range(level))

    def _existing_number(self, level: int, text: str) -> list[int] | None:
        """
        Распознаёт номер, уже написанный автором (``## 2.3 Обзор``).

        Номер принимается, только если в нём ровно *level* частей, каждая
        похожа на номер раздела, родительские части совпадают с текущими
        счётчиками, а последняя часть больше текущего значения. Всё остальное
        («2024 год», «## 5 Пункт») — обычный текст, и заголовок получает
        собственный номер.
        """
        match = _NUMBER_PREFIX_RE.match(text.strip())
        if not match:
            return None
        parts = [int(part) for part in match.group("number").split(".")]
        if len(parts) != level:
            return None
        if any(part <= 0 or part > _MAX_SECTION_NUMBER for part in parts):
            return None
        if any(parts[i] != self.counters[i] for i in range(level - 1)):
            return None
        if parts[-1] <= self.counters[level - 1]:
            return None
        return parts


def _number_headings(text: str) -> str:
    """
    Добавляет номера разделов к заголовкам H1–H3 в тексте Markdown
    (ATX ``#`` и setext ``===``/``---``).

    Без изменений и без учёта в счётчиках остаются:
        * заголовки с ``-`` или ``.unnumbered`` в блоке атрибутов в любом
          сочетании: ``{-}``, ``{.unnumbered .unlisted}``,
          ``{#intro .unnumbered}``, ``{- #intro}`` — вместе с подзаголовками;
        * заголовки, задающие собственный ``label`` в блоке атрибутов
          (``{label="А"}``, любое значение), — вместе с подзаголовками;
        * строки внутри блоков кода, HTML-комментариев и YAML front matter;
        * заголовки H4 и глубже.

    Примеры:
        # Введение {-}    → # Введение {-}
        # Обзор           → # 1 Обзор
        ## Методы         → ## 1.1 Методы
        ### Детали        → ### 1.1.1 Детали
        # 2024 год        → # 2 2024 год
        # ПРИЛОЖЕНИЕ А {label="А"} → # ПРИЛОЖЕНИЕ А {label="А"}
    """
    # Заголовки ищутся в маскированной копии, а изменяются строки оригинала:
    # длина и разбиение на строки у них совпадают.
    lines = text.split("\n")
    masked_lines = mask_markdown_code(
        text, mask_inline=False, mask_html_comments=True, mask_front_matter=True,
    ).split("\n")
    numberer = _HeadingNumberer()
    result: list[str] = []

    for index, line in enumerate(lines):
        masked = masked_lines[index]

        atx = _ATX_HEADING_RE.match(masked)
        if atx:
            original = _ATX_HEADING_RE.match(line)
            if original is not None:
                level = len(original.group("hashes"))
                title = original.group("title")
                num = numberer.number(level, title)
                if num is not None:
                    line = (
                        f"{original.group('indent')}{original.group('hashes')} "
                        f"{num} {title}"
                    )
            result.append(line)
            continue

        # Setext-заголовок: непустая строка после пустой, за которой идёт
        # подчёркивание из «=» (H1) или «-» (H2).
        next_masked = masked_lines[index + 1] if index + 1 < len(masked_lines) else ""
        underline = _SETEXT_UNDERLINE_RE.match(next_masked)
        previous_blank = index == 0 or not masked_lines[index - 1].strip()
        if (
            underline
            and masked.strip()
            and previous_blank
            and not _NOT_SETEXT_TEXT_RE.match(masked)
            and not _SETEXT_UNDERLINE_RE.match(masked)
        ):
            level = 1 if underline.group("marker")[0] == "=" else 2
            indent = line[: len(line) - len(line.lstrip(" "))]
            title = line.strip()
            num = numberer.number(level, title)
            if num is not None:
                line = f"{indent}{num} {title}"
            result.append(line)
            continue

        result.append(line)
    return "\n".join(result)


_WORKER_PATH = Path(__file__).parent / "_diagram_worker.py"


# Значение атрибута type → функция рендеринга; заполняется декоратором _register.
_RENDERERS: dict = {}


def _register(name: str):
    """Декоратор: регистрирует функцию как рендерер диаграмм типа name."""
    def decorator(fn):
        _RENDERERS[name] = fn
        return fn
    return decorator


def _run_in_worker(diag_type: str, code: str, attrs: dict,
                   diagram_timeout: int = 30) -> bytes:
    """
    Передаёт рендеринг отдельному подпроцессу (_diagram_worker.py).

    Подпроцесс получает JSON через stdin и возвращает байты PNG в stdout.
    Отдельный процесс нужен, чтобы упавшая или зависшая диаграмма не
    остановила весь конвейер; песочницей для безопасности он НЕ является.

    При ошибке выбрасывает RuntimeError с текстом из stderr подпроцесса,
    при превышении diagram_timeout секунд — RuntimeError о тайм-ауте.
    """
    payload = json.dumps(
        {"type": diag_type, "code": code, "attrs": attrs},
        ensure_ascii=False,
    ).encode("utf-8")

    try:
        result = subprocess.run(
            [sys.executable, str(_WORKER_PATH)],
            input=payload,
            capture_output=True,
            timeout=diagram_timeout,
            creationflags=_NO_WINDOW_FLAGS,
        )
    except subprocess.TimeoutExpired:
        logger.error(
            "Диаграмма [%s]: рендеринг прерван по таймауту (%d с).",
            diag_type, diagram_timeout,
        )
        raise RuntimeError(
            f"Таймаут рендеринга диаграммы '{diag_type}': "
            f"превышен лимит {diagram_timeout} секунд. "
            "Проверьте, не зависает ли код диаграммы (бесконечный цикл, plt.show() и т.п.)."
        ) from None

    if result.returncode != 0:
        error = result.stderr.decode("utf-8", errors="replace").strip()
        raise RuntimeError(error)

    return result.stdout


@_register("matplotlib")
def _render_matplotlib(code: str, attrs: dict,
                        diagram_timeout: int = 30) -> bytes:
    return _run_in_worker("matplotlib", code, attrs, diagram_timeout)


# Блок <diagram ...>...</diagram>, в том числе многострочный. Строки
# в кавычках внутри тега разбираются отдельно, поэтому символ «>» в значении
# атрибута (например, caption="x > y") не закрывает тег.
_BLOCK_RE = re.compile(
    r'<diagram\s((?:[^>"\']*(?:"[^"]*"|\'[^\']*\'))*[^>"\']*)>(.*?)</diagram>',
    re.DOTALL,
)
_ATTR_PAIR_RE = re.compile(
    r"""([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')"""
)
_LABEL_RE = re.compile(r"fig:[\w-]+")
_SIZE_RE = re.compile(r"\d+(?:\.\d+)?(?:%|cm|mm|in|px|pt|em)?")
# Формулы $...$ и $$...$$ в подписи: квадратные скобки в них не экранируются.
_MATH_SPAN_RE = re.compile(r"(\$\$.+?\$\$|\$[^$\n]+?\$)", re.DOTALL)
_MAX_DPI = 1200


def _parse_attrs(attr_string: str) -> dict:
    """
    Разбирает пары key="value" / key='value' из открывающего тега <diagram>.

    HTML-сущности в значениях раскодируются, поэтому работают и ``A & B``,
    и ``A &amp; B``. Нераспознанные фрагменты пропускаются с предупреждением,
    а не отменяют весь блок.
    """
    attrs: dict = {}
    position = 0
    for match in _ATTR_PAIR_RE.finditer(attr_string):
        gap = attr_string[position:match.start()].strip()
        if gap:
            logger.warning(
                "Диаграмма: в теге <diagram> не распознан фрагмент %r — он проигнорирован.",
                gap,
            )
        key = match.group(1).lower()
        raw_value = match.group(2) if match.group(2) is not None else match.group(3)
        if key in attrs:
            logger.warning(
                "Диаграмма: атрибут %s указан несколько раз — используется последнее значение.",
                key,
            )
        attrs[key] = html.unescape(raw_value)
        position = match.end()

    # Завершающий «/» (как в самозакрывающемся теге) ошибкой не считается.
    tail = attr_string[position:].strip().rstrip("/").strip()
    if tail:
        logger.warning(
            "Диаграмма: в теге <diagram> не распознан фрагмент %r — он проигнорирован.",
            tail,
        )
    return attrs


def _escape_caption(caption: str) -> str:
    """
    Делает подпись безопасной для вставки внутрь ``![...]``.

    Квадратные скобки вне формул экранируются, чтобы подпись не закрыла
    синтаксис изображения раньше времени и не превратилась в ссылку;
    переводы строк заменяются пробелами; обратные косые черты в конце
    удаляются, чтобы они не экранировали закрывающую скобку.
    """
    caption = re.sub(r"\s*\n\s*", " ", caption).strip().rstrip("\\")
    # split с группой даёт чередование: чётные элементы — текст, нечётные — формулы.
    parts = _MATH_SPAN_RE.split(caption)
    for i in range(0, len(parts), 2):
        parts[i] = re.sub(r"(?<!\\)([\[\]])", r"\\\1", parts[i])
    return "".join(parts)


def _validate_diagram_attrs(diag_type: str, attrs: dict) -> tuple[str, str, str, str] | None:
    """
    Проверяет атрибуты одного блока.

    Возвращает (caption, label, width, height), готовые для Markdown, или
    None, если блок нужно пропустить (недопустимый dpi). Некорректные label,
    width и height отбрасываются с предупреждением. Все сообщения помечены
    как ошибки атрибутов, чтобы их не путали с ошибками в коде диаграммы.
    """
    caption = attrs.get("caption", "")
    label = attrs.get("label", "").strip()
    width = attrs.get("width", "").strip()
    height = attrs.get("height", "").strip()

    if "dpi" in attrs:
        try:
            dpi = float(attrs["dpi"])
        except ValueError:
            dpi = -1
        if not 0 < dpi <= _MAX_DPI:
            logger.error(
                "Диаграмма [%s]: ошибка в атрибутах — dpi=%r, допустимо число от 1 до %d. Блок пропущен.",
                diag_type, attrs["dpi"], _MAX_DPI,
            )
            return None

    if label and not _LABEL_RE.fullmatch(label):
        logger.warning(
            "Диаграмма [%s]: ошибка в атрибутах — label=%r должен иметь вид fig:имя "
            "(буквы, цифры, _ и -). Метка проигнорирована.",
            diag_type, label,
        )
        label = ""
    for name, value in (("width", width), ("height", height)):
        if value and not _SIZE_RE.fullmatch(value):
            logger.warning(
                "Диаграмма [%s]: ошибка в атрибутах — %s=%r, ожидается число с единицей "
                "(%%, cm, mm, in, px, pt, em). Атрибут проигнорирован.",
                diag_type, name, value,
            )
    # «-» заведомо не проходит проверку, поэтому пустое значение тоже даёт "".
    width = width if _SIZE_RE.fullmatch(width or "-") else ""
    height = height if _SIZE_RE.fullmatch(height or "-") else ""

    if not label:
        logger.warning(
            "Диаграмма [%s]: атрибут label не задан — "
            "рисунок не будет пронумерован и ссылки на него невозможны.",
            diag_type,
        )
    if not caption:
        logger.warning(
            "Диаграмма [%s]: атрибут caption не задан — "
            "рисунок не получит подпись.",
            diag_type,
        )
    return _escape_caption(caption), label, width, height


def _img_filename(img_bytes: bytes) -> str:
    """
    Возвращает имя файла по хешу содержимого изображения: одинаковые
    диаграммы дают один файл. Выбрасывает ValueError, если данные не PNG —
    единственный формат, который сейчас выдаёт подпроцесс.
    """
    if img_bytes[:4] != b'\x89PNG':
        raise ValueError(
            f"Рендерер вернул данные в неожиданном формате "
            f"(magic bytes: {img_bytes[:4]!r}). Ожидался PNG."
        )
    digest = hashlib.sha256(img_bytes).hexdigest()[:20]
    return f"diagram_{digest}.png"


def _process_diagram_block(
    attr_str: str,
    code: str,
    work_dir: Path,
    diagram_timeout: int,
) -> str:
    """
    Рендерит один блок <diagram> в PNG внутри work_dir и возвращает ссылку
    на изображение в синтаксисе Markdown.

    Если блок отрендерить нельзя, возвращает пустую строку (ошибка
    записывается в журнал).
    """
    code = code.strip()
    attrs = _parse_attrs(attr_str)

    diag_type = attrs.get("type", "").strip().lower()
    if not diag_type:
        logger.error(
            "Диаграмма: ошибка в атрибутах — не задан обязательный атрибут type. Блок пропущен."
        )
        return ""

    renderer = _RENDERERS.get(diag_type)
    if renderer is None:
        logger.error(
            "Диаграмма: ошибка в атрибутах — неизвестный type=%r (доступно: %s). Блок пропущен.",
            diag_type, ", ".join(sorted(_RENDERERS)),
        )
        return ""

    validated = _validate_diagram_attrs(diag_type, attrs)
    if validated is None:
        return ""
    caption, label, width, height = validated

    try:
        img_bytes = renderer(code, attrs, diagram_timeout)
        fname = _img_filename(img_bytes)
    except Exception as exc:
        logger.error(
            "Диаграмма [%s]: %s. Блок пропущен.", diag_type, str(exc).strip().rstrip("."),
        )
        return ""

    img_path = work_dir / fname
    img_path.write_bytes(img_bytes)
    logger.info("Диаграмма [%s] → %s", diag_type, fname)

    img = f"![{caption}]({fname})"
    attr_parts = []
    if label:
        attr_parts.append(f"#{label}")
    if width:
        attr_parts.append(f'width="{width}"')
    if height:
        attr_parts.append(f'height="{height}"')
    if attr_parts:
        # Блок атрибутов должен идти сразу за изображением без пробела,
        # иначе pandoc воспримет его как обычный текст.
        img += "{" + " ".join(attr_parts) + "}"

    return img


def preprocess_markdown(md_text: str, work_dir: Path,
                         diagram_timeout: int = 30) -> str:
    """
    Находит в *md_text* блоки <diagram>, рендерит каждый в PNG внутри
    *work_dir* и возвращает текст, где каждый блок заменён ссылкой
    на изображение Markdown.

    Блоки внутри блоков кода, встроенного кода и HTML-комментариев
    пропускаются: поиск идёт по маскированной копии текста той же длины,
    а по позициям совпадений вырезаются фрагменты оригинала.

    Блоки, которые не удалось отрендерить, удаляются из текста, а ошибка
    записывается в журнал.
    """
    masked = mask_markdown_code(md_text, mask_html_comments=True)
    pieces: list[str] = []
    position = 0
    for match in _BLOCK_RE.finditer(masked):
        attr_str = md_text[match.start(1):match.end(1)]
        code = md_text[match.start(2):match.end(2)]
        pieces.append(md_text[position:match.start()])
        pieces.append(_process_diagram_block(attr_str, code, work_dir, diagram_timeout))
        position = match.end()
    pieces.append(md_text[position:])
    return "".join(pieces)

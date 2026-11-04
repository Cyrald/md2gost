"""
Маскирование участков Markdown, которые нельзя разбирать как разметку.

Каждый маскируемый символ заменяется пробелом (переводы строк сохраняются),
поэтому маскированный текст совпадает с исходным по длине и разбиению на
строки. Вызывающий код ищет совпадения в маскированном тексте, а по их
позициям вырезает фрагменты из исходного.

Маскируются:
    * огороженные блоки кода (``` / ~~~, с любым отступом, закрытие по правилам CommonMark);
    * встроенный код в обратных кавычках (не через пустую строку, т. е. в пределах абзаца);
    * по запросу — HTML-комментарии ``<!-- ... -->``;
    * по запросу — блок YAML front matter в самом начале документа.
"""
import re

_FENCE_RE = re.compile(r"^[ \t]*(`{3,}|~{3,})(.*)$")
_FRONT_MATTER_OPEN = "---"
_FRONT_MATTER_CLOSE = ("---", "...")

# Встроенный код — это N обратных кавычек, содержимое и ровно N обратных
# кавычек. Внутри допустимы одиночные переводы строки (один абзац), но не
# пустая строка: непарная кавычка не должна поглощать следующие абзацы.
_INLINE_CODE_RE = re.compile(
    r"(?<!`)(`+)(?!`)((?:(?!\n[ \t]*\r?\n).)*?)(?<!`)\1(?!`)",
    re.DOTALL,
)
_HTML_COMMENT_RE = re.compile(r"<!--.*?(?:-->|\Z)", re.DOTALL)


def _blank(fragment: str) -> str:
    return "".join(char if char in "\r\n" else " " for char in fragment)


def _blank_match(match: re.Match) -> str:
    return _blank(match.group(0))


def _front_matter_end(lines: list[str]) -> int:
    """Возвращает число начальных строк, занятых YAML front matter (0, если блока нет или он не закрыт)."""
    if not lines or lines[0].rstrip("\r\n").rstrip() != _FRONT_MATTER_OPEN:
        return 0
    for index in range(1, len(lines)):
        if lines[index].rstrip("\r\n").rstrip() in _FRONT_MATTER_CLOSE:
            return index + 1
    return 0


def mask_markdown_code(
    text: str,
    *,
    mask_inline: bool = True,
    mask_html_comments: bool = False,
    mask_front_matter: bool = False,
) -> str:
    """
    Возвращает *text*, в котором код (и по запросу комментарии и front matter)
    заменён пробелами.

    Длина результата всегда равна длине *text*, поэтому найденные в нём
    позиции действительны и для исходного текста.
    """
    lines = text.splitlines(keepends=True)
    masked: list[str] = []

    start = _front_matter_end(lines) if mask_front_matter else 0
    masked.extend(_blank(line) for line in lines[:start])

    in_fence = False
    fence_char = ""
    fence_len = 0

    for line in lines[start:]:
        fence = _FENCE_RE.match(line.rstrip("\r\n"))
        if in_fence:
            if (
                fence
                and fence.group(1)[0] == fence_char
                and len(fence.group(1)) >= fence_len
                and not fence.group(2).strip()
            ):
                in_fence = False
                fence_char = ""
                fence_len = 0
            masked.append(_blank(line))
            continue

        # По CommonMark строка-описание после ``` не может содержать обратных
        # кавычек: такая строка — встроенный код, а не начало блока.
        if fence and not (fence.group(1)[0] == "`" and "`" in fence.group(2)):
            in_fence = True
            fence_char = fence.group(1)[0]
            fence_len = len(fence.group(1))
            masked.append(_blank(line))
            continue

        masked.append(line)

    result = "".join(masked)
    if mask_inline:
        result = _INLINE_CODE_RE.sub(_blank_match, result)
    if mask_html_comments:
        result = _HTML_COMMENT_RE.sub(_blank_match, result)
    return result

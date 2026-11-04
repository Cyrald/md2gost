"""
Изолированный подпроцесс для рендеринга одного блока <diagram>.

Читает из stdin один JSON-объект:
    {"type": str, "code": str, "attrs": {str: str}}

При успехе пишет в stdout байты PNG и завершается с кодом 0.
При ошибке пишет сообщение в stderr и завершается с кодом 1.

ВНИМАНИЕ: код диаграмм исполняется без ограничений через exec().
Запускайте только доверенный код.
"""
import io
import sys
import json
import traceback

_MAX_DPI = 1200
_CODE_FILENAME = "<diagram>"


class DiagramAttributeError(ValueError):
    """Недопустимое значение атрибута тега <diagram>."""


class DiagramCodeError(RuntimeError):
    """Код диаграммы выбросил исключение или ничего не нарисовал."""


def _describe_code_exception(exc: BaseException) -> str:
    """
    Возвращает «строка N: Тип: сообщение», где N — номер строки в коде
    диаграммы (самый глубокий кадр трассировки, относящийся к этому коду).
    """
    line_no = None
    if isinstance(exc, SyntaxError) and exc.filename == _CODE_FILENAME:
        line_no = exc.lineno
    else:
        for frame in traceback.extract_tb(exc.__traceback__):
            if frame.filename == _CODE_FILENAME:
                line_no = frame.lineno
    where = f"строка {line_no}: " if line_no else ""
    return f"{where}{type(exc).__name__}: {exc}"


def _figure_is_empty(fig) -> bool:
    """Истина, если на фигуре и её осях нет ни одного видимого элемента."""
    if fig.texts or fig.images or fig.lines or fig.patches or fig.artists or fig.legends:
        return False
    for ax in fig.axes:
        if ax.has_data() or ax.texts or ax.artists or ax.tables or ax.get_title():
            return False
    return True


def _render_matplotlib(code: str, attrs: dict) -> bytes:
    """Выполняет код matplotlib и возвращает текущую фигуру в формате PNG."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    raw_dpi = attrs.get("dpi", 150)
    try:
        dpi = int(float(raw_dpi))
    except (TypeError, ValueError):
        raise DiagramAttributeError(
            f"dpi={raw_dpi!r} не является числом."
        ) from None
    if not 0 < dpi <= _MAX_DPI:
        raise DiagramAttributeError(
            f"dpi={raw_dpi!r}, допустимо число от 1 до {_MAX_DPI}."
        )

    namespace: dict = {
        "__name__": "__diagram__",
        "plt": plt,
    }

    # С бэкендом Agg plt.show() ничего не показывает, но выдаёт UserWarning,
    # который попал бы в журнал GUI; подменяем его пустой функцией.
    plt.show = lambda *args, **kwargs: None  # noqa: E731

    try:
        exec(compile(code, _CODE_FILENAME, "exec"), namespace)  # noqa: S102
    except Exception as exc:
        raise DiagramCodeError(_describe_code_exception(exc)) from exc

    if not plt.get_fignums() or _figure_is_empty(plt.gcf()):
        plt.close("all")
        raise DiagramCodeError(
            "код выполнился, но ничего не нарисовал (нет графиков, текста или "
            "изображений на текущей фигуре)."
        )

    buf = io.BytesIO()
    plt.savefig(buf, format="png", bbox_inches="tight", dpi=dpi)
    plt.close("all")
    buf.seek(0)
    return buf.read()


# Значение атрибута type тега <diagram> → функция рендеринга.
_RENDERERS: dict = {
    "matplotlib": _render_matplotlib,
}


def main() -> None:
    # Перенаправляем stdout на stderr, чтобы print() из кода диаграммы
    # не засорял бинарный PNG-канал. PNG пишется только в _real_stdout.
    _real_stdout = sys.stdout.buffer
    sys.stdout   = sys.stderr

    try:
        payload   = json.loads(sys.stdin.buffer.read().decode("utf-8"))
        diag_type = payload["type"]
        code      = payload["code"]
        attrs     = payload["attrs"]

        renderer = _RENDERERS.get(diag_type)
        if renderer is None:
            print(
                f"Неизвестный тип диаграммы: {diag_type!r}.",
                file=sys.stderr,
            )
            sys.exit(1)

        png_bytes = renderer(code, attrs)
        _real_stdout.write(png_bytes)
        sys.exit(0)

    except DiagramAttributeError as exc:
        print(f"ошибка в атрибутах — {exc}", file=sys.stderr)
        sys.exit(1)
    except DiagramCodeError as exc:
        print(f"ошибка в коде диаграммы — {exc}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(f"ошибка рендеринга — {type(exc).__name__}: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

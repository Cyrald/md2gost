"""Вспомогательные средства журналирования, не зависящие от GUI-библиотеки."""
import logging
import sys
from typing import List


# Маркер «аргумент не передан»: None здесь не годится, потому что
# sys.stdout/sys.stderr сами бывают None (например, при запуске через pythonw).
_USE_PROCESS_STREAM = object()


class StartupLogBuffer(logging.Handler):
    """Хранит предупреждения, возникшие до появления консоли журнала GUI."""

    def __init__(self) -> None:
        super().__init__(logging.WARNING)
        self.records: List[logging.LogRecord] = []

    def emit(self, record: logging.LogRecord) -> None:
        self.records.append(record)


def mirror_log_message(
    msg: str,
    level: str,
    stdout=_USE_PROCESS_STREAM,
    stderr=_USE_PROCESS_STREAM,
) -> None:
    """
    Дублирует сообщение консоли GUI в поток консоли процесса.

    Предупреждения и ошибки пишутся в stderr, остальное — в stdout. Если
    нужный поток отсутствует, используется второй; если нет обоих, сообщение
    просто не дублируется.
    """
    stdout = sys.stdout if stdout is _USE_PROCESS_STREAM else stdout
    stderr = sys.stderr if stderr is _USE_PROCESS_STREAM else stderr
    to_stderr = level in {"WARNING", "ERROR", "CRITICAL"}
    stream = stderr if to_stderr else stdout
    if stream is None:
        stream = stdout if to_stderr else stderr
    if stream is None:
        return

    try:
        stream.write(msg)
        stream.flush()
    except UnicodeEncodeError:
        # Кодировка консоли Windows (например, cp866) может не вмещать часть
        # символов: заменяем их, чтобы сообщение всё же было выведено.
        encoding = getattr(stream, "encoding", None) or "utf-8"
        safe_msg = msg.encode(encoding, errors="replace").decode(
            encoding, errors="replace",
        )
        try:
            stream.write(safe_msg)
            stream.flush()
        except (OSError, UnicodeError, ValueError):
            pass
    except (OSError, ValueError):
        # Закрытая консоль не должна мешать обновлению журнала в GUI.
        pass

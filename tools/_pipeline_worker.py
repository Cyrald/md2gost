"""
Точка входа изолированного подпроцесса, выполняющего весь конвейер конвертации.

Протокол
--------
stdin   — один JSON-объект (UTF-8) со всеми параметрами задачи.
stdout  — журнал в виде JSON-строк, по одной на запись:
              {"level": "INFO", "msg": "..."}
stderr  — «сырой» вывод: трассировки Python и print(); EventHandler._read_stderr
          показывает его в консоли GUI с уровнем WARNING.
код 0   — конвертация успешна.
код 1   — конвертация не удалась (ошибка уже записана в stdout).

EventHandler запускает процесс в отдельной группе (Unix — start_new_session=True,
Windows — CREATE_NEW_PROCESS_GROUP), чтобы при отмене одним действием завершить
и его, и все дочерние процессы: pandoc, pandoc-crossref, рендеринг диаграмм.
"""
import os
import sys
import json
import logging
from pathlib import Path

# Делаем корень проекта импортируемым независимо от текущего каталога:
# файл лежит в <корень>/tools/, поэтому корень — parent.parent.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


class _JsonLineHandler(logging.Handler):
    """
    Пишет записи журнала JSON-строками в текстовый поток (stdout процесса).
    Каждая строка сбрасывается сразу, чтобы GUI получал сообщения по мере
    выполнения, а не одним блоком после завершения процесса.
    """

    def __init__(self, stream) -> None:
        super().__init__()
        self._stream = stream

    def emit(self, record: logging.LogRecord) -> None:
        try:
            line = json.dumps(
                {"level": record.levelname, "msg": self.format(record)},
                ensure_ascii=False,
            )
            self._stream.write(line + "\n")
            self._stream.flush()
        except Exception:
            self.handleError(record)


def _setup_logging(stream) -> None:
    """
    Подключает JSON-обработчик только к логгерам проекта (core, tools).
    Корневой логгер не трогаем, чтобы сообщения сторонних библиотек
    не попадали в консоль GUI.
    """
    handler = _JsonLineHandler(stream)
    handler.setFormatter(
        logging.Formatter(
            "[%(asctime)s] %(levelname)s  %(message)s",
            datefmt="%H:%M:%S",
        )
    )
    for name in ("core", "tools"):
        lgr = logging.getLogger(name)
        lgr.setLevel(logging.DEBUG)
        lgr.addHandler(handler)


def main() -> None:
    """Читает задачу из stdin, выполняет конвейер и завершает процесс с кодом 0 или 1."""
    # В Windows каналы по умолчанию используют системную кодовую страницу
    # (например, cp1251), в которой нет многих символов из сообщений журнала.
    # Родитель передаёт PYTHONIOENCODING=utf-8, но если переменная проигнорирована
    # или переопределена, явная перенастройка потоков на UTF-8 страхует от этого.
    import io
    if hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True
        )
    if hasattr(sys.stderr, "buffer"):
        sys.stderr = io.TextIOWrapper(
            sys.stderr.buffer, encoding="utf-8", errors="replace", line_buffering=True
        )

    # stdout зарезервирован под JSON-журнал, поэтому сохраняем его отдельно,
    # а print() и трассировки перенаправляем в stderr — GUI покажет их как WARNING.
    real_stdout = sys.stdout
    sys.stdout  = sys.stderr

    _setup_logging(real_stdout)

    # Логгер ещё не может сообщить о битом JSON в нужном формате,
    # поэтому ошибку разбора пишем в протокол вручную.
    try:
        payload = json.loads(sys.stdin.read())
    except Exception as exc:
        real_stdout.write(
            json.dumps({"level": "ERROR", "msg": f"Ошибка чтения задачи: {exc}"},
                       ensure_ascii=False) + "\n"
        )
        real_stdout.flush()
        sys.exit(1)

    # UserFacingError импортируется отдельно: он нужен в except ниже, и если
    # модули конвертера не загружаются, об этом надо сообщить до запуска задачи.
    try:
        from core.pipeline import UserFacingError
    except Exception:
        logging.getLogger("core").exception("Не удалось загрузить модули конвертера:")
        sys.exit(1)

    try:
        from core.config import TaskConfig
        from core.pipeline import PipelineOrchestrator

        task = TaskConfig(
            md_path=Path(payload["md_path"]),
            bib_path=Path(payload["bib_path"]) if payload.get("bib_path") else None,
            title_path=Path(payload["title_path"]) if payload.get("title_path") else None,
            style_path=Path(payload["style_path"]) if payload.get("style_path") else None,
            csl_path=Path(payload["csl_path"]) if payload.get("csl_path") else None,
            use_toc=payload["use_toc"],
            number_headings=payload["number_headings"],
            output_path=Path(payload["output_path"]),
            lua_filter_path=(
                Path(payload["lua_filter_path"])
                if payload.get("lua_filter_path") else None
            ),
            crossref_config_path=(
                Path(payload["crossref_config_path"])
                if payload.get("crossref_config_path") else None
            ),
            pandoc_args_path=(
                Path(payload["pandoc_args_path"])
                if payload.get("pandoc_args_path") else None
            ),
        )

        pandoc_path = Path(payload["pandoc_path"])
        crossref_path = (
            Path(payload["crossref_path"])
            if payload.get("crossref_path")
            else None
        )
        diagram_timeout = int(payload.get("diagram_timeout", 30))

        work_dir = Path(os.environ["MD2GOST_TMPDIR"])

        orchestrator = PipelineOrchestrator(
            task,
            pandoc_path,
            crossref_path,
            work_dir=work_dir,
            diagram_timeout=diagram_timeout,
        )

        success = orchestrator.execute()
        sys.exit(0 if success else 1)

    except UserFacingError as exc:
        logging.getLogger("core").error("%s", exc)
        sys.exit(1)
    except PermissionError as exc:
        target = exc.filename or "файлу"
        logging.getLogger("core").error(
            "Нет доступа к %s. Возможно, файл открыт в Word или другой программе "
            "либо недостаточно прав — закройте его и повторите конвертацию.",
            f"'{target}'" if exc.filename else target,
        )
        sys.exit(1)
    except Exception:
        logging.getLogger("core").exception("Необработанная ошибка в воркере:")
        sys.exit(1)


if __name__ == "__main__":
    main()

"""
Обработка событий интерфейса: проверка входных данных, запуск конвейера
в отдельном процессе и передача его журнала в консоль через цикл событий Tk.

Конвейер выполняется в отдельном процессе (tools/_pipeline_worker.py),
который вместе со всеми дочерними (pandoc, pandoc-crossref, рендеринг
диаграмм) образует собственную группу процессов. Поэтому отмена мгновенна
и не оставляет «осиротевших» процессов: в Unix вся группа завершается
через os.killpg(SIGKILL), в Windows всё дерево — через taskkill /F /T.
"""
import os
import sys
import json
import time
import shutil
import signal
import platform
import tempfile
import subprocess
import threading
import logging
from pathlib import Path
from typing import TYPE_CHECKING, List, Optional
from tkinter import messagebox, filedialog, TclError
from .components import LogConsole
from .pandoc_setup import PandocSetup
from core.config import ProjectConfig

if TYPE_CHECKING:
    from .app import MainWindow

logger = logging.getLogger(__name__)

_WORKER_PATH = Path(__file__).resolve().parent.parent / "tools" / "_pipeline_worker.py"

# Скрывает консольные окна рабочего процесса (а через наследование — и pandoc,
# pandoc-crossref, рендеринга диаграмм), когда GUI запущен через pythonw.
_NO_WINDOW_FLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0)

# Сколько ждать потоки чтения журнала после завершения рабочего процесса,
# чтобы последние строки (часто — текст ошибки) появились до диалога с итогом.
_READER_DRAIN_TIMEOUT_S = 2.0
# Сколько ждать завершения убитого рабочего процесса при закрытии окна.
_CLOSE_WAIT_TIMEOUT_S = 3.0


def _open_file(path: Path) -> None:
    """
    Открывает *path* программой, назначенной в системе (Windows, macOS, Linux).
    Ошибку сообщает отдельным диалогом: файл к этому моменту уже сохранён,
    и пользователь должен понимать, что не удалось только его открытие.
    """
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(str(path))
        elif system == "Darwin":
            result = subprocess.run(
                ["open", str(path)],
                check=False,
                capture_output=True,
                text=True,
            )
            if result.returncode != 0:
                raise OSError(result.stderr.strip() or "команда open завершилась с ошибкой")
        else:
            result = subprocess.run(
                ["xdg-open", str(path)],
                check=False,
                capture_output=True,
                text=True,
            )
            if result.returncode != 0:
                raise OSError(
                    result.stderr.strip() or "команда xdg-open завершилась с ошибкой"
                )
    except Exception as exc:
        logger.exception("Не удалось открыть файл: %s", path)
        try:
            messagebox.showerror(
                "Не удалось открыть файл",
                f"Документ сохранён, но открыть его автоматически не удалось:\n\n"
                f"{path}\n\nПричина: {exc}",
            )
        except TclError:
            pass


class EventHandler:
    """
    Связующее звено между GUI и конвейером конвертации.

    Одновременно выполняется не больше одной конвертации. Пока она идёт,
    кнопка «Конвертировать» заблокирована, а «Отмена» доступна.

    Журнал рабочего процесса читается двумя фоновыми потоками: из stdout
    (JSON-строки) и из stderr (простой текст). Потоки лишь передают строки
    в LogConsole.append(), который ставит их в очередь, а в виджет их
    вставляет главный поток: из потоков чтения Tk никогда не вызывается.
    """

    def __init__(
        self,
        root: "MainWindow",
        console: LogConsole,
        pandoc_setup: PandocSetup,
        cfg: Optional[ProjectConfig] = None,
    ) -> None:
        self.ui       = root
        self.console  = console
        self._pandoc_setup = pandoc_setup
        self._cfg     = cfg
        self._process: Optional[subprocess.Popen] = None
        self._cancelled: bool = False
        self._closing:   bool = False
        self._worker_tmpdir: Optional[str] = None
        # Путь итогового файла нужен после завершения, чтобы предложить его открыть.
        self._output_path: Optional[Path] = None
        self._readers: List[threading.Thread] = []
        self._reader_deadline: Optional[float] = None

    def on_convert_click(self) -> None:
        """
        Проверяет входные данные, запрашивает путь сохранения, сериализует
        задачу в JSON и передаёт её рабочему процессу через stdin.

        Ничего не делает, если конвертация уже идёт или работает установщик
        pandoc. Если pandoc ещё не признан готовым, вместо конвертации
        запускается его проверка.
        """
        if self._process is not None and self._process.poll() is None:
            return
        if self._pandoc_setup.busy:
            return
        if not self._pandoc_setup.ready:
            self._pandoc_setup.start_check()
            return

        params = self.ui.get_task_parameters()

        if not params["md_path"] or not Path(params["md_path"]).is_file():
            messagebox.showerror("Ошибка", "Выберите существующий Markdown-файл.")
            return

        if params["bib_path"] and not Path(params["bib_path"]).is_file():
            messagebox.showerror(
                "Ошибка",
                f"Выбранный .bib файл не найден:\n{params['bib_path']}",
            )
            return

        out = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=(("Word", "*.docx"),),
            initialfile="result.docx",
        )
        if not out:
            return

        out_path = Path(out)
        if not out_path.parent.exists():
            messagebox.showerror(
                "Ошибка",
                f"Папка для сохранения не существует:\n{out_path.parent}\n\n"
                "Создайте папку и попробуйте снова.",
            )
            return

        cfg = self._cfg if self._cfg is not None else ProjectConfig()
        # Готовность определяет install_pandoc.py --check; здесь программы лишь
        # находятся в том же порядке, что и у установщика: сначала bin/, затем PATH.
        pandoc_path   = cfg.resolve_binary("pandoc")
        crossref_path = cfg.resolve_binary("pandoc-crossref")
        if not pandoc_path or not crossref_path:
            self._pandoc_setup.ready = False
            self._pandoc_setup.start_check()
            return

        title_path = cfg.resolve_template("titles", params["title_name"])
        profile = cfg.resolve_gost_profile(params["style_name"])
        style_path = profile.reference_docx if profile else None
        # CSL — часть требований ГОСТа: берётся только из профиля.
        csl_path   = profile.csl if profile else None

        missing_resources = []
        if params["title_name"] and title_path is None:
            missing_resources.append(f"титульный лист: {params['title_name']}")
        if params["style_name"] and style_path is None:
            missing_resources.append(f"стили тела: {params['style_name']}")
        if missing_resources:
            messagebox.showerror(
                "Ошибка",
                "Выбранные ресурсы не найдены:\n\n"
                + "\n".join(f"  • {item}" for item in missing_resources)
                + "\n\nВыберите доступные файлы и повторите попытку.",
            )
            return

        # pandoc-crossref запускается всегда, даже без профиля: иначе ссылки
        # @fig:/@tbl:/@eq: дойдут до --citeproc, и он примет их за цитаты.
        # YAML-конфиг профиля, если есть, лишь меняет подписи и нумерацию.
        payload_str = json.dumps(
            {
                "md_path":       params["md_path"],
                "bib_path":      params["bib_path"],
                "title_path":    str(title_path)    if title_path    else None,
                "style_path":    str(style_path)    if style_path    else None,
                "csl_path":      str(csl_path)      if csl_path      else None,
                "use_toc":       params["use_toc"],
                "number_headings": params["number_headings"],
                "output_path":   str(out_path),
                "pandoc_path":   str(pandoc_path),
                "crossref_path": str(crossref_path),
                "lua_filter_path": (
                    str(profile.lua_filter) if profile and profile.lua_filter else None
                ),
                "crossref_config_path": (
                    str(profile.crossref_config)
                    if profile and profile.crossref_config else None
                ),
                "pandoc_args_path": (
                    str(profile.pandoc_args) if profile and profile.pandoc_args else None
                ),
                "diagram_timeout": cfg.diagram_timeout,
            },
            ensure_ascii=False,
        )

        self._cancelled    = False
        self._reader_deadline = None
        self._output_path  = out_path
        # Временную папку создаёт и удаляет GUI, а не рабочий процесс: так она
        # будет удалена, даже если процесс убит при отмене.
        self._worker_tmpdir = tempfile.mkdtemp(prefix="md2gost_")
        self.console.clear()
        self.ui.convert_btn.config(state="disabled")
        self.ui.cancel_btn.config(state="normal")

        # PYTHONIOENCODING включает UTF-8 для stdout/stderr дочернего процесса:
        # в Windows каналы по умолчанию используют системную кодовую страницу
        # (например, cp1251), в которой нет многих символов журнала.
        worker_env = {**os.environ, "PYTHONIOENCODING": "utf-8", "MD2GOST_TMPDIR": self._worker_tmpdir}

        # Отдельная группа процессов нужна для отмены одним действием.
        # В Unix это start_new_session; в Windows он не поддерживается, и
        # аналогом служит CREATE_NEW_PROCESS_GROUP.
        if sys.platform == "win32":
            popen_kwargs = {
                "creationflags": subprocess.CREATE_NEW_PROCESS_GROUP | _NO_WINDOW_FLAGS,
            }
        else:
            popen_kwargs = {"start_new_session": True}

        try:
            self._process = subprocess.Popen(
                [sys.executable, str(_WORKER_PATH)],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                env=worker_env,
                **popen_kwargs,
            )
            if self._process.stdin is None:
                raise RuntimeError("Не удалось открыть канал передачи задачи worker.")
            self._process.stdin.write(payload_str)
            self._process.stdin.close()
        except Exception as exc:
            logger.exception("Не удалось запустить worker конвертации")
            process = self._process
            self._process = None
            if process is not None:
                try:
                    process.kill()
                    process.wait(timeout=2)
                except (OSError, subprocess.TimeoutExpired):
                    pass
            self._cleanup_worker_tmpdir()
            self._reset_buttons()
            messagebox.showerror(
                "Ошибка запуска",
                f"Не удалось запустить конвертацию:\n\n{exc}",
            )
            return

        self._readers = [
            threading.Thread(target=self._read_stdout, args=(self._process,), daemon=True),
            threading.Thread(target=self._read_stderr, args=(self._process,), daemon=True),
        ]
        for reader in self._readers:
            reader.start()
        self.ui.after(100, self._poll)

    def request_cancel(self) -> None:
        """
        Немедленно завершает рабочий процесс и все его дочерние процессы.

        Ничего не делает, если конвертация не идёт, в том числе если процесс
        уже завершился сам: так нажатие, совпавшее с обычным окончанием,
        не превратит записанный файл в «Отменено».

        Unix    — os.killpg(SIGKILL) завершает всю группу, созданную
                  start_new_session=True: рабочий процесс, pandoc,
                  pandoc-crossref и рендеринг диаграмм.
        Windows — taskkill /F /T завершает всё дерево процессов от PID
                  рабочего процесса.
        """
        process = self._process
        if process is None or process.poll() is not None:
            return
        self._cancelled = True
        try:
            self.ui.cancel_btn.config(state="disabled")
        except TclError:
            pass
        try:
            if sys.platform == "win32":
                subprocess.run(
                    ["taskkill", "/F", "/T", "/PID", str(process.pid)],
                    check=False,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=_NO_WINDOW_FLAGS,
                )
            else:
                os.killpg(os.getpgid(process.pid), signal.SIGKILL)
        except (ProcessLookupError, OSError):
            pass

    def set_closing(self) -> None:
        """
        Помечает обработчик как закрывающийся, чтобы _poll() и _done()
        не обращались к виджетам уничтоженного окна.
        """
        self._closing = True

    def shutdown(self) -> None:
        """
        Вызывается непосредственно перед уничтожением главного окна.

        После этого цикл событий Tk останавливается, и _poll()/_done() больше
        не выполнятся. Поэтому всё, что они бы очистили, делается здесь
        синхронно: процесс завершается, его выход дожидается, временная
        папка удаляется.
        """
        self.set_closing()
        process = self._process
        if process is None:
            self._cleanup_worker_tmpdir()
            return
        self.request_cancel()
        try:
            process.wait(timeout=_CLOSE_WAIT_TIMEOUT_S)
        except subprocess.TimeoutExpired:
            logger.warning("Процесс конвертации не завершился за %.0f с.", _CLOSE_WAIT_TIMEOUT_S)
        for reader in self._readers:
            reader.join(timeout=0.5)
        self._cleanup_worker_tmpdir()
        self._process = None

    def _read_stdout(self, process: subprocess.Popen) -> None:
        """
        Читает JSON-строки журнала из stdout рабочего процесса и передаёт их
        в консоль. Строка, не являющаяся JSON, выводится как есть с уровнем INFO.
        """
        try:
            for line in process.stdout:
                line = line.rstrip("\n")
                if not line:
                    continue
                try:
                    entry = json.loads(line)
                    msg   = entry.get("msg", line)
                    level = entry.get("level", "INFO")
                except json.JSONDecodeError:
                    msg, level = line, "INFO"
                self.console.append(msg + "\n", level)
        except Exception:
            # Сбой потока чтения не должен ронять GUI, но и не должен теряться
            # бесследно: при уровне журнала DEBUG трассировка будет видна.
            logger.debug("Поток чтения stdout завершился с ошибкой", exc_info=True)

    def _read_stderr(self, process: subprocess.Popen) -> None:
        """Читает простой текст из stderr рабочего процесса и передаёт его в консоль с уровнем WARNING."""
        try:
            for line in process.stderr:
                line = line.rstrip("\n")
                if line:
                    self.console.append(line + "\n", "WARNING")
        except Exception:
            logger.debug("Поток чтения stderr завершился с ошибкой", exc_info=True)

    def _poll(self) -> None:
        """Периодически (через after) проверяет, завершился ли рабочий процесс."""
        if self._process is None or self._closing:
            return
        rc = self._process.poll()
        if rc is None:
            self.ui.after(100, self._poll)
            return

        # Процесс завершился; даём потокам чтения дочитать остаток вывода
        # (EOF приходит сразу после выхода), прежде чем показывать итог.
        if self._reader_deadline is None:
            self._reader_deadline = time.monotonic() + _READER_DRAIN_TIMEOUT_S
        readers_alive = any(reader.is_alive() for reader in self._readers)
        if readers_alive and time.monotonic() < self._reader_deadline:
            self.ui.after(50, self._poll)
            return

        self.console.flush_pending()
        self._done(rc)

    def _done(self, returncode: int) -> None:
        """
        Вызывается в главном потоке после завершения рабочего процесса
        и прочтения его вывода. Восстанавливает интерфейс и сообщает итог.

        Итог определяется кодом возврата: процесс, завершившийся с кодом 0,
        уже записал итоговый файл, поэтому это успех, даже если пользователь
        нажал «Отмена» в самый последний момент.
        """
        ok = returncode == 0
        self._process = None
        self._cleanup_worker_tmpdir()
        if self._closing:
            self._reset_buttons()
            return
        try:
            self._reset_buttons()
            if ok:
                if messagebox.askyesno(
                    "Успех",
                    "Документ успешно создан!\n\nОткрыть файл?",
                ):
                    if self._output_path is not None:
                        _open_file(self._output_path)
            elif self._cancelled:
                messagebox.showinfo("Отменено", "Конвертация была отменена.")
            else:
                messagebox.showerror(
                    "Ошибка",
                    "Конвертация завершилась с ошибкой. Проверьте лог.",
                )
        except TclError:
            pass

    def _reset_buttons(self) -> None:
        """Возвращает кнопки «Конвертировать» и «Отмена» в исходное состояние."""
        try:
            self.ui.convert_btn.config(state="normal")
            self.ui.cancel_btn.config(state="disabled", text="✕ Отмена")
        except TclError:
            pass

    def _cleanup_worker_tmpdir(self) -> None:
        """
        Удаляет временную папку рабочего процесса. Ею владеет GUI (рабочий
        процесс её не удаляет), поэтому метод вызывается после каждой
        конвертации: из _done() в обычном случае и из shutdown() при закрытии
        окна. Отсутствующая или уже удалённая папка пропускается без ошибки.
        """
        path, self._worker_tmpdir = self._worker_tmpdir, None
        if path:
            shutil.rmtree(path, ignore_errors=True)

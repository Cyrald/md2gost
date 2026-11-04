"""
Проверка и установка pandoc и pandoc-crossref из GUI.

Готовность определяет только tools/install_pandoc.py — GUI сам не проверяет
ни версии, ни файлы программ:

    install_pandoc.py --check          код 0 — готово; любой другой — проблема,
                                       которую объясняет вывод самого скрипта
    install_pandoc.py --auto           установка без вопросов
    install_pandoc.py --interactive    интерактивное меню скрипта; ответы
                                       вводятся в консоли журнала GUI

Всё происходит без вопросов пользователю:

* проверка выполняется незаметно, её итог виден только в строке состояния
  (версии, если всё готово);
* при неудачной проверке сразу запускается `--auto`. Вывод установщика
  в консоль не попадает: из него извлекаются этап и процент для строки
  состояния, а полный текст пишется в журнал уровня DEBUG. В консоль
  выводится одна строка в начале установки и одна при успехе;
* если повторная проверка после установки снова неудачна, в консоль
  выводится краткий перечень причин, а строка состояния предлагает повтор.

Ручной режим (--interactive) поддерживается этим модулем, но текущий GUI
его не запускает.

stdout и stderr установщика объединены и читаются «сырыми» блоками, а не
построчно: так приглашение input() (без перевода строки в конце) появляется
в консоли до того, как скрипт остановится в ожидании ответа.
"""
import codecs
import os
import re
import sys
import time
import signal
import subprocess
import threading
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Callable, Dict, List, Optional, Tuple
from tkinter import messagebox
from .components import LogConsole

if TYPE_CHECKING:
    import tkinter as tk
    from .components import PandocStatusBar

logger = logging.getLogger(__name__)

_INSTALLER_PATH = Path(__file__).resolve().parent.parent / "tools" / "install_pandoc.py"
_NO_WINDOW_FLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0)
_POLL_MS = 100
_READER_DRAIN_TIMEOUT_S = 2.0
_CLOSE_WAIT_TIMEOUT_S = 3.0

_MODE_CHECK = "check"
_MODE_AUTO = "auto"
_MODE_MANUAL = "manual"

# Шаблоны строк, которые печатает install_pandoc.py (см. его функции
# info/step/ok/warn/fail). При изменении текста сообщений установщика
# эти выражения нужно обновить, иначе прогресс перестанет отображаться.
_RE_STATE_LINE = re.compile(r"^(pandoc-crossref|pandoc)\s+(\d[^\s]*)\s")
_RE_PLAN_ITEM = re.compile(
    r"^\*\s+установка\s+(pandoc-crossref|pandoc)\s+([^\s,]+).*\((?:~(\d+)\s*МБ|[^)]*)\)\s*$"
)
_RE_DOWNLOAD_STEP = re.compile(r"^==>\s+(pandoc-crossref|pandoc)\s+([^\s:]+).*:\s+загрузка\b")
_RE_PERCENT = re.compile(r"\[[#\-]+\]\s+(\d{1,3})%")
_RE_INSTALLED = re.compile(r"^\[OK\]\s+(pandoc-crossref|pandoc)\s+\S+\s+установлен\b")


def installer_line_level(line: str) -> str:
    """Сопоставляет метки строк установщика ([X], [!]) уровням журнала консоли."""
    stripped = line.lstrip()
    if stripped.startswith("[X]"):
        return "ERROR"
    if stripped.startswith("[!]"):
        return "WARNING"
    return "INFO"


def _problems(output: str) -> List[str]:
    """Тексты строк [X] из вывода установщика без метки, по порядку и без повторов."""
    seen: List[str] = []
    for line in output.splitlines():
        stripped = line.strip()
        if stripped.startswith("[X]"):
            text = stripped[3:].strip()
            if text and text not in seen:
                seen.append(text)
    return seen


def _versions(output: str) -> Dict[str, str]:
    """Версии из блока «Текущее состояние установки» в выводе установщика."""
    found: Dict[str, str] = {}
    for line in output.splitlines():
        match = _RE_STATE_LINE.match(line.strip())
        if match:
            found[match.group(1)] = match.group(2)
    return found


def _versions_text(output: str) -> str:
    """Строка вида «pandoc X · pandoc-crossref Y» для строки состояния."""
    versions = _versions(output)
    parts = [f"{name} {versions[name]}" for name in ("pandoc", "pandoc-crossref") if name in versions]
    return " · ".join(parts)


@dataclass
class _Component:
    """Устанавливаемый компонент: размер из плана установки и доля выполнения (0…1)."""
    name: str
    tag: str
    size_mb: int
    fraction: float = 0.0


class _InstallProgress:
    """
    Превращает потоковый вывод `install_pandoc.py --auto` в описание этапа
    и общий процент для строки состояния.

    feed() вызывается из потока чтения, snapshot() — из главного потока,
    поэтому состояние защищено блокировкой.

    Шкала: 0–5 % — подготовка, 5–92 % — загрузка и установка компонентов
    (пропорционально их размеру), 95 % — завершение, 98 % — повторная проверка.
    """

    _PREP_DONE = 5.0
    _DOWNLOAD_SPAN = 87.0
    _FINISHING = 95.0
    _VERIFYING = 98.0

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._components: List[_Component] = []
        self._current: Optional[_Component] = None
        self._phase = "prep"
        self._detail = "Подготовка к установке"
        self._prep_percent = 1.0
        self._in_plan = False
        self._dirty = True

    def feed(self, line: str) -> None:
        text = line.strip()
        if not text:
            return
        with self._lock:
            self._dirty = self._apply(text) or self._dirty

    def _apply(self, text: str) -> bool:
        """Обновляет состояние по одной строке вывода; True, если отображение нужно перерисовать."""
        # Любой новый заголовок «==>» закрывает блок плана установки.
        if text.startswith("==>"):
            self._in_plan = False

        if text.startswith("==> Текущее состояние") and self._phase == "prep":
            self._detail, self._prep_percent = "Анализ текущей установки", 2.0
            return True
        if text.startswith("==> Получение списка релизов"):
            self._detail, self._prep_percent = "Поиск подходящей версии", 3.0
            return True
        if text.startswith("==> План установки"):
            self._in_plan = True
            self._components = []
            self._detail, self._prep_percent = "Подготовка загрузки", self._PREP_DONE
            return True
        if self._in_plan:
            match = _RE_PLAN_ITEM.match(text)
            if match:
                size = int(match.group(3)) if match.group(3) else 0
                self._components.append(_Component(match.group(1), match.group(2), size))
                return True
            return False

        match = _RE_DOWNLOAD_STEP.match(text)
        if match:
            name, tag = match.group(1), match.group(2)
            component = next((c for c in self._components if c.name == name), None)
            if component is None:
                component = _Component(name, tag, 0)
                self._components.append(component)
            self._current, self._phase = component, "component"
            self._detail = f"Загрузка {name} {tag}"
            return True

        if text.startswith("==> Установка завершена"):
            for component in self._components:
                component.fraction = 1.0
            self._phase, self._current = "finishing", None
            self._detail = "Завершение установки"
            return True

        current = self._current
        if current is None:
            return False

        # Загрузка занимает 85 % доли компонента; остальное приходится
        # на проверку суммы, распаковку и пробный запуск.
        match = _RE_PERCENT.search(text)
        if match:
            current.fraction = max(current.fraction, min(int(match.group(1)), 100) / 100 * 0.85)
            return True
        if text.startswith("[!]") and "Повторная" in text:
            self._detail = f"Повторная загрузка {current.name} {current.tag}"
            return True
        if text.startswith("[OK] Контрольная сумма") or text.startswith("[!]  Контрольная сумма"):
            current.fraction = max(current.fraction, 0.88)
            self._detail = f"Проверка целостности {current.name}"
            return True
        if text.startswith("[OK] Архив"):
            current.fraction = max(current.fraction, 0.92)
            self._detail = f"Распаковка {current.name}"
            return True
        if text.startswith("[OK] Пробный запуск"):
            current.fraction = max(current.fraction, 0.96)
            self._detail = f"Проверка запуска {current.name}"
            return True
        match = _RE_INSTALLED.match(text)
        if match:
            current.fraction = 1.0
            self._detail = f"{current.name} {current.tag} установлен"
            return True
        return False

    def mark_verifying(self) -> None:
        """Переводит прогресс в этап повторной проверки после успешной установки."""
        with self._lock:
            self._phase, self._current = "verifying", None
            self._detail = "Проверка установки"
            self._dirty = True

    def snapshot(self) -> Optional[Tuple[str, float]]:
        """(описание, процент), если с прошлого вызова что-то изменилось, иначе None."""
        with self._lock:
            if not self._dirty:
                return None
            self._dirty = False
            return self._describe(), self._percent()

    def _describe(self) -> str:
        """
        «Этап N из M · …», где этапы — подготовка, каждый компонент и завершение.
        Пока план неизвестен, выводится только описание.
        """
        if not self._components:
            return self._detail
        total = len(self._components) + 2
        if self._phase == "prep":
            index = 1
        elif self._phase == "component" and self._current in self._components:
            index = self._components.index(self._current) + 2
        else:
            index = total
        return f"Этап {index} из {total} · {self._detail}"

    def _percent(self) -> float:
        """Общий процент установки по текущему этапу и долям выполнения компонентов."""
        if self._phase == "verifying":
            return self._VERIFYING
        if self._phase == "finishing":
            return self._FINISHING
        if self._phase == "prep" or not self._components:
            return self._prep_percent
        # Компоненты взвешиваются по размеру загрузки; если размер хотя бы
        # одного неизвестен, все считаются равными.
        sizes = [c.size_mb for c in self._components]
        weights = sizes if all(sizes) else [1] * len(sizes)
        total_weight = float(sum(weights))
        done = sum(w * c.fraction for w, c in zip(weights, self._components))
        return self._PREP_DONE + self._DOWNLOAD_SPAN * done / total_weight


def _kill_process_tree(process: subprocess.Popen) -> None:
    """Принудительно завершает установщик со всеми дочерними процессами (Windows — taskkill /T, Unix — группа процессов)."""
    if process.poll() is not None:
        return
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


def _spawn_installer(args: List[str], with_stdin: bool) -> subprocess.Popen:
    """
    Запускает install_pandoc.py в отдельной группе процессов, чтобы отмена
    завершала и его загрузки. Вывод не буферизуется (-u, PYTHONUNBUFFERED),
    иначе прогресс приходил бы рывками; stdin открыт только в ручном режиме.
    """
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"}
    if sys.platform == "win32":
        group_kwargs = {"creationflags": subprocess.CREATE_NEW_PROCESS_GROUP | _NO_WINDOW_FLAGS}
    else:
        group_kwargs = {"start_new_session": True}
    return subprocess.Popen(
        [sys.executable, "-u", str(_INSTALLER_PATH), *args],
        stdin=subprocess.PIPE if with_stdin else subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        env=env,
        **group_kwargs,
    )


class PandocSetup:
    """
    Запускает установщик подпроцессом, не более одного одновременно.

    on_busy_change(True/False) позволяет окну блокировать конвертацию, пока
    работает установщик, и направлять сюда кнопку «Отмена». status —
    необязательная строка состояния для готовности и прогресса установки.

    Все методы, кроме _read_output() и вызываемого из него _emit(),
    выполняются в главном потоке.
    """

    def __init__(
        self,
        root: "tk.Misc",
        console: LogConsole,
        on_busy_change: Callable[[bool], None],
        status: Optional["PandocStatusBar"] = None,
    ) -> None:
        self._root = root
        self._console = console
        self._on_busy_change = on_busy_change
        self._status = status
        self.ready = False
        self._process: Optional[subprocess.Popen] = None
        self._reader: Optional[threading.Thread] = None
        self._reader_deadline: Optional[float] = None
        self._mode: Optional[str] = None
        self._output: List[str] = []
        self._progress: Optional[_InstallProgress] = None
        self._install_problems: List[str] = []
        self._auto_attempted = False
        self._cancelled = False
        self._closing = False

    @property
    def busy(self) -> bool:
        return self._process is not None

    def start_check(self) -> None:
        """
        Запускает незаметную проверку `install_pandoc.py --check`; при неудаче
        далее автоматически выполняется установка `--auto`.
        """
        if self.busy or self._closing:
            return
        self._auto_attempted = False
        self._install_problems = []
        self._progress = None
        if self._status is not None:
            self._status.show_checking()
        self._run(_MODE_CHECK, ["--check"])

    def cancel(self) -> None:
        """Прерывает работающий установщик; итог обработает _finish() как отмену."""
        process = self._process
        if process is None:
            return
        self._cancelled = True
        _kill_process_tree(process)

    def shutdown(self) -> None:
        """
        Вызывается перед уничтожением главного окна: синхронно завершает
        установщик, потому что после закрытия окна _poll() уже не выполнится.
        """
        self._closing = True
        process = self._process
        if process is None:
            return
        _kill_process_tree(process)
        try:
            process.wait(timeout=_CLOSE_WAIT_TIMEOUT_S)
        except subprocess.TimeoutExpired:
            logger.warning("Установщик pandoc не завершился за %.0f с.", _CLOSE_WAIT_TIMEOUT_S)
        if self._reader is not None:
            self._reader.join(timeout=0.5)
        self._process = None

    def _run(self, mode: str, args: List[str]) -> None:
        """Запускает установщик в режиме mode, поток чтения его вывода и опрос завершения."""
        self._mode = mode
        self._cancelled = False
        self._output = []
        self._reader_deadline = None
        try:
            self._process = _spawn_installer(args, with_stdin=mode == _MODE_MANUAL)
        except OSError as exc:
            self._process = None
            self.ready = False
            logger.exception("Не удалось запустить установщик pandoc")
            if self._status is not None:
                self._status.show_error("Не удалось запустить установщик")
            messagebox.showerror(
                "Ошибка запуска",
                f"Не удалось запустить установщик pandoc:\n{_INSTALLER_PATH}\n\nПричина: {exc}",
            )
            return

        self._on_busy_change(True)
        if mode == _MODE_MANUAL:
            self._console.enable_input(self._send_input)
        self._reader = threading.Thread(
            target=self._read_output, args=(self._process, mode), daemon=True,
        )
        self._reader.start()
        self._root.after(_POLL_MS, self._poll)

    def _read_output(self, process: subprocess.Popen, mode: str) -> None:
        """
        Поток чтения: декодирует блоки вывода и передаёт их построчно.
        Инкрементальный декодер нужен, потому что граница блока может
        прийтись на середину многобайтового символа UTF-8.
        """
        decoder = codecs.getincrementaldecoder("utf-8")(errors="replace")
        pending = ""
        stream = process.stdout
        try:
            while stream is not None:
                chunk = stream.read1(4096)
                if not chunk:
                    break
                *lines, pending = (pending + decoder.decode(chunk)).replace("\r", "").split("\n")
                for line in lines:
                    self._emit(line + "\n", mode)
                # Незавершённый хвост обычно — приглашение input(): показываем его сразу.
                if pending and mode == _MODE_MANUAL:
                    self._emit(pending, mode)
                    pending = ""
        except (OSError, ValueError):
            logger.debug("Поток чтения вывода установщика завершился с ошибкой", exc_info=True)
        tail = pending + decoder.decode(b"", final=True)
        if tail:
            self._emit(tail + "\n", mode)

    def _emit(self, text: str, mode: str) -> None:
        """
        В ручном режиме выводит текст в консоль; в остальных копит его для
        анализа после завершения, а при --auto ещё и обновляет прогресс.
        """
        if mode == _MODE_MANUAL:
            self._console.append(text, installer_line_level(text))
            return
        self._output.append(text)
        progress = self._progress
        if mode == _MODE_AUTO and progress is not None:
            progress.feed(text)

    def _send_input(self, line: str) -> None:
        """Передаёт строку, введённую в консоли, в stdin установщика (ручной режим)."""
        process = self._process
        if process is None or process.stdin is None:
            return
        try:
            process.stdin.write((line + "\n").encode("utf-8"))
            process.stdin.flush()
        except (OSError, ValueError):
            pass

    def _poll(self) -> None:
        """Обновляет прогресс и проверяет завершение установщика; вызывается через after."""
        process = self._process
        if process is None or self._closing:
            return
        self._show_progress()
        returncode = process.poll()
        if returncode is None:
            self._root.after(_POLL_MS, self._poll)
            return
        # Процесс завершился: даём потоку чтения дочитать остаток вывода.
        if self._reader_deadline is None:
            self._reader_deadline = time.monotonic() + _READER_DRAIN_TIMEOUT_S
        if self._reader is not None and self._reader.is_alive() \
                and time.monotonic() < self._reader_deadline:
            self._root.after(50, self._poll)
            return
        self._finish(returncode)

    def _show_progress(self) -> None:
        """Передаёт в строку состояния прогресс установки, если он изменился с прошлого раза."""
        progress = self._progress
        if progress is None or self._status is None:
            return
        update = progress.snapshot()
        if update is not None:
            self._status.show_installing(*update)

    def _finish(self, returncode: int) -> None:
        """
        Освобождает ресурсы завершившегося установщика и решает, что дальше.
        После --check итог разбирает _on_check_done(); после установки
        всегда выполняется повторная --check, так как готовность определяет
        только она.
        """
        process, mode = self._process, self._mode
        self._process = None
        self._reader = None
        for stream in (process.stdin, process.stdout) if process else ():
            if stream is not None:
                try:
                    stream.close()
                except OSError:
                    pass
        if mode == _MODE_MANUAL:
            self._console.disable_input()
        self._on_busy_change(False)
        if self._closing:
            return

        output = "".join(self._output)
        if mode == _MODE_CHECK:
            self._on_check_done(returncode, output)
            return

        self.ready = False
        if self._cancelled:
            self._on_cancelled()
            return
        logger.debug("install_pandoc.py завершился (код %s):\n%s", returncode, output)
        if mode == _MODE_AUTO:
            # Причины неудачи установки сохраняются, чтобы показать их вместе
            # с итогом повторной проверки.
            self._install_problems = _problems(output)
            if returncode == 0 and self._progress is not None:
                self._progress.mark_verifying()
                self._show_progress()
            self._run(_MODE_CHECK, ["--check"])
        else:
            self.start_check()

    def _on_cancelled(self) -> None:
        self._progress = None
        if self._status is not None:
            self._status.show_cancelled()
        self._console.append("Установка pandoc отменена.\n", "WARNING")

    def _on_check_done(self, returncode: int, output: str) -> None:
        """
        Разбирает итог --check: успех — pandoc готов; первая неудача —
        запуск автоматической установки; неудача после установки — отчёт
        об ошибке. Автоустановка пробуется не более одного раза за start_check().
        """
        logger.debug("install_pandoc.py --check (код %s):\n%s", returncode, output)

        if returncode == 0:
            self.ready = True
            versions = _versions_text(output)
            if self._status is not None:
                self._status.show_ready(versions)
            if self._auto_attempted:
                installed = versions.replace(" · ", " и ") or "pandoc и pandoc-crossref"
                self._console.append(f"Установка завершена: {installed} готовы к работе.\n", "INFO")
            self._progress = None
            return

        self.ready = False
        if self._cancelled:
            self._on_cancelled()
            return

        if not self._auto_attempted:
            self._start_auto_install(_problems(output))
            return

        self._progress = None
        self._report_failure(self._install_problems + _problems(output))

    def _start_auto_install(self, problems: List[str]) -> None:
        """Сообщает в консоли о причинах и запускает `install_pandoc.py --auto`."""
        self._auto_attempted = True
        self._progress = _InstallProgress()
        reason = f" ({'; '.join(problems)})" if problems else ""
        self._console.append(
            f"Компоненты pandoc не готовы к работе{reason}. Запущена автоматическая "
            f"установка — ход выполнения отображается в строке состояния внизу окна.\n",
            "INFO",
        )
        self._show_progress()
        self._run(_MODE_AUTO, ["--auto"])

    def _report_failure(self, problems: List[str]) -> None:
        """Сообщает о неудачной установке: причины без повторов и подсказка, как повторить."""
        reasons = list(dict.fromkeys(problems))
        if self._status is not None:
            self._status.show_error()
        self._console.append("Не удалось автоматически установить pandoc.\n", "ERROR")
        for reason in reasons:
            self._console.append(f"  • {reason}\n", "ERROR")
        self._console.append(
            "Проверьте подключение к интернету и нажмите «Повторить» в строке состояния "
            "(или «Конвертировать») — установка будет запущена снова.\n",
            "WARNING",
        )

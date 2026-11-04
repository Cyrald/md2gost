#!/usr/bin/env python3
"""
Установщик pandoc и pandoc-crossref для md2gost.

Загружает с GitHub официальные сборки pandoc (версии 3.9 и выше) и совместимую
с ними сборку pandoc-crossref и размещает их в каталоге <проект>/bin/ вместе
с текстами лицензий. Использует только стандартную библиотеку Python 3.8+.

Запуск:
    python tools/install_pandoc.py              интерактивный режим (без терминала - автоматический)
    python tools/install_pandoc.py --auto       автоматическая установка без подтверждений
    python tools/install_pandoc.py --pandoc 3.10.2
    python tools/install_pandoc.py --upgrade    установка последней совместимой пары версий
    python tools/install_pandoc.py --list       список версий и совместимость с pandoc-crossref
    python tools/install_pandoc.py --check      проверка установки (код возврата 0 - ошибок нет)

Каталог bin/ имеет приоритет при поиске исполняемых файлов
(core/config.py: resolve_binary), поэтому оценивается только тот pandoc,
который фактически будет использоваться md2gost:
  * если существует файл bin/pandoc, используется он, системный pandoc игнорируется;
  * если файла bin/pandoc нет, используется системный pandoc из PATH.

Автоматический режим:
  * если используемый pandoc имеет версию 3.9 и выше и для него существует сборка
    pandoc-crossref, pandoc сохраняется, а в bin/ устанавливается только
    pandoc-crossref, собранный для этой версии;
  * если pandoc отсутствует, не запускается, имеет версию ниже 3.9 или для него
    нет сборки pandoc-crossref, в bin/ устанавливается последняя версия pandoc,
    для которой выпущен pandoc-crossref.

Установка выполняется только в bin/; системный pandoc не изменяется и не удаляется.
"""

from __future__ import annotations

import argparse
import hashlib
import html
import importlib
import json
import os
import platform
import re
import shutil
import ssl
import stat
import subprocess
import sys
import tarfile
import tempfile
import time
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence, Tuple

# Минимальная поддерживаемая версия pandoc.
MIN_PANDOC: Tuple[int, ...] = (3, 9)

PANDOC_REPO = "jgm/pandoc"
CROSSREF_REPO = "lierdakil/pandoc-crossref"
API = "https://api.github.com/repos/{repo}/releases?per_page=100"
ATOM = "https://github.com/{repo}/releases.atom"
DOWNLOAD = "https://github.com/{repo}/releases/download/{tag}/{name}"
RAW = "https://raw.githubusercontent.com/{repo}/{ref}/{path}"

PANDOC_LICENSE_FILES = ("COPYING.md", "COPYRIGHT")
CROSSREF_LICENSE_FILES = ("LICENSE",)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_BIN = PROJECT_ROOT / "bin"
MANIFEST_NAME = "pandoc-install.json"
# Запас свободного места для загрузки, распаковки и установки, МБ.
NEEDED_FREE_MB = 400

IS_WINDOWS = os.name == "nt"
EXE = ".exe" if IS_WINDOWS else ""
USER_AGENT = "md2gost-pandoc-installer/1.0"


# ════════════════════════════════════════════════════════════════════════════
# Вывод в консоль
# ════════════════════════════════════════════════════════════════════════════

def _setup_console() -> None:
    """Переключает потоки вывода на UTF-8 и включает обработку ANSI-последовательностей в Windows."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[union-attr]
        except Exception:
            pass
    if IS_WINDOWS:
        try:
            # Пустая команда включает обработку ANSI-последовательностей в классической консоли conhost.
            os.system("")
        except Exception:
            pass


# Метки «==>», [OK], [!] и [X] разбирает GUI (gui/pandoc_setup.py) для
# отображения прогресса и ошибок; при изменении формата обновите и его.
def info(msg: str = "") -> None:
    print(msg, flush=True)


def step(msg: str) -> None:
    print(f"\n==> {msg}", flush=True)


def ok(msg: str) -> None:
    print(f"  [OK] {msg}", flush=True)


def warn(msg: str) -> None:
    print(f"  [!]  {msg}", flush=True)


def fail(msg: str) -> None:
    print(f"  [X]  {msg}", file=sys.stderr, flush=True)


class InstallError(Exception):
    pass


class Ui:
    """В интерактивном режиме на вопросы отвечает пользователь, иначе берутся ответы по умолчанию."""

    def __init__(self, interactive: bool) -> None:
        self.interactive = interactive

    def confirm(self, question: str, default: bool = True) -> bool:
        if not self.interactive:
            return default
        hint = "[Y/n]" if default else "[y/N]"
        while True:
            answer = self._input(f"{question} {hint}: ").strip().lower()
            if not answer:
                return default
            if answer in ("y", "yes", "д", "да"):
                return True
            if answer in ("n", "no", "н", "нет"):
                return False

    def choose(self, question: str, options: Sequence[str], default: int = 0) -> int:
        if not self.interactive:
            return default
        for i, text in enumerate(options, 1):
            mark = "  <- по умолчанию" if i - 1 == default else ""
            info(f"  [{i}] {text}{mark}")
        while True:
            answer = self._input(f"{question} [1-{len(options)}, Enter = {default + 1}]: ").strip()
            if not answer:
                return default
            if answer.isdigit() and 1 <= int(answer) <= len(options):
                return int(answer) - 1
            warn("Введите номер из списка.")

    def ask(self, question: str) -> str:
        return self._input(question) if self.interactive else ""

    @staticmethod
    def _input(prompt: str) -> str:
        """input(), у которого закрытый stdin означает пустой ответ, то есть вариант по умолчанию."""
        try:
            return input(prompt)
        except EOFError:
            return ""


# ════════════════════════════════════════════════════════════════════════════
# Версии
# ════════════════════════════════════════════════════════════════════════════

def parse_version(text: str) -> Optional[Tuple[int, ...]]:
    """Извлекает первую последовательность вида 1.2.3 из текста; None, если её нет."""
    m = re.search(r"(\d+(?:\.\d+)*)", text or "")
    return tuple(int(p) for p in m.group(1).split(".")) if m else None


def vstr(v: Optional[Tuple[int, ...]]) -> str:
    return ".".join(map(str, v)) if v else "?"


def _pad(v: Tuple[int, ...], n: int) -> Tuple[int, ...]:
    return (tuple(v) + (0,) * n)[:n]


def _strip_zeros(v: Tuple[int, ...]) -> Tuple[int, ...]:
    v = tuple(v)
    while len(v) > 1 and v[-1] == 0:
        v = v[:-1]
    return v


def compat_tier(built_with: Tuple[int, ...], pandoc: Tuple[int, ...]) -> Optional[int]:
    """
    Уровень совместимости pandoc-crossref, собранного для built_with, с pandoc:
    0 — точное совпадение (без учёта завершающих нулей);
    1 — совпадают первые три компонента (3.9 и 3.9.0.2): безопасно;
    2 — совпадают major.minor: работает, но crossref предупреждает о несовпадении;
    None — несовместимы.
    """
    if _strip_zeros(built_with) == _strip_zeros(pandoc):
        return 0
    if _pad(built_with, 3) == _pad(pandoc, 3):
        return 1
    if _pad(built_with, 2) == _pad(pandoc, 2):
        return 2
    return None


TIER_LABEL = {0: "точное совпадение", 1: "совместим", 2: "совместим с предупреждением"}


# ════════════════════════════════════════════════════════════════════════════
# Платформа
# ════════════════════════════════════════════════════════════════════════════

@dataclass(frozen=True)
class Target:
    """Платформа установки и имена соответствующих ей архивов pandoc и pandoc-crossref."""
    os_name: str
    arch: str
    pandoc_suffix: str
    crossref_asset: str

    @property
    def label(self) -> str:
        return f"{self.os_name} {self.arch}"


def detect_target() -> Target:
    """Определяет платформу и соответствующие ей архивы; неподдерживаемая платформа — InstallError."""
    system = platform.system()
    machine = platform.machine().lower()
    is_arm = machine in ("arm64", "aarch64", "armv8", "armv8l")
    is_x64 = machine in ("x86_64", "amd64", "x64")

    if system == "Windows":
        if not (is_x64 or is_arm):
            raise InstallError(f"Архитектура {machine} не поддерживается: сборки pandoc для Windows "
                               f"выпускаются только для 64-разрядных систем.")
        # Сборок для Windows на ARM нет; x64-сборка работает там через эмуляцию.
        return Target("Windows", "x64", "windows-x86_64.zip", "pandoc-crossref-Windows-X64.7z")
    if system == "Darwin":
        # x64-Python под Rosetta сообщает x86_64, но машина — Apple Silicon:
        # ставим нативную arm64-сборку.
        if is_arm or _is_rosetta():
            return Target("macOS", "arm64", "arm64-macOS.zip", "pandoc-crossref-macOS-ARM64.tar.xz")
        return Target("macOS", "x64", "x86_64-macOS.zip", "pandoc-crossref-macOS-X64.tar.xz")
    if system == "Linux":
        if is_arm:
            return Target("Linux", "arm64", "linux-arm64.tar.gz", "pandoc-crossref-Linux-ARM64.tar.xz")
        if is_x64:
            return Target("Linux", "x64", "linux-amd64.tar.gz", "pandoc-crossref-Linux-X64.tar.xz")
        raise InstallError(f"Архитектура {machine} не поддерживается сборками pandoc.")
    raise InstallError(f"Операционная система {system} не поддерживается. "
                       f"Инструкция по ручной установке pandoc: https://pandoc.org/installing.html")


def _is_rosetta() -> bool:
    try:
        out = subprocess.run(["sysctl", "-in", "sysctl.proc_translated"],
                             capture_output=True, text=True, timeout=5)
        return out.stdout.strip() == "1"
    except Exception:
        return False


# ════════════════════════════════════════════════════════════════════════════
# HTTP
# ════════════════════════════════════════════════════════════════════════════

class Http:
    """
    HTTP-клиент на urllib. При ошибке проверки SSL-сертификата переходит
    сначала на хранилище сертификатов certifi (если пакет установлен),
    затем на системный curl. Токен GITHUB_TOKEN/GH_TOKEN, если задан,
    передаётся только в api.github.com — для повышения лимита запросов.
    """

    def __init__(self) -> None:
        self.token = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
        self._ctx: Optional[ssl.SSLContext] = None
        self._use_curl = False
        self._certifi_tried = False

    def _context(self) -> ssl.SSLContext:
        if self._ctx is None:
            self._ctx = ssl.create_default_context()
        return self._ctx

    def _switch_to_fallback_tls(self) -> bool:
        """Вызывается после ошибки проверки сертификата; False, если запасных вариантов не осталось."""
        if self._certifi_tried:
            if shutil.which("curl"):
                warn("Ошибка проверки SSL-сертификата. Загрузка продолжается через системный curl.")
                self._use_curl = True
                return True
            return False
        try:
            import certifi  # type: ignore
            ctx = ssl.create_default_context(cafile=certifi.where())
            self._certifi_tried = True
            self._ctx = ctx
            warn("Ошибка проверки SSL-сертификата. Используется хранилище сертификатов certifi.")
            return True
        except Exception:
            if shutil.which("curl"):
                warn("Ошибка проверки SSL-сертификата. Загрузка продолжается через системный curl.")
                self._use_curl = True
                return True
            return False

    def _headers(self, url: str, accept: Optional[str]) -> Dict[str, str]:
        """Заголовки запроса; токен GitHub добавляется только для api.github.com."""
        headers = {"User-Agent": USER_AGENT}
        if accept:
            headers["Accept"] = accept
        if self.token and "api.github.com" in url:
            headers["Authorization"] = f"Bearer {self.token}"
        return headers

    def get_bytes(self, url: str, accept: Optional[str] = None, retries: int = 3) -> bytes:
        """
        Загружает ресурс целиком. Ошибки клиента (4xx) пробрасываются сразу —
        повтор их не исправит; сетевые и серверные ошибки повторяются
        с нарастающей паузой.
        """
        last: Optional[Exception] = None
        for attempt in range(1, retries + 1):
            if self._use_curl:
                return self._curl(url, None, accept)
            try:
                req = urllib.request.Request(url, headers=self._headers(url, accept))
                with urllib.request.urlopen(req, timeout=30, context=self._context()) as resp:
                    return resp.read()
            except urllib.error.HTTPError as e:
                if e.code in (403, 404, 429) or 400 <= e.code < 500:
                    raise
                last = e
            except Exception as e:
                if _is_cert_error(e) and self._switch_to_fallback_tls():
                    continue
                last = e
            if attempt < retries:
                time.sleep(2 * attempt)
        raise InstallError(f"Не удалось загрузить {url}: {last}")

    def get_text(self, url: str, accept: Optional[str] = None) -> str:
        return self.get_bytes(url, accept).decode("utf-8", errors="replace")

    def download(self, url: str, dest: Path, expected_size: int = 0,
                 retries: int = 4) -> None:
        """Скачивает файл в dest с повторами; 404 — сразу InstallError."""
        last: Optional[Exception] = None
        for attempt in range(1, retries + 1):
            if self._use_curl:
                self._curl(url, dest, None)
                return
            try:
                self._download_once(url, dest, expected_size)
                return
            except urllib.error.HTTPError as e:
                if e.code == 404:
                    raise InstallError(f"Файл не найден на GitHub: {url}") from None
                last = e
            except Exception as e:
                if _is_cert_error(e) and self._switch_to_fallback_tls():
                    continue
                last = e
            if attempt < retries:
                warn(f"Ошибка загрузки ({last}). Повторная попытка {attempt + 1}/{retries} через {3 * attempt} с.")
                time.sleep(3 * attempt)
        raise InstallError(f"Не удалось скачать {url}: {last}")

    def _download_once(self, url: str, dest: Path, expected_size: int) -> None:
        """Одна попытка загрузки с докачкой и проверкой итогового размера файла."""
        # Если от прошлой попытки остался неполный файл, докачиваем его через
        # Range; файл неизвестного или полного размера качаем заново.
        have = dest.stat().st_size if dest.exists() else 0
        headers = self._headers(url, None)
        if have and expected_size and have < expected_size:
            headers["Range"] = f"bytes={have}-"
        elif have:
            dest.unlink()
            have = 0
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=60, context=self._context()) as resp:
            # Сервер без поддержки Range отдаёт файл целиком (не 206) — пишем с начала.
            if have and resp.status != 206:
                have = 0
            total = expected_size or (int(resp.headers.get("Content-Length") or 0) + have)
            mode = "ab" if have else "wb"
            progress = Progress(total, already=have)
            with open(dest, mode) as fh:
                while True:
                    chunk = resp.read(1 << 16)
                    if not chunk:
                        break
                    fh.write(chunk)
                    progress.add(len(chunk))
            progress.done()
        if expected_size and dest.stat().st_size != expected_size:
            raise InstallError(
                f"размер файла ({dest.stat().st_size} байт) не соответствует ожидаемому ({expected_size} байт)"
            )

    def _curl(self, url: str, dest: Optional[Path], accept: Optional[str]) -> bytes:
        """Загрузка через системный curl: в файл dest или, если dest не задан, в память."""
        cmd = ["curl", "-fL", "--retry", "3", "--connect-timeout", "30", "-A", USER_AGENT]
        if accept:
            cmd += ["-H", f"Accept: {accept}"]
        if self.token and "api.github.com" in url:
            cmd += ["-H", f"Authorization: Bearer {self.token}"]
        if dest is not None:
            cmd += ["-o", str(dest), "-#"]
        else:
            cmd += ["-s"]
        cmd.append(url)
        res = subprocess.run(cmd, stdout=subprocess.PIPE if dest is None else None)
        if res.returncode != 0:
            raise InstallError(f"Ошибка загрузки {url} через curl (код возврата {res.returncode}).")
        return res.stdout or b""


def _is_cert_error(exc: Exception) -> bool:
    reason = getattr(exc, "reason", exc)
    return isinstance(reason, ssl.SSLCertVerificationError) or "CERTIFICATE_VERIFY_FAILED" in str(exc)


class Progress:
    """
    Индикатор загрузки. В терминале строка перерисовывается на месте не чаще
    раза в 0,2 с; без терминала (запуск из GUI) новая строка печатается
    только при переходе очередной четверти, чтобы не засорять вывод.
    """

    def __init__(self, total: int, already: int = 0) -> None:
        self.total = total
        self.done_bytes = already
        self.tty = sys.stdout.isatty()
        self.last_print = 0.0
        self.last_pct = -1
        self.start = time.monotonic()

    def add(self, n: int) -> None:
        """Учитывает очередную порцию загруженных байтов и при необходимости обновляет индикатор."""
        self.done_bytes += n
        now = time.monotonic()
        if now - self.last_print < 0.2:
            return
        self.last_print = now
        self._render()

    def _render(self, final: bool = False) -> None:
        """Печатает строку индикатора; final — принудительный вывод в конце загрузки."""
        mb = self.done_bytes / 1048576
        speed = (self.done_bytes / 1048576) / max(time.monotonic() - self.start, 0.001)
        if self.total:
            pct = min(100, int(self.done_bytes * 100 / self.total))
            bar = "#" * (pct // 4) + "-" * (25 - pct // 4)
            line = f"  [{bar}] {pct:3d}%  {mb:6.1f}/{self.total / 1048576:.1f} МБ  {speed:5.1f} МБ/с"
        else:
            pct = 0
            line = f"  {mb:6.1f} МБ  {speed:5.1f} МБ/с"
        if self.tty:
            sys.stdout.write("\r" + line)
            sys.stdout.flush()
        elif final or pct // 25 != self.last_pct // 25:
            print(line, flush=True)
        self.last_pct = pct

    def done(self) -> None:
        self._render(final=True)
        if self.tty:
            sys.stdout.write("\n")


# ════════════════════════════════════════════════════════════════════════════
# Каталог релизов
# ════════════════════════════════════════════════════════════════════════════

@dataclass
class Asset:
    """Файл релиза. size и sha256 известны только при загрузке каталога через GitHub API."""
    name: str
    url: str
    size: int = 0
    sha256: Optional[str] = None


@dataclass
class PandocRelease:
    """Релиз pandoc; assets пуст, если каталог получен из ленты Atom."""
    tag: str
    version: Tuple[int, ...]
    published: str
    assets: Dict[str, Asset] = field(default_factory=dict)

    def asset_for(self, target: Target) -> Asset:
        """
        Архив pandoc для платформы. Если список файлов неизвестен (каталог
        из Atom), ссылка строится по стандартной схеме имён релизов pandoc.
        """
        name = f"pandoc-{self.tag}-{target.pandoc_suffix}"
        return self.assets.get(name) or Asset(name, DOWNLOAD.format(repo=PANDOC_REPO, tag=self.tag, name=name))


@dataclass
class CrossrefRelease:
    """Релиз pandoc-crossref; built_with — версия pandoc, для которой он собран."""
    tag: str
    built_with: Tuple[int, ...]
    published: str
    assets: Dict[str, Asset] = field(default_factory=dict)

    def asset_for(self, target: Target) -> Optional[Asset]:
        """
        Архив pandoc-crossref для платформы. Если список файлов релиза
        известен, а нужного архива в нём нет, возвращает None; если список
        неизвестен (Atom), ссылка строится по стандартному имени.
        """
        if self.assets:
            return self.assets.get(target.crossref_asset)
        name = target.crossref_asset
        return Asset(name, DOWNLOAD.format(repo=CROSSREF_REPO, tag=self.tag, name=name))


class Catalogue:
    """
    Список релизов pandoc и pandoc-crossref с GitHub. Основной источник —
    GitHub API (с размерами и контрольными суммами файлов); если он
    недоступен (например, исчерпан лимит запросов), используется лента Atom.
    """

    def __init__(self, http: Http, target: Target) -> None:
        self.http = http
        self.target = target
        self.pandoc: List[PandocRelease] = []
        self.crossref: List[CrossrefRelease] = []
        self.source = ""

    def load(self) -> None:
        """Загружает каталог релизов, оставляя только pandoc поддерживаемых версий."""
        step("Получение списка релизов с GitHub")
        try:
            self.pandoc = self._pandoc_from_api()
            self.crossref = self._crossref_from_api()
            self.source = "GitHub API"
        except Exception as e:
            reason = _rate_limit_reason(e) or str(e)
            warn(f"GitHub API недоступен ({reason}). Используется лента релизов Atom; "
                 f"проверка контрольных сумм недоступна.")
            self.pandoc = self._pandoc_from_atom()
            self.crossref = self._crossref_from_atom()
            self.source = "лента релизов Atom"
        # pandoc — от новой версии к старой, pandoc-crossref — от нового
        # релиза к старому: на этот порядок рассчитан выбор пары ниже.
        self.pandoc = [r for r in self.pandoc if _pad(r.version, 2) >= MIN_PANDOC]
        self.pandoc.sort(key=lambda r: r.version, reverse=True)
        self.crossref.sort(key=lambda r: r.published, reverse=True)
        if not self.pandoc:
            raise InstallError(f"Релизы pandoc версии {vstr(MIN_PANDOC)} и выше не найдены.")
        ok(f"Найдено версий pandoc {vstr(MIN_PANDOC)}+: {len(self.pandoc)}; "
           f"релизов pandoc-crossref: {len(self.crossref)} (источник: {self.source}).")

    # ---- GitHub API ---------------------------------------------------------
    def _api(self, repo: str) -> list:
        """Опубликованные релизы репозитория (без черновиков и предварительных)."""
        data = json.loads(self.http.get_text(API.format(repo=repo), accept="application/vnd.github+json"))
        if not isinstance(data, list):
            raise InstallError(str(data.get("message", data)))
        return [r for r in data if not r.get("draft") and not r.get("prerelease")]

    @staticmethod
    def _assets(raw: list) -> Dict[str, Asset]:
        """Файлы релиза по имени; SHA-256 берётся из поля digest вида «sha256:…», если оно есть."""
        out = {}
        for a in raw:
            digest = a.get("digest") or ""
            out[a["name"]] = Asset(
                a["name"], a["browser_download_url"], int(a.get("size") or 0),
                digest.split(":", 1)[1] if digest.startswith("sha256:") else None,
            )
        return out

    def _pandoc_from_api(self) -> List[PandocRelease]:
        out = []
        for r in self._api(PANDOC_REPO):
            v = parse_version(r["tag_name"])
            if v:
                out.append(PandocRelease(r["tag_name"], v, r.get("published_at") or "", self._assets(r["assets"])))
        return out

    def _crossref_from_api(self) -> List[CrossrefRelease]:
        """Релизы pandoc-crossref; версия pandoc, для которой собран релиз, берётся из его описания."""
        out = []
        for r in self._api(CROSSREF_REPO):
            built = _built_with(r.get("body") or "")
            if built:
                out.append(CrossrefRelease(r["tag_name"], built, r.get("published_at") or "", self._assets(r["assets"])))
        return out

    # ---- Лента Atom (запасной источник) -------------------------------------
    def _atom(self, repo: str) -> List[Tuple[str, str, str]]:
        """(тег, дата обновления, текст описания) для каждой записи ленты релизов."""
        text = self.http.get_text(ATOM.format(repo=repo))
        entries = []
        for entry in re.findall(r"<entry>(.*?)</entry>", text, re.S):
            tag = re.search(r'/releases/tag/([^"]+)"', entry)
            updated = re.search(r"<updated>([^<]+)</updated>", entry)
            content = re.search(r"<content[^>]*>(.*?)</content>", entry, re.S)
            if tag:
                entries.append((tag.group(1), updated.group(1) if updated else "",
                                html.unescape(content.group(1)) if content else ""))
        return entries

    def _pandoc_from_atom(self) -> List[PandocRelease]:
        out = []
        for tag, updated, _ in self._atom(PANDOC_REPO):
            v = parse_version(tag)
            # Лента не отличает предварительные релизы, поэтому берутся только
            # теги из чисел и точек.
            if v and re.fullmatch(r"[\d.]+", tag):
                out.append(PandocRelease(tag, v, updated))
        return out

    def _crossref_from_atom(self) -> List[CrossrefRelease]:
        out = []
        for tag, updated, body in self._atom(CROSSREF_REPO):
            built = _built_with(body)
            if built:
                out.append(CrossrefRelease(tag, built, updated))
        return out

    # ---- Запросы к каталогу -------------------------------------------------
    def find_pandoc(self, version: Tuple[int, ...]) -> Optional[PandocRelease]:
        """Релиз pandoc с указанной версией (3.9 и 3.9.0 считаются одной версией)."""
        for r in self.pandoc:
            if _strip_zeros(r.version) == _strip_zeros(version):
                return r
        return None

    def best_crossref(self, pandoc: Tuple[int, ...],
                      max_tier: int = 2) -> Optional[Tuple[CrossrefRelease, int]]:
        """
        Подбирает pandoc-crossref для версии pandoc, у которого есть сборка для
        текущей платформы: сначала по лучшему уровню совместимости, а среди
        равных — самый новый релиз (с последними исправлениями).
        """
        best: Optional[Tuple[CrossrefRelease, int]] = None
        # Список отсортирован от новых к старых, поэтому при равном уровне
        # сохраняется первый найденный — самый новый.
        for r in self.crossref:
            tier = compat_tier(r.built_with, pandoc)
            if tier is None or tier > max_tier or r.asset_for(self.target) is None:
                continue
            if best is None or tier < best[1]:
                best = (r, tier)
        return best

    def latest_compatible_pair(self) -> Tuple[PandocRelease, CrossrefRelease, int]:
        """
        Самая новая версия pandoc, для которой есть безопасно совместимый
        (уровень 0–1) pandoc-crossref; только если такой нет — пара
        с уровнем 2 (с предупреждением о несовпадении версий).
        """
        for rel in self.pandoc:
            match = self.best_crossref(rel.version, max_tier=1)
            if match:
                return rel, match[0], match[1]
        for rel in self.pandoc:
            match = self.best_crossref(rel.version, max_tier=2)
            if match:
                return rel, match[0], match[1]
        raise InstallError(f"Для платформы {self.target.label} отсутствуют совместимые версии "
                           f"pandoc {vstr(MIN_PANDOC)}+ и pandoc-crossref.")


def _built_with(body: str) -> Optional[Tuple[int, ...]]:
    """Версия pandoc из фразы «built with Pandoc vX.Y» (описание релиза или вывод --version)."""
    found = re.findall(r"built with Pandoc v([\d.]+)", body, re.I)
    return parse_version(found[0]) if found else None


def _rate_limit_reason(exc: Exception) -> Optional[str]:
    """Понятное пояснение, если GitHub API отказал из-за лимита запросов (403/429)."""
    if isinstance(exc, urllib.error.HTTPError) and exc.code in (403, 429):
        return "превышен лимит запросов; для увеличения лимита задайте переменную окружения GITHUB_TOKEN"
    return None


# ════════════════════════════════════════════════════════════════════════════
# Локальная установка
# ════════════════════════════════════════════════════════════════════════════

@dataclass
class Found:
    """
    Найденная программа: version — её версия; built_with — версия pandoc,
    для которой собран pandoc-crossref; raw — первая строка вывода --version.
    """
    path: Path
    version: Tuple[int, ...]
    location: str  # "bin" или "system"
    built_with: Optional[Tuple[int, ...]] = None
    raw: str = ""


def _run_version(path: Path) -> Optional[str]:
    """Вывод `<программа> --version` или None, если программа не запускается или завершилась с ошибкой."""
    try:
        res = subprocess.run([str(path), "--version"], capture_output=True, text=True,
                             timeout=60, errors="replace")
        return (res.stdout or res.stderr).strip() if res.returncode == 0 else None
    except Exception:
        return None


def probe_pandoc(path: Path, location: str) -> Optional[Found]:
    """Запускает pandoc и определяет его версию; None, если запуск или разбор версии не удался."""
    out = _run_version(path)
    if not out:
        return None
    m = re.match(r"pandoc(?:\.exe)?\s+([\d.]+)", out.splitlines()[0])
    v = parse_version(m.group(1)) if m else None
    return Found(path, v, location, raw=out.splitlines()[0]) if v else None


def probe_crossref(path: Path, location: str) -> Optional[Found]:
    """
    Запускает pandoc-crossref и определяет его версию и версию pandoc,
    для которой он собран. Нераспознанная собственная версия заменяется
    на (0,): для совместимости важна только built_with.
    """
    out = _run_version(path)
    if not out:
        return None
    m = re.search(r"pandoc-crossref\s+v?([\w.]+)", out)
    v = parse_version(m.group(1)) if m else None
    return Found(path, v or (0,), location, built_with=_built_with(out), raw=out.splitlines()[0])


@dataclass
class LocalState:
    """
    Что установлено локально: программы в bin/, в системном PATH и список
    файлов в bin/, которые есть, но не запускаются (broken_bin).
    """
    bin_dir: Path
    bin_pandoc: Optional[Found] = None
    sys_pandoc: Optional[Found] = None
    bin_crossref: Optional[Found] = None
    sys_crossref: Optional[Found] = None
    broken_bin: List[str] = field(default_factory=list)

    def in_bin(self, name: str) -> bool:
        """
        Есть ли файл программы в bin/. Учитываются и неработающие файлы:
        resolve_binary выбирает любой существующий файл в bin/, не проверяя,
        запускается ли он, поэтому приоритет определяет само наличие файла.
        """
        attr = "bin_pandoc" if name == "pandoc" else "bin_crossref"
        return getattr(self, attr) is not None or name in self.broken_bin

    @property
    def pandoc(self) -> Optional[Found]:
        """pandoc, который фактически запустит md2gost; None, если его нет или он не запускается."""
        return self.bin_pandoc if self.in_bin("pandoc") else self.sys_pandoc

    @property
    def crossref(self) -> Optional[Found]:
        """pandoc-crossref, который фактически запустит md2gost; правило выбора то же, что у pandoc."""
        return self.bin_crossref if self.in_bin("pandoc-crossref") else self.sys_crossref


def scan_local(bin_dir: Path, use_system: bool = True) -> LocalState:
    """
    Проверяет программы в bin_dir и (если use_system) в PATH. Если PATH
    указывает на тот же файл, что лежит в bin/, системной копией он
    не считается.
    """
    st = LocalState(bin_dir)
    for name, probe, attr in (("pandoc", probe_pandoc, "pandoc"), ("pandoc-crossref", probe_crossref, "crossref")):
        local = _bin_file(bin_dir, name)
        if local:
            found = probe(local, "bin")
            if found:
                setattr(st, f"bin_{attr}", found)
            else:
                st.broken_bin.append(name)
        if use_system:
            sys_path = shutil.which(name)
            if sys_path and (not local or Path(sys_path).resolve() != local.resolve()):
                setattr(st, f"sys_{attr}", probe(Path(sys_path), "system"))
    return st


def _bin_file(bin_dir: Path, name: str) -> Optional[Path]:
    """Файл программы в bin_dir в том же порядке поиска, что и у core/config.py: resolve_binary."""
    for candidate in (bin_dir / name, bin_dir / f"{name}.exe"):
        if candidate.exists():
            return candidate
    return None


def describe_state(st: LocalState) -> None:
    """
    Печатает блок «Текущее состояние установки» и предупреждения о проблемах.
    Формат строк версий разбирает GUI (gui/pandoc_setup.py: _RE_STATE_LINE).
    """
    step("Текущее состояние установки")
    for label, chosen, other in (("pandoc", st.bin_pandoc, st.sys_pandoc),
                                 ("pandoc-crossref", st.bin_crossref, st.sys_crossref)):
        if chosen:
            extra = f", собран с pandoc {vstr(chosen.built_with)}" if chosen.built_with else ""
            info(f"  {label:<16} {vstr(chosen.version):<10} bin/{extra}")
            if other:
                info(f"  {'':<16} (системная версия {vstr(other.version)} не используется: приоритет у bin/)")
        elif label in st.broken_bin:
            info(f"  {label:<16} {'-':<10} bin/ (не запускается)")
            if other:
                info(f"  {'':<16} (системная версия {vstr(other.version)} не используется: приоритет у bin/)")
        elif other:
            extra = f", собран с pandoc {vstr(other.built_with)}" if other.built_with else ""
            info(f"  {label:<16} {vstr(other.version):<10} системный: {other.path}{extra}")
        else:
            info(f"  {label:<16} не найден")
    for name in st.broken_bin:
        warn(f"bin/{name} не запускается. Файл будет заменён.")
    p, c = st.pandoc, st.crossref
    if p and _pad(p.version, 2) < MIN_PANDOC:
        warn(f"Версия pandoc {vstr(p.version)} ниже минимально поддерживаемой ({vstr(MIN_PANDOC)}). "
             f"Требуется обновление.")
    if p and c and c.built_with:
        tier = compat_tier(c.built_with, p.version)
        if tier is None:
            warn(f"pandoc-crossref собран для pandoc {vstr(c.built_with)} и несовместим "
                 f"с используемой версией {vstr(p.version)}.")
        elif tier == 2:
            warn("pandoc-crossref собран для другой патч-версии pandoc; при конвертации "
                 "возможны предупреждения о несоответствии версий.")


# ════════════════════════════════════════════════════════════════════════════
# Планирование установки
# ════════════════════════════════════════════════════════════════════════════

@dataclass
class Plan:
    """План установки; notes — пояснения, выводимые пользователю перед выполнением."""
    pandoc: Optional[PandocRelease]          # None — оставить текущий pandoc
    pandoc_version: Tuple[int, ...]          # версия pandoc, которая будет использоваться после установки
    crossref: Optional[CrossrefRelease]      # None — оставить текущий pandoc-crossref
    crossref_tier: int
    notes: List[str] = field(default_factory=list)

    @property
    def nothing_to_do(self) -> bool:
        return self.pandoc is None and self.crossref is None


def _crossref_is_fine(st: LocalState, pandoc_v: Tuple[int, ...], wanted_tier: int) -> bool:
    """Подходит ли уже используемый pandoc-crossref для pandoc_v не хуже уровня wanted_tier."""
    c = st.crossref
    if not c or not c.built_with or "pandoc-crossref" in st.broken_bin:
        return False
    tier = compat_tier(c.built_with, pandoc_v)
    return tier is not None and tier <= wanted_tier


def plan_for_version(cat: Catalogue, st: LocalState, rel: PandocRelease, force: bool) -> Plan:
    """
    План для явно выбранной версии pandoc. Уже установленные pandoc этой
    версии и подходящий pandoc-crossref без force не переустанавливаются.
    """
    match = cat.best_crossref(rel.version)
    if not match:
        raise InstallError(
            f"Сборка pandoc-crossref для pandoc {rel.tag} на платформе {cat.target.label} отсутствует. "
            f"Список доступных версий: --list."
        )
    cross, tier = match
    current = st.pandoc
    keep_pandoc = (not force and current is not None
                   and _strip_zeros(current.version) == _strip_zeros(rel.version))
    plan = Plan(None if keep_pandoc else rel, rel.version, cross, tier)
    if keep_pandoc and current is not None:
        where = "bin/" if current.location == "bin" else "системный"
        plan.notes.append(f"pandoc {rel.tag} уже установлен ({where}); переустановка не требуется.")
    if not force and _crossref_is_fine(st, rel.version, tier):
        plan.crossref = None
    if tier == 2:
        plan.notes.append(
            f"Сборка pandoc-crossref для pandoc {rel.tag} отсутствует. Будет установлена ближайшая "
            f"сборка (для pandoc {vstr(cross.built_with)}); при конвертации возможны предупреждения "
            f"о несоответствии версий."
        )
    return plan


def plan_auto(cat: Catalogue, st: LocalState, upgrade: bool, force: bool) -> Plan:
    """
    План автоматического режима (правила — в описании модуля). С upgrade
    или force всегда ставится последняя совместимая пара.
    """
    latest_rel, latest_cross, latest_tier = cat.latest_compatible_pair()
    newest = cat.pandoc[0]
    lag_note = None
    if newest.version != latest_rel.version:
        lag_note = (f"Для последней версии pandoc {newest.tag} сборка pandoc-crossref пока не выпущена. "
                    f"Будет установлена версия {latest_rel.tag} - последняя, для которой доступен pandoc-crossref.")

    def fresh_install(reason: str) -> Plan:
        """План установки последней совместимой пары с пояснением reason."""
        p = Plan(latest_rel, latest_rel.version, latest_cross, latest_tier)
        p.notes.append(reason)
        if lag_note:
            p.notes.append(lag_note)
        return p

    if upgrade or force:
        return fresh_install("Будет установлена последняя совместимая пара pandoc и pandoc-crossref.")

    # Оценивается только тот pandoc, который фактически запустит md2gost: из bin/,
    # если там есть файл pandoc (bin/ приоритетнее PATH), иначе системный.
    # Изменения вносятся только в bin/.
    cur = st.pandoc
    system_note = ""
    if st.sys_pandoc:
        system_note = (f" Системный pandoc {vstr(st.sys_pandoc.version)} не изменяется; "
                       f"будет использоваться версия из bin/, так как bin/ имеет приоритет.")

    if cur is None:
        if st.in_bin("pandoc"):
            return fresh_install("bin/pandoc не запускается и будет заменён." + system_note)
        return fresh_install("pandoc не найден и будет установлен в bin/.")

    where = "в bin/" if cur.location == "bin" else f"в системе ({cur.path})"
    if _pad(cur.version, 2) < MIN_PANDOC:
        target = "pandoc будет обновлён" if cur.location == "bin" else "в bin/ будет установлена поддерживаемая версия"
        return fresh_install(f"Версия pandoc {vstr(cur.version)} {where} ниже минимально поддерживаемой "
                             f"({vstr(MIN_PANDOC)}); {target}." + system_note)

    match = cat.best_crossref(cur.version, max_tier=1)
    if not match:
        return fresh_install(f"Для pandoc {vstr(cur.version)} {where} отсутствует совместимый pandoc-crossref; "
                             f"в bin/ будет установлена совместимая пара." + system_note)

    cross, tier = match
    plan = Plan(None, cur.version, cross, tier)
    plan.notes.append(f"Используемый pandoc {vstr(cur.version)} {where} сохраняется; "
                      f"в bin/ будет установлен соответствующий ему pandoc-crossref.")
    if st.crossref and _crossref_is_fine(st, cur.version, tier):
        plan.crossref = None
        plan.notes.append(f"Установленный pandoc-crossref {vstr(st.crossref.version)} совместим; "
                          f"переустановка не требуется.")
    if newest.version > cur.version:
        plan.notes.append(f"Доступна новая версия pandoc {latest_rel.tag} (обновление: --upgrade).")
    return plan


def show_plan(plan: Plan, cat: Catalogue) -> None:
    """
    Печатает план установки. Строки «* установка …» с размером разбирает GUI
    (gui/pandoc_setup.py: _RE_PLAN_ITEM) для расчёта прогресса.
    """
    step("План установки")
    for note in plan.notes:
        info(f"  - {note}")
    total = 0
    if plan.pandoc:
        a = plan.pandoc.asset_for(cat.target)
        total += a.size
        info(f"  * установка pandoc {plan.pandoc.tag}  ({_mb(a.size)})")
    if plan.crossref:
        ca = plan.crossref.asset_for(cat.target)
        total += ca.size if ca else 0
        info(f"  * установка pandoc-crossref {plan.crossref.tag}, собран с pandoc "
             f"{vstr(plan.crossref.built_with)} - {TIER_LABEL[plan.crossref_tier]}  ({_mb(ca.size if ca else 0)})")
    if total:
        info(f"  Общий объём загрузки: {_mb(total)}")
    if plan.nothing_to_do:
        info("  Все компоненты установлены и совместимы. Действия не требуются.")


def _mb(size: int) -> str:
    return f"~{size / 1048576:.0f} МБ" if size else "раз��ер неизвестен"


# ════════════════════════════════════════════════════════════════════════════
# Распаковка архивов
# ════════════════════════════════════════════════════════════════════════════

def extract(archive: Path, dest: Path, ui: Ui) -> None:
    """
    Распаковывает zip, tar.* или 7z в dest. Перед распаковкой пути внутри
    архива проверяются на выход за пределы dest; из tar извлекаются только
    обычные файлы и каталоги (без ссылок и устройств).
    """
    name = archive.name.lower()
    dest.mkdir(parents=True, exist_ok=True)
    if name.endswith(".zip"):
        with zipfile.ZipFile(archive) as z:
            _safe_members(dest, z.namelist())
            z.extractall(dest)
    elif name.endswith((".tar.gz", ".tgz", ".tar.xz", ".tar.bz2", ".tar")):
        with tarfile.open(archive) as t:
            members = [m for m in t.getmembers() if m.isfile() or m.isdir()]
            _safe_members(dest, [m.name for m in members])
            # Фильтр «data» есть только в новых версиях Python (3.12 и обновления
            # безопасности старых веток); без него защиту даёт _safe_members.
            if hasattr(tarfile, "data_filter"):
                t.extractall(dest, members=members, filter="data")
            else:
                t.extractall(dest, members=members)
    elif name.endswith(".7z"):
        _extract_7z(archive, dest, ui)
    else:
        raise InstallError(f"Неизвестный формат архива: {archive.name}")


def _safe_members(dest: Path, names: List[str]) -> None:
    """Отклоняет архив, если хотя бы один путь в нём ведёт за пределы dest (Zip Slip)."""
    root = dest.resolve()
    for n in names:
        target = (dest / n).resolve()
        if target != root and root not in target.parents:
            raise InstallError(f"Архив содержит небезопасный путь: {n}")


def _extract_7z(archive: Path, dest: Path, ui: Ui) -> None:
    """
    В стандартной библиотеке нет поддержки 7z, поэтому по очереди пробуются
    py7zr, программа 7-Zip и системный tar. Если ничего не подошло, py7zr
    с разрешения пользователя ставится во временный каталог и после
    распаковки выгружается — установленные в системе пакеты не меняются.
    """
    attempts: List[Tuple[str, Callable[[], bool]]] = [
        ("py7zr", lambda: _7z_py7zr(archive, dest)),
        ("7-Zip", lambda: _7z_exe(archive, dest)),
        ("системный tar", lambda: _7z_tar(archive, dest)),
    ]
    for label, attempt in attempts:
        try:
            if attempt() and any(dest.rglob("*")):
                ok(f"Архив .7z распакован ({label}).")
                return
        except Exception:
            pass
    warn("Для распаковки архива .7z требуется модуль py7zr или программа 7-Zip; ни один из них не обнаружен.")
    if not ui.confirm("Загрузить py7zr во временный каталог? Системные пакеты изменены не будут.", True):
        raise InstallError("Распаковка .7z отменена. Установите 7-Zip или выполните: pip install py7zr")
    with tempfile.TemporaryDirectory(prefix="md2gost-py7zr-") as lib:
        res = subprocess.run([sys.executable, "-m", "pip", "install", "--quiet",
                              "--disable-pip-version-check", "--target", lib, "py7zr"])
        # Возможно, pip отсутствует (встраиваемый Python): ставим его через ensurepip и повторяем.
        if res.returncode != 0:
            subprocess.run([sys.executable, "-m", "ensurepip", "--default-pip"], capture_output=True)
            res = subprocess.run([sys.executable, "-m", "pip", "install", "--quiet",
                                  "--disable-pip-version-check", "--target", lib, "py7zr"])
        if res.returncode != 0:
            raise InstallError("Не удалось установить py7zr. Установите 7-Zip (https://7-zip.org) "
                               "и повторите запуск.")
        sys.path.insert(0, lib)
        try:
            importlib.invalidate_caches()
            if not _7z_py7zr(archive, dest):
                raise InstallError("Ошибка распаковки архива с помощью py7zr.")
            ok("Архив .7z распакован (py7zr).")
        finally:
            # Временный каталог сейчас будет удалён, поэтому выгружаем py7zr
            # и его зависимости, загруженные оттуда.
            sys.path.remove(lib)
            for mod in [m for m in sys.modules if m.split(".")[0] in ("py7zr", "pyzstd", "pyppmd", "pybcj", "multivolumefile", "inflate64", "brotli", "texttable", "Cryptodome")]:
                sys.modules.pop(mod, None)


def _7z_py7zr(archive: Path, dest: Path) -> bool:
    """Распаковка через py7zr; False, если модуль не установлен."""
    try:
        import py7zr  # type: ignore
    except ImportError:
        return False
    with py7zr.SevenZipFile(archive, "r") as z:
        _safe_members(dest, z.getnames())
        z.extractall(path=dest)
    return True


def _7z_exe(archive: Path, dest: Path) -> bool:
    """Распаковка программой 7-Zip из PATH или (в Windows) из стандартной папки установки."""
    candidates = [shutil.which(n) for n in ("7z", "7za", "7zr", "7zz")]
    if IS_WINDOWS:
        for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
            if base:
                candidates.append(str(Path(base) / "7-Zip" / "7z.exe"))
    for exe in candidates:
        if exe and Path(exe).exists():
            res = subprocess.run([exe, "x", "-y", f"-o{dest}", str(archive)], capture_output=True)
            if res.returncode == 0:
                return True
    return False


def _7z_tar(archive: Path, dest: Path) -> bool:
    """Распаковка системным tar (bsdtar)."""
    # В Windows 10+ tar.exe — это bsdtar (libarchive), который умеет читать 7z.
    tar = shutil.which("bsdtar") or shutil.which("tar")
    if not tar:
        return False
    res = subprocess.run([tar, "-xf", str(archive), "-C", str(dest)], capture_output=True)
    return res.returncode == 0


def find_binary(root: Path, name: str) -> Path:
    """
    Ищет исполняемый файл в распакованном архиве. При нескольких совпадениях
    берётся самый большой: настоящий исполняемый файл, а не обёртка или ярлык.
    """
    wanted = {name, f"{name}.exe"}
    hits = [p for p in root.rglob("*") if p.is_file() and p.name in wanted]
    if not hits:
        raise InstallError(f"В архиве не найден файл {name}{EXE}.")
    return max(hits, key=lambda p: p.stat().st_size)


# ════════════════════════════════════════════════════════════════════════════
# Установка
# ════════════════════════════════════════════════════════════════════════════

class Installer:
    """
    Выполняет план: скачивает, проверяет и размещает программы в bin/,
    сохраняет тексты лицензий и записывает манифест pandoc-install.json.
    """

    def __init__(self, http: Http, cat: Catalogue, bin_dir: Path, ui: Ui) -> None:
        self.http = http
        self.cat = cat
        self.bin_dir = bin_dir
        self.ui = ui
        self.manifest: dict = self._load_manifest()

    def run(self, plan: Plan) -> None:
        """Выполняет план во временной папке внутри bin/ и записывает манифест установки."""
        self.bin_dir.mkdir(parents=True, exist_ok=True)
        self._check_writable()
        self._check_space(plan)
        # Промежуточная папка внутри bin/ — на том же диске, поэтому
        # os.replace в _place() переносит файл атомарно, без копирования.
        staging_root = self.bin_dir / ".install-tmp"
        staging_root.mkdir(exist_ok=True)
        staging = Path(tempfile.mkdtemp(dir=staging_root))
        try:
            if plan.pandoc:
                self._install_pandoc(plan.pandoc, staging)
            if plan.crossref:
                self._install_crossref(plan.crossref, plan.pandoc_version, staging)
            self.manifest["installed_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
            self.manifest["platform"] = self.cat.target.label
            (self.bin_dir / MANIFEST_NAME).write_text(
                json.dumps(self.manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        finally:
            shutil.rmtree(staging_root, ignore_errors=True)

    # ---- pandoc ------------------------------------------------------------
    def _install_pandoc(self, rel: PandocRelease, staging: Path) -> None:
        """Устанавливает pandoc; пробный запуск должен показать именно ожидаемую версию."""
        asset = rel.asset_for(self.cat.target)
        step(f"pandoc {rel.tag}: загрузка {asset.name}")
        archive = self._fetch(asset, staging)
        unpacked = staging / "pandoc"
        extract(archive, unpacked, self.ui)
        binary = find_binary(unpacked, "pandoc")
        self._verify_runs(binary, "pandoc", lambda f: f is not None and _strip_zeros(f.version) == _strip_zeros(rel.version))
        self._place(binary, "pandoc")
        self._licenses("pandoc", PANDOC_REPO, rel.tag, PANDOC_LICENSE_FILES, unpacked)
        self.manifest["pandoc"] = {"version": rel.tag, "asset": asset.name,
                                   "sha256": asset.sha256, "source": asset.url}
        ok(f"pandoc {rel.tag} установлен в {self.bin_dir}")

    # ---- pandoc-crossref ---------------------------------------------------
    def _install_crossref(self, rel: CrossrefRelease, pandoc_v: Tuple[int, ...], staging: Path) -> None:
        """Устанавливает pandoc-crossref; для пробного запуска достаточно, чтобы он запустился."""
        asset = rel.asset_for(self.cat.target)
        if asset is None:
            raise InstallError(f"В релизе pandoc-crossref {rel.tag} отсутствует сборка для платформы "
                               f"{self.cat.target.label}.")
        step(f"pandoc-crossref {rel.tag} (собран с pandoc {vstr(rel.built_with)}): загрузка {asset.name}")
        archive = self._fetch(asset, staging)
        unpacked = staging / "crossref"
        extract(archive, unpacked, self.ui)
        binary = find_binary(unpacked, "pandoc-crossref")
        self._verify_runs(binary, "pandoc-crossref", lambda f: f is not None)
        self._place(binary, "pandoc-crossref")
        self._licenses("pandoc-crossref", CROSSREF_REPO, rel.tag, CROSSREF_LICENSE_FILES, unpacked)
        self.manifest["pandoc-crossref"] = {"version": rel.tag, "built_with_pandoc": vstr(rel.built_with),
                                            "for_pandoc": vstr(pandoc_v), "asset": asset.name,
                                            "sha256": asset.sha256, "source": asset.url}
        ok(f"pandoc-crossref {rel.tag} установлен в {self.bin_dir}")

    # ---- Общие шаги --------------------------------------------------------
    def _fetch(self, asset: Asset, staging: Path) -> Path:
        """
        Скачивает файл и сверяет SHA-256 с опубликованной на GitHub суммой
        (одна повторная загрузка при несовпадении). Если сумма неизвестна,
        целостность подтверждает только последующий пробный запуск.
        """
        dest = staging / asset.name
        self.http.download(asset.url, dest, asset.size)
        if asset.sha256:
            digest = _sha256(dest)
            if digest != asset.sha256:
                dest.unlink()
                warn("Контрольная сумма не совпадает. Выполняется повторная загрузка.")
                self.http.download(asset.url, dest, asset.size)
                if _sha256(dest) != asset.sha256:
                    raise InstallError(f"Контрольная сумма {asset.name} не совпадает с опубликованной на GitHub.")
            ok("Контрольная сумма SHA-256 совпадает с опубликованной на GitHub.")
        else:
            warn("Контрольная сумма на GitHub не опубликована. Целостность файла проверяется пробным запуском.")
        return dest

    def _verify_runs(self, binary: Path, label: str, accept: Callable[[Optional[Found]], bool]) -> None:
        """Пробный запуск загруженной программы до её переноса в bin/; accept решает, годится ли результат."""
        _make_executable(binary)
        _clear_quarantine(binary)
        probe = probe_pandoc if label == "pandoc" else probe_crossref
        found = probe(binary, "staging")
        if not accept(found):
            raise InstallError(f"Загруженный {label} не запускается на платформе {self.cat.target.label}.")
        ok(f"Пробный запуск выполнен: {found.raw if found else label}")

    def _place(self, binary: Path, name: str) -> None:
        """
        Переносит программу в bin/. Старый файл сначала переименовывается
        в .old и возвращается на место, если перенос не удался, — bin/
        не остаётся без рабочей программы. Если файл занят другим процессом,
        в интерактивном режиме предлагается повторить попытку.
        """
        dest = self.bin_dir / f"{name}{EXE}"
        while True:
            try:
                if dest.exists():
                    old = dest.with_name(dest.name + ".old")
                    if old.exists():
                        old.unlink()
                    os.replace(dest, old)
                    try:
                        os.replace(binary, dest)
                    except Exception:
                        os.replace(old, dest)
                        raise
                    # В Windows старый файл может оставаться заблокированным;
                    # тогда он будет удалён при следующей установке.
                    try:
                        old.unlink()
                    except OSError:
                        pass
                else:
                    os.replace(binary, dest)
                _make_executable(dest)
                return
            except PermissionError:
                fail(f"Файл {dest.name} используется другим процессом. "
                     f"Завершите работу md2gost и активные конвертации.")
                if not self.ui.interactive or not self.ui.confirm(
                        "Нажмите Enter для повторной попытки", True):
                    raise InstallError(f"Не удалось заменить {dest}: файл используется другим процессом.") from None

    def _licenses(self, name: str, repo: str, tag: str, files: Sequence[str], unpacked: Path) -> None:
        """
        Сохраняет тексты лицензий в bin/licenses/<name>/: сначала из репозитория
        (по тегу релиза, затем из master/main), затем недостающие — из архива.
        GPL требует распространять программу вместе с лицензией и указанием
        на исходный код — для этого пишется SOURCE.txt.
        """
        target = self.bin_dir / "licenses" / name
        shutil.rmtree(target, ignore_errors=True)
        target.mkdir(parents=True, exist_ok=True)
        got: List[str] = []
        for fname in files:
            data = None
            for ref in (tag, "master", "main"):
                try:
                    data = self.http.get_bytes(RAW.format(repo=repo, ref=ref, path=fname), retries=2)
                    break
                except Exception:
                    continue
            if data:
                (target / fname).write_bytes(data)
                got.append(fname)
        for local in unpacked.rglob("*"):
            if local.is_file() and re.match(r"(COPYING|COPYRIGHT|LICEN[CS]E)", local.name, re.I) \
                    and local.name not in got:
                shutil.copy2(local, target / local.name)
                got.append(local.name)
        (target / "SOURCE.txt").write_text(
            f"{name} {tag}\nhttps://github.com/{repo}/releases/tag/{tag}\n"
            f"Распространяется на условиях GNU GPL v2 (или более поздней); "
            f"исходный код доступен по ссылке выше.\n", encoding="utf-8")
        # Об успешной загрузке лицензий не сообщаем, предупреждаем только о неудаче.
        for fname in files:
            if fname not in got:
                warn(f"Ошибка загрузки лицензии {name} ({fname}). Текст лицензии доступен по ссылке: "
                     f"https://github.com/{repo}/blob/{tag}/{fname}")

    def _check_writable(self) -> None:
        """Проверяет право записи в bin/ до загрузки, чтобы не скачивать файлы впустую."""
        probe = self.bin_dir / ".write-test"
        try:
            probe.write_text("x")
            probe.unlink()
        except OSError:
            raise InstallError(
                f"Отсутствуют права на запись в каталог {self.bin_dir}. Переместите проект в каталог "
                f"пользователя или укажите другой каталог параметром --bin-dir <путь>."
            ) from None

    def _check_space(self, plan: Plan) -> None:
        """При нехватке места спрашивает, продолжать ли; без интерактива установка прерывается."""
        if not (plan.pandoc or plan.crossref):
            return
        free_mb = shutil.disk_usage(self.bin_dir).free / 1048576
        if free_mb < NEEDED_FREE_MB:
            msg = f"Свободное место на диске: {free_mb:.0f} МБ; для установки требуется около {NEEDED_FREE_MB} МБ."
            if not self.ui.confirm(f"{msg} Продолжить установку?", False):
                raise InstallError(msg)

    def _load_manifest(self) -> dict:
        """Манифест прошлой установки; отсутствующий или повреждённый файл даёт пустой словарь."""
        try:
            return json.loads((self.bin_dir / MANIFEST_NAME).read_text(encoding="utf-8"))
        except Exception:
            return {}


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _make_executable(path: Path) -> None:
    if not IS_WINDOWS:
        path.chmod(path.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)


def _clear_quarantine(path: Path) -> None:
    """В macOS снимает с загруженного файла атрибут карантина, иначе Gatekeeper заблокирует запуск."""
    if platform.system() == "Darwin":
        subprocess.run(["xattr", "-d", "com.apple.quarantine", str(path)], capture_output=True)


# ════════════════════════════════════════════════════════════════════════════
# Команды
# ════════════════════════════════════════════════════════════════════════════

def cmd_list(cat: Catalogue, st: LocalState) -> None:
    """--list: версии pandoc и подходящий для каждой pandoc-crossref."""
    step(f"Доступные версии pandoc >= {vstr(MIN_PANDOC)} для {cat.target.label}")
    current = st.pandoc.version if st.pandoc else None
    info(f"  {'pandoc':<10} {'дата':<11} pandoc-crossref")
    for rel in cat.pandoc:
        match = cat.best_crossref(rel.version)
        cross = (f"{match[0].tag} (собран с {vstr(match[0].built_with)}, {TIER_LABEL[match[1]]})"
                 if match else "сборка отсутствует")
        mark = "  <- установлен" if current and _strip_zeros(current) == _strip_zeros(rel.version) else ""
        info(f"  {rel.tag:<10} {rel.published[:10]:<11} {cross}{mark}")


def cmd_check(st: LocalState) -> int:
    """
    --check: 0, если используемые pandoc и pandoc-crossref запускаются
    и совместимы, иначе 1. Каждая проблема выводится строкой [X] —
    GUI показывает их пользователю как причины.
    """
    p, c = st.pandoc, st.crossref
    problems = []
    if not p:
        problems.append("bin/pandoc не запускается" if st.in_bin("pandoc") else "pandoc не найден")
    elif _pad(p.version, 2) < MIN_PANDOC:
        problems.append(f"версия pandoc {vstr(p.version)} ниже минимально поддерживаемой ({vstr(MIN_PANDOC)})")
    if not c:
        problems.append("bin/pandoc-crossref не запускается" if st.in_bin("pandoc-crossref")
                        else "pandoc-crossref не найден")
    elif p and c.built_with and compat_tier(c.built_with, p.version) is None:
        problems.append(f"pandoc-crossref собран для pandoc {vstr(c.built_with)} и несовместим "
                        f"с используемой версией {vstr(p.version)}")
    if problems:
        for pr in problems:
            fail(pr)
        info("\n  Для исправления выполните: python tools/install_pandoc.py --auto")
        return 1
    ok("pandoc и pandoc-crossref установлены и совместимы.")
    return 0


def pick_version(cat: Catalogue, ui: Ui) -> PandocRelease:
    """
    Интерактивный выбор версии pandoc из 20 последних, для которых есть
    pandoc-crossref. По умолчанию предлагается самая новая безопасно
    совместимая версия.
    """
    step("Выбор версии pandoc")
    usable: List[Tuple[PandocRelease, Tuple[CrossrefRelease, int]]] = []
    for candidate in cat.pandoc:
        match = cat.best_crossref(candidate.version)
        if match:
            usable.append((candidate, match))
    if not usable:
        raise InstallError("Отсутствуют версии pandoc, для которых доступен pandoc-crossref.")
    options = []
    for rel, (cross, tier) in usable[:20]:
        extra = "" if tier <= 1 else ", возможны предупреждения"
        options.append(f"pandoc {rel.tag:<9} {rel.published[:10]}  crossref {cross.tag}{extra}")
    latest_idx = next((i for i, (_, (_, t)) in enumerate(usable[:20]) if t <= 1), 0)
    idx = ui.choose("Версия", options + ["Ввести номер версии вручную"], default=latest_idx)
    if idx < len(options):
        return usable[idx][0]
    while True:
        text = ui.ask("Версия pandoc (например 3.10.2): ").strip()
        v = parse_version(text)
        chosen = cat.find_pandoc(v) if v else None
        if chosen and cat.best_crossref(chosen.version):
            return chosen
        warn(f"Указанная версия не найдена, ниже {vstr(MIN_PANDOC)} или не имеет сборки pandoc-crossref. "
             f"Выберите версию из списка выше.")


def interactive_menu(cat: Catalogue, st: LocalState, ui: Ui, force: bool) -> Optional[Plan]:
    """Главное меню интерактивного режима; None — пользователь выбрал выход."""
    auto = plan_auto(cat, st, upgrade=False, force=False)
    latest_rel, _, _ = cat.latest_compatible_pair()
    auto_text = ("Автоматический режим: " + ("проверка (все компоненты установлены)" if auto.nothing_to_do
                 else "подбор и установка необходимых компонентов"))
    options = [auto_text,
               f"Обновление до последней совместимой версии (pandoc {latest_rel.tag})",
               "Выбор версии pandoc вручную",
               "Список доступных версий",
               "Выход"]
    step("Выберите действие")
    while True:
        choice = ui.choose("Действие", options, default=0)
        if choice == 0:
            return plan_auto(cat, st, upgrade=False, force=force)
        if choice == 1:
            return plan_auto(cat, st, upgrade=True, force=force)
        if choice == 2:
            return plan_for_version(cat, st, pick_version(cat, ui), force)
        if choice == 3:
            cmd_list(cat, st)
            step("Выберите действие")
            continue
        return None


# ════════════════════════════════════════════════════════════════════════════
# Точка входа
# ════════════════════════════════════════════════════════════════════════════

def build_parser() -> argparse.ArgumentParser:
    """Параметры командной строки установщика."""
    p = argparse.ArgumentParser(
        prog="install_pandoc",
        description="Установка pandoc (>= 3.9) и подходящего pandoc-crossref в bin/ проекта.",
    )
    mode = p.add_mutually_exclusive_group()
    mode.add_argument("--auto", "-y", action="store_true", help="автоматическая установка без подтверждений")
    mode.add_argument("--pandoc", metavar="ВЕРСИЯ", help="установка указанной версии pandoc (>= 3.9) или latest")
    mode.add_argument("--upgrade", action="store_true", help="установка последней совместимой пары версий")
    mode.add_argument("--list", action="store_true", help="список версий и сведения о совместимости")
    mode.add_argument("--check", action="store_true", help="проверка установки (код возврата 0 - ошибок нет)")
    p.add_argument("--force", action="store_true", help="переустановка независимо от текущего состояния")
    p.add_argument("--no-system", action="store_true", help="не учитывать pandoc из системного PATH")
    p.add_argument("--bin-dir", type=Path, default=DEFAULT_BIN, help=f"каталог установки (по умолчанию {DEFAULT_BIN})")
    p.add_argument("--interactive", action="store_true",
                   help="задавать вопросы, даже если stdin не терминал (используется GUI md2gost)")
    return p


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Возвращает код завершения: 0 — успех (для --check — ошибок нет), 1 — ошибка."""
    _setup_console()
    args = build_parser().parse_args(argv)
    explicit = args.auto or args.pandoc or args.upgrade or args.list or args.check
    # Вопросы задаются только при наличии терминала (или с --interactive) и никогда — с --auto.
    ui = Ui(interactive=(args.interactive or sys.stdin.isatty()) and not args.auto)
    bin_dir: Path = args.bin_dir.expanduser().resolve()

    info("md2gost - установка pandoc и pandoc-crossref")
    try:
        target = detect_target()
        info(f"  Платформа: {target.label}    Каталог установки: {bin_dir}")
        st = scan_local(bin_dir, use_system=not args.no_system)
        describe_state(st)
        # --check работает без сети: каталог релизов ему не нужен.
        if args.check:
            return cmd_check(st)

        http = Http()
        cat = Catalogue(http, target)
        cat.load()

        if args.list:
            cmd_list(cat, st)
            return 0
        plan: Optional[Plan]
        if args.pandoc:
            if args.pandoc.lower() == "latest":
                plan = plan_auto(cat, st, upgrade=True, force=args.force)
            else:
                v = parse_version(args.pandoc)
                if not v or _pad(v, 2) < MIN_PANDOC:
                    raise InstallError(f"Требуется версия pandoc не ниже {vstr(MIN_PANDOC)} "
                                       f"(указано: {args.pandoc}).")
                rel = cat.find_pandoc(v)
                if not rel:
                    raise InstallError(f"Версия pandoc {args.pandoc} не найдена среди релизов. "
                                       f"Список доступных версий: --list.")
                plan = plan_for_version(cat, st, rel, args.force)
        elif args.upgrade or args.auto or not ui.interactive or explicit:
            plan = plan_auto(cat, st, upgrade=args.upgrade, force=args.force)
        else:
            plan = interactive_menu(cat, st, ui, args.force)
            if plan is None:
                return 0

        show_plan(plan, cat)
        if plan.nothing_to_do:
            return 0
        if not ui.confirm("\nПродолжить?", True):
            info("Установка отменена.")
            return 0

        Installer(http, cat, bin_dir, ui).run(plan)

        final = scan_local(bin_dir, use_system=not args.no_system)
        step("Установка завершена")
        describe_state(final)
        return cmd_check(final)
    except KeyboardInterrupt:
        fail("Операция прервана пользователем.")
        return 130
    except InstallError as e:
        fail(str(e))
        return 1
    # Последний рубеж: трассировка выводится только при MD2GOST_DEBUG=1.
    except Exception as e:
        fail(f"Непредвиденная ошибка: {type(e).__name__}: {e}")
        if os.environ.get("MD2GOST_DEBUG"):
            raise
        info("  Для вывода подробной информации запустите установщик с переменной окружения MD2GOST_DEBUG=1.")
        return 1


def _pause_if_double_clicked() -> None:
    """
    Не даёт окну консоли закрыться сразу при запуске двойным щелчком в Windows
    (признак — запуск в терминале без аргументов), чтобы пользователь успел
    прочитать результат.
    """
    if IS_WINDOWS and sys.stdin.isatty() and len(sys.argv) == 1:
        try:
            input("\nНажмите Enter, чтобы закрыть окно...")
        except EOFError:
            pass


if __name__ == "__main__":
    code = main()
    _pause_if_double_clicked()
    sys.exit(code)

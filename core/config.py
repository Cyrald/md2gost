"""
Пути проекта, поиск внешних программ и обнаружение ресурсов
(титульных листов и профилей ГОСТ со всеми их файлами, включая CSL).
"""
import shutil
import logging
from pathlib import Path, PureWindowsPath
from dataclasses import dataclass, replace
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class TaskConfig:
    """Проверенные входные данные одной задачи конвертации."""
    md_path: Path
    bib_path: Optional[Path]
    title_path: Optional[Path]
    style_path: Optional[Path]
    use_toc: bool
    output_path: Path
    csl_path: Optional[Path] = None
    number_headings: bool = False
    lua_filter_path: Optional[Path] = None
    crossref_config_path: Optional[Path] = None
    pandoc_args_path: Optional[Path] = None


@dataclass(frozen=True)
class GostProfile:
    """Проверенный профиль ГОСТ: папка в resources/styles/ с config.cfg и файлами, на которые он ссылается."""
    name: str
    directory: Path
    config_path: Path
    reference_docx: Path
    lua_filter: Optional[Path] = None
    crossref_config: Optional[Path] = None
    csl: Optional[Path] = None
    pandoc_args: Optional[Path] = None
    gost_generator: Optional[Path] = None
    docs_path: Optional[Path] = None
    prompt_path: Optional[Path] = None


class ProjectConfig:
    """
    Единый источник путей проекта и расположения внешних программ.

    Корень проекта по умолчанию вычисляется от расположения этого файла,
    а не от текущего каталога, поэтому запуск из другой папки ничего не ломает.
    Явный base_dir нужен, когда корень известен заранее.
    """

    def __init__(self, base_dir: Optional[Path] = None) -> None:
        self.base_dir = base_dir or Path(__file__).resolve().parent.parent
        self.resources_dir = self.base_dir / "resources"

        # Предельное время рендеринга одного блока <diagram>, в секундах.
        # Единственное место настройки: увеличьте, если тяжёлые диаграммы
        # прерываются по тайм-ауту.
        self.diagram_timeout: int = 30

        # scan_gost_profiles() выполняется при каждом поиске профиля, поэтому
        # предупреждение о повторяющемся gost_name выводится один раз за сеанс.
        self._warned_duplicate_names: set = set()

    def resolve_binary(self, name: str) -> Optional[Path]:
        """
        Ищет программу по имени: сначала в локальной папке bin/ (с расширением
        .exe и без него), затем в системном PATH.

        bin/ проверяется первой, потому что туда tools/install_pandoc.py ставит
        согласованные между собой версии pandoc и pandoc-crossref; они должны
        иметь приоритет над случайной системной установкой.
        """
        bin_dir = self.base_dir / "bin"
        for candidate in [bin_dir / name, bin_dir / f"{name}.exe"]:
            if candidate.exists():
                return candidate

        system_path = shutil.which(name)
        if system_path:
            return Path(system_path)

        return None

    def scan_templates(self, folder_name: str) -> List[Path]:
        """
        Возвращает отсортированный по имени список .docx из resources/<folder_name>/.
        Если папки нет, возвращает пустой список, чтобы вызывающему коду
        не приходилось обрабатывать этот случай отдельно.
        """
        target = self.resources_dir / folder_name
        if not target.is_dir():
            return []
        return sorted(target.glob("*.docx"))

    def scan_gost_profiles(self) -> List[GostProfile]:
        """
        Возвращает корректные профили ГОСТ из непосредственных подпапок styles/.

        Обязательны: config.cfg, непустой gost_name и существующий reference_docx
        в той же папке. Lua-фильтр, YAML pandoc-crossref, стиль CSL, файл
        дополнительных аргументов pandoc, генератор, документация и промпт
        для ИИ необязательны, но если указаны — разрешаются относительно
        папки профиля и должны вести к существующим файлам внутри resources/.
        Некорректный профиль пропускается с предупреждением и не мешает
        загрузке остальных.
        """
        styles_dir = self.resources_dir / "styles"
        if not styles_dir.is_dir():
            return []

        profiles: List[GostProfile] = []
        for directory in sorted(p for p in styles_dir.iterdir() if p.is_dir()):
            config_path = directory / "config.cfg"
            if not config_path.is_file():
                logger.warning(
                    "Профиль ГОСТа '%s' пропущен: отсутствует config.cfg.",
                    directory.name,
                )
                continue

            try:
                values = self._read_profile_config(config_path)
            except (OSError, UnicodeError) as exc:
                logger.warning(
                    "Профиль ГОСТа '%s' пропущен: не удалось прочитать %s: %s.",
                    directory.name, config_path.name, exc,
                )
                continue

            name = values.get("gost_name", "").strip()
            if not name:
                logger.warning(
                    "Профиль ГОСТа '%s' пропущен: в config.cfg не указан "
                    "обязательный параметр gost_name.",
                    directory.name,
                )
                continue

            reference_value = values.get("reference_docx", "").strip()
            reference_docx = self._resolve_profile_file(
                directory, reference_value, "reference_docx", required=True,
            )
            if reference_docx is None:
                continue

            lua_filter = self._resolve_resource_file(
                self.resources_dir, directory, values.get("lua_filter"), "lua_filter",
            )
            crossref_config = self._resolve_resource_file(
                self.resources_dir, directory, values.get("crossref_config"),
                "crossref_config",
            )
            csl = self._resolve_resource_file(
                self.resources_dir, directory, values.get("csl"), "csl",
            )
            pandoc_args = self._resolve_resource_file(
                self.resources_dir, directory, values.get("pandoc_args"), "pandoc_args",
            )
            gost_generator = self._resolve_resource_file(
                self.resources_dir, directory, values.get("gost_generator"),
                "gost_generator",
            )
            docs_path = self._resolve_resource_file(
                self.resources_dir, directory, values.get("docs"), "docs",
            )
            prompt_path = self._resolve_resource_file(
                self.resources_dir, directory, values.get("prompt"), "prompt",
            )

            # Отсутствие необязательного ключа допустимо, а указанный, но
            # не найденный файл — нет: такой профиль неполон и исключается.
            optional_invalid = (
                ("lua_filter" in values and lua_filter is None)
                or ("crossref_config" in values and crossref_config is None)
                or ("csl" in values and csl is None)
                or ("pandoc_args" in values and pandoc_args is None)
                or ("gost_generator" in values and gost_generator is None)
                or ("docs" in values and docs_path is None)
                or ("prompt" in values and prompt_path is None)
            )
            if optional_invalid:
                logger.warning(
                    "Профиль ГОСТа '%s' пропущен: config.cfg содержит "
                    "отсутствующий или недопустимый необязательный файл.",
                    directory.name,
                )
                continue

            profiles.append(
                GostProfile(
                    name=name,
                    directory=directory,
                    config_path=config_path,
                    reference_docx=reference_docx,
                    lua_filter=lua_filter,
                    crossref_config=crossref_config,
                    csl=csl,
                    pandoc_args=pandoc_args,
                    gost_generator=gost_generator,
                    docs_path=docs_path,
                    prompt_path=prompt_path,
                )
            )

        return self._disambiguate_profile_names(profiles)

    def _disambiguate_profile_names(self, profiles: List[GostProfile]) -> List[GostProfile]:
        """
        Делает отображаемые имена профилей уникальными.

        Если несколько профилей объявляют один gost_name, первый (в порядке
        папок) сохраняет его, а следующие получают суффиксы «(1)», «(2)», …
        Уже занятые суффиксы пропускаются: «ГОСТ», «ГОСТ», «ГОСТ(1)» дают
        «ГОСТ», «ГОСТ(2)», «ГОСТ(1)». Порядок стабилен, поэтому выбранное в GUI
        имя всегда указывает на тот же профиль. Предупреждение выводится один
        раз на каждое повторяющееся имя.
        """
        by_name: Dict[str, List[int]] = {}
        for index, profile in enumerate(profiles):
            by_name.setdefault(profile.name, []).append(index)

        taken = {profile.name for profile in profiles}
        result = list(profiles)
        for name, indices in by_name.items():
            if len(indices) < 2:
                continue
            renamed: List[str] = []
            counter = 0
            for index in indices[1:]:
                counter += 1
                candidate = f"{name}({counter})"
                while candidate in taken:
                    counter += 1
                    candidate = f"{name}({counter})"
                taken.add(candidate)
                result[index] = replace(profiles[index], name=candidate)
                renamed.append(f"'{profiles[index].directory.name}' → '{candidate}'")
            if name not in self._warned_duplicate_names:
                self._warned_duplicate_names.add(name)
                logger.warning(
                    "Несколько профилей ГОСТа имеют одинаковое название gost_name='%s'. "
                    "Профиль '%s' сохранил название, остальные переименованы: %s. "
                    "Исправьте gost_name в их config.cfg.",
                    name, profiles[indices[0]].directory.name, ", ".join(renamed),
                )
        return result

    def resolve_gost_profile(self, display_name: Optional[str]) -> Optional[GostProfile]:
        """Находит профиль по отображаемому в GUI имени; None, если имя пустое или не найдено."""
        if not display_name:
            return None
        for profile in self.scan_gost_profiles():
            if profile.name == display_name:
                return profile
        logger.warning("Профиль ГОСТа '%s' не найден.", display_name)
        return None

    @staticmethod
    def _read_profile_config(path: Path) -> Dict[str, str]:
        """
        Читает config.cfg в формате «ключ=значение».

        Пустые строки, комментарии (# и ;), строки без «=» и неизвестные ключи
        пропускаются намеренно: так старая версия программы не ломается
        на config.cfg из более новой.
        """
        values: Dict[str, str] = {}
        for raw_line in path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith(("#", ";")) or "=" not in line:
                continue
            key, value = (part.strip() for part in line.split("=", 1))
            if key == "promt":
                # Ключ с опечаткой из прежних версий; если указаны оба, побеждает «prompt».
                logger.warning(
                    "%s: ключ 'promt' устарел, переименуйте его в 'prompt'.", path,
                )
                values.setdefault("prompt", value)
                continue
            if key in {
                "gost_name", "reference_docx", "lua_filter",
                "crossref_config", "csl", "pandoc_args",
                "gost_generator", "docs", "prompt",
            }:
                values[key] = value
        return values

    @staticmethod
    def _resolve_profile_file(
        directory: Path,
        value: Optional[str],
        key: str,
        required: bool = False,
    ) -> Optional[Path]:
        """
        Разрешает имя файла профиля. Допускается только простое имя файла,
        лежащего в папке профиля: пути, «..» и выход за пределы папки отклоняются.
        """
        if value is None or not value.strip():
            if required:
                logger.warning(
                    "Профиль ГОСТа '%s' некорректен: не указан %s.",
                    directory.name, key,
                )
            return None

        filename = value.strip()
        candidate_name = Path(filename)
        if (
            candidate_name.is_absolute()
            or "/" in filename
            or "\\" in filename
            or ".." in candidate_name.parts
            or candidate_name.name != filename
        ):
            logger.warning(
                "Профиль ГОСТа '%s' некорректен: %s должен быть именем "
                "файла в той же папке.",
                directory.name, key,
            )
            return None

        candidate = directory / candidate_name
        try:
            candidate.resolve().relative_to(directory.resolve())
        except ValueError:
            logger.warning(
                "Профиль ГОСТа '%s' некорректен: %s выходит за пределы папки.",
                directory.name, key,
            )
            return None

        if not candidate.is_file():
            logger.warning(
                "Профиль ГОСТа '%s' некорректен: файл %s '%s' не найден.",
                directory.name, key, filename,
            )
            return None
        return candidate

    @staticmethod
    def _resolve_resource_file(
        resources_dir: Path,
        directory: Path,
        value: Optional[str],
        key: str,
    ) -> Optional[Path]:
        """
        Разрешает необязательный путь из профиля относительно его папки.
        Допускаются только относительные POSIX-пути, итог которых остаётся
        внутри resources/; абсолютные пути, диски и «\\» отклоняются.
        """
        if value is None or not value.strip():
            return None

        filename = value.strip()
        relative_path = Path(filename)
        windows_path = PureWindowsPath(filename)
        if (
            relative_path.is_absolute()
            or windows_path.is_absolute()
            or windows_path.drive
            or "\\" in filename
        ):
            logger.warning(
                "Профиль ГОСТа '%s' некорректен: %s должен быть относительным "
                "путём внутри resources/.",
                directory.name, key,
            )
            return None

        resources_root = resources_dir.resolve()
        candidate = (directory / relative_path).resolve()
        try:
            candidate.relative_to(resources_root)
        except ValueError:
            logger.warning(
                "Профиль ГОСТа '%s' некорректен: путь %s выходит за пределы "
                "resources/.",
                directory.name, key,
            )
            return None

        if not candidate.is_file():
            logger.warning(
                "Профиль ГОСТа '%s' некорректен: файл %s '%s' не найден.",
                directory.name, key, filename,
            )
            return None
        return candidate

    def resolve_template(self, folder_name: str, display_name: Optional[str]) -> Optional[Path]:
        """
        Находит шаблон по отображаемому имени (с расширением или без).
        Возвращает None, если имя пустое (пользователь выбрал «не использовать»)
        или файл не найден (во втором случае с предупреждением).
        """
        if not display_name:
            return None

        templates = self.scan_templates(folder_name)
        for tpl in templates:
            if tpl.name == display_name or tpl.stem == display_name:
                return tpl

        logger.warning("Шаблон '%s' не найден в %s/", display_name, folder_name)
        return None

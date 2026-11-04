"""
Присоединение титульного листа к телу документа на уровне пакета OOXML.

Оба файла .docx разбираются как ZIP-архивы с XML-частями. Тело документа
остаётся основой результата, а содержимое титульного листа вместе с его
стилями, нумерацией, сносками, колонтитулами и медиафайлами переносится
в начало документа как изолированный блок (см. attach_title_page).
"""
import contextlib
import copy
import os
import posixpath
import re
import uuid
import logging
import zipfile
from pathlib import Path
from collections.abc import Iterable
from lxml import etree as ET

logger = logging.getLogger(__name__)

_RELS_PATH       = "word/_rels/document.xml.rels"
_CT_PATH         = "[Content_Types].xml"
_STYLES_PATH     = "word/styles.xml"
_DOC_PATH        = "word/document.xml"
_NUMBERING_PATH  = "word/numbering.xml"
_FONT_TABLE_PATH = "word/fontTable.xml"
_FOOTNOTES_PATH  = "word/footnotes.xml"
_ENDNOTES_PATH   = "word/endnotes.xml"
_FOOTNOTES_RELS_PATH = "word/_rels/footnotes.xml.rels"
_ENDNOTES_RELS_PATH  = "word/_rels/endnotes.xml.rels"

_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS  = "http://schemas.openxmlformats.org/package/2006/content-types"
_WNS    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R_NS   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_WP_NS  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_A_NS   = "http://schemas.openxmlformats.org/drawingml/2006/main"

_NUMBERING_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
)
_NUMBERING_CT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
)
_FOOTNOTES_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
)
_ENDNOTES_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
)
_FONT_TABLE_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
)
_FOOTNOTES_CT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
)
_ENDNOTES_CT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
)
_FONT_TABLE_CT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
)

_HEADER_REL_SUFFIX = "/header"
_FOOTER_REL_SUFFIX = "/footer"
_HEADER_CT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
)
_FOOTER_CT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
)

# Части титульного листа, которые не копируются как есть: у результата они
# общие и берутся из тела документа (стили, нумерация и сноски титула
# переносятся отдельно, со слиянием).
_SKIPPABLE_REL_SUFFIXES = frozenset({
    "/styles",
    "/settings",
    "/webSettings",
    "/theme",
    "/fontTable",
    "/numbering",
    "/footnotes",
    "/endnotes",
    "/glossaryDocument",
    "/printerSettings",
    "/stylesWithEffects",
})

# Примечания Word из титульного листа не переносятся: их части удаляются
# вместе со всеми ссылающимися на них якорями.
_COMMENT_REL_SUFFIXES = frozenset({
    "/comments",
    "/commentsExtended",
    "/commentsIds",
    "/commentsExtensible",
    "/people",
})
_COMMENT_ANCHOR_TAGS = frozenset(
    f"{{{_WNS}}}{local}"
    for local in ("commentRangeStart", "commentRangeEnd", "commentReference")
)
# Контейнеры, удаляемые целиком, если что-то внутри ссылается на связь,
# которая не переносится (диаграмма, OLE-объект, SmartArt и т. п.).
_DROPPED_REF_OWNER_TAGS = frozenset(
    f"{{{_WNS}}}{local}" for local in ("drawing", "pict", "object")
)
# Элементы, которые сами являются ссылкой и удаляются непосредственно.
_DROPPED_REF_SELF_TAGS = frozenset(
    f"{{{_WNS}}}{local}" for local in ("headerReference", "footerReference")
)

_IMAGE_EXTENSIONS = {
    "png", "jpg", "jpeg", "gif", "bmp", "svg", "emf", "wmf", "tiff", "tif",
}

# MIME-типы медиафайлов на случай, если в [Content_Types].xml титула их нет.
_MIME_FALLBACKS: dict[str, str] = {
    "png":  "image/png",
    "jpg":  "image/jpeg",
    "jpeg": "image/jpeg",
    "gif":  "image/gif",
    "bmp":  "image/bmp",
    "svg":  "image/svg+xml",
    "emf":  "image/x-emf",
    "wmf":  "image/x-wmf",
    "tiff": "image/tiff",
    "tif":  "image/tiff",
    "mp3":  "audio/mpeg",
    "mp4":  "video/mp4",
    "m4a":  "audio/mp4",
    "m4v":  "video/x-m4v",
    "wav":  "audio/x-wav",
    "wma":  "audio/x-ms-wma",
    "wmv":  "video/x-ms-wmv",
    "avi":  "video/x-msvideo",
    "mov":  "video/quicktime",
    "webm": "video/webm",
    "ogg":  "audio/ogg",
    "oga":  "audio/ogg",
    "ogv":  "video/ogg",
    "aac":  "audio/aac",
    "flac": "audio/flac",
}

_MEDIA_EXTENSIONS = _IMAGE_EXTENSIONS | set(_MIME_FALLBACKS)

_GENERIC_MEDIA_CT_TYPE = "application/octet-stream"

# Теги, у которых w:val содержит id стиля.
_STYLE_REF_TAGS = frozenset({
    f"{{{_WNS}}}pStyle",
    f"{{{_WNS}}}rStyle",
    f"{{{_WNS}}}tblStyle",
    f"{{{_WNS}}}basedOn",
    f"{{{_WNS}}}next",
    f"{{{_WNS}}}link",
    f"{{{_WNS}}}numStyleLink",
    f"{{{_WNS}}}styleLink",
})

# Ссылки на необъявленный стиль, которые просто удаляются; pStyle обрабатывается
# отдельно — заменяется стилем абзаца по умолчанию.
_UNDECLARED_STYLE_DROP_TAGS = frozenset({
    f"{{{_WNS}}}rStyle",
    f"{{{_WNS}}}tblStyle",
    f"{{{_WNS}}}basedOn",
    f"{{{_WNS}}}next",
    f"{{{_WNS}}}link",
    f"{{{_WNS}}}numStyleLink",
    f"{{{_WNS}}}styleLink",
})

# Элементы с id объекта рисования, уникальным в пределах документа.
_DRAWING_ID_TAGS = frozenset({
    f"{{{_WP_NS}}}docPr",
    f"{{{_WP_NS}}}cNvPr",
    f"{{{_A_NS}}}cNvPr",
})

# Закладки и исправления: их w:id должны быть уникальны в документе.
_BOOKMARK_ID_TAGS = frozenset({
    f"{{{_WNS}}}bookmarkStart",
    f"{{{_WNS}}}bookmarkEnd",
    f"{{{_WNS}}}ins",
    f"{{{_WNS}}}del",
    f"{{{_WNS}}}moveFrom",
    f"{{{_WNS}}}moveTo",
})

_BOOKMARK_START_TAG = f"{{{_WNS}}}bookmarkStart"
_HYPERLINK_TAG       = f"{{{_WNS}}}hyperlink"
_INSTR_TEXT_TAG      = f"{{{_WNS}}}instrText"
_FLD_SIMPLE_TAG      = f"{{{_WNS}}}fldSimple"



# Безопасный парсер: внешние сущности, DTD и сеть отключены (защита от XXE);
# huge_tree снимает ограничения lxml для очень больших документов.
_XML_PARSER = ET.XMLParser(
    resolve_entities=False,
    no_network=True,
    load_dtd=False,
    dtd_validation=False,
    huge_tree=True,
)


def _fromstring(data: bytes) -> ET._Element:
    """Разбирает XML безопасным парсером; синтаксическую ошибку превращает в ValueError с понятным сообщением."""
    try:
        return ET.fromstring(data, parser=_XML_PARSER)
    except ET.XMLSyntaxError as exc:
        raise ValueError(f"Повреждённое XML-содержимое: {exc}") from exc



def _xml_bytes(root: ET._Element) -> bytes:
    """Сериализует элемент в UTF-8 с XML-декларацией."""
    return (
        b'<?xml version="1.0" encoding="UTF-8"?>\n'
        + ET.tostring(root, encoding="unicode").encode("utf-8")
    )



def _validate_input_file(path: Path) -> None:
    """Проверяет, что входной путь — читаемый файл и ZIP-архив (.docx)."""
    if not path.exists():
        raise ValueError(f"Входной файл не найден: '{path}'.")
    if not path.is_file():
        raise ValueError(f"Входной путь не является файлом: '{path}'.")
    if not os.access(path, os.R_OK):
        raise ValueError(f"Нет прав на чтение входного файла '{path}'.")
    if not zipfile.is_zipfile(path):
        raise ValueError(
            f"'{path.name}' не является корректным файлом .docx (это не ZIP-архив). "
            "Возможно, файл повреждён."
        )


def _validate_output_path(path: Path) -> None:
    """Проверяет, что папка назначения существует и доступна для записи, а сам путь не указывает на папку."""
    parent = path.parent
    if not parent.exists():
        raise ValueError(f"Папка для сохранения '{parent}' не существует.")
    if not parent.is_dir():
        raise ValueError(f"Родительский путь '{parent}' для выходного файла не является папкой.")
    if not os.access(parent, os.W_OK):
        raise ValueError(f"Нет прав на запись в папку '{parent}'.")
    if path.is_dir():
        raise ValueError(
            f"Выходной путь '{path}' указывает на папку, а не на файл."
        )


# Допустимые значения w:type разрыва раздела.
_VALID_SECTION_BREAKS = frozenset({
    "nextPage", "continuous", "evenPage", "oddPage", "nextColumn",
})


def _validate_section_break(value: str) -> None:
    """Проверяет, что тип разрыва раздела — одно из значений w:type, допустимых в OOXML."""
    if not isinstance(value, str) or value not in _VALID_SECTION_BREAKS:
        valid = ", ".join(sorted(_VALID_SECTION_BREAKS))
        raise ValueError(
            f"Недопустимое значение section_break {value!r}. Допустимые значения: {valid}."
        )


# Префикс id стилей титула; отображаемые имена получают «title: ».
TITLE_STYLE_PREFIX = "_title_"


def attach_title_page(
    title: Path,
    document: Path,
    output: Path,
    section_break: str = "nextPage",
) -> None:
    """
    Добавляет ``title`` перед ``document`` и записывает результат в ``output``.

    Основой остаётся документ: его настройки, тема, docDefaults и имена
    встроенных стилей (Heading1, TOC1, Hyperlink, …) сохраняются как есть,
    поэтому поля вроде оглавления по-прежнему опираются на стили документа.
    Титульный лист импортируется изолированным самодостаточным блоком:
    его стили переименовываются с префиксом ``TITLE_STYLE_PREFIX``,
    перепривязываются к docDefaults документа так, чтобы выглядеть точно
    как в исходном файле титула, и лишаются уровней структуры — ни один
    абзац титула не попадёт в оглавление и в область навигации.
    Титул отделяется от тела разрывом раздела типа ``section_break``.
    """
    try:
        title_path = Path(title)
        doc_path   = Path(document)
        out        = Path(output)
    except TypeError as exc:
        raise ValueError(f"Некорректный путь: {exc}") from exc

    _validate_input_file(title_path)
    _validate_input_file(doc_path)
    _validate_output_path(out)
    _validate_section_break(section_break)

    out_resolved = out.resolve()
    if out_resolved in (title_path.resolve(), doc_path.resolve()):
        raise ValueError(
            f"Выходной файл '{out}' не должен совпадать ни с одним из входных документов."
        )
    if title_path.resolve() == doc_path.resolve():
        raise ValueError("Титульный лист и документ должны быть разными файлами.")

    _full_merge(
        base_path=doc_path,
        src_path=title_path,
        output_path=out,
        src_prefix=TITLE_STYLE_PREFIX,
        break_type=section_break,
    )



def _read_if_present(zf: zipfile.ZipFile, names: set[str], path: str) -> bytes | None:
    """Содержимое части архива или None, если её нет."""
    return zf.read(path) if path in names else None


def _full_merge(
    base_path: Path,
    src_path: Path,
    output_path: Path,
    src_prefix: str,
    break_type: str,
) -> None:
    """
    Собирает итоговый пакет: base_path — тело документа (основа),
    src_path — присоединяемый титульный лист.

    Порядок работы:
      1. чтение нужных частей обоих архивов;
      2. новые уникальные имена для медиафайлов и колонтитулов титула
         и перенумерация его связей (rId);
      3. очистка титула: неподдерживаемые объекты, примечания, ссылки
         на тему и уровни структуры;
      4. переименование стилей титула с префиксом src_prefix и явное
         назначение его стилей по умолчанию абзацам, фрагментам и таблицам;
      5. перенумерация id рисунков, закладок, списков и сносок, чтобы они
         не пересекались с документом;
      6. разрыв раздела после титула;
      7. сборка styles.xml, numbering.xml, fontTable.xml, сносок,
         document.xml, .rels и [Content_Types].xml;
      8. запись архива.
    """
    with zipfile.ZipFile(base_path, "r") as bzf:
        b_names = set(bzf.namelist())

        if _DOC_PATH not in b_names:
            raise ValueError(
                f"В '{base_path.name}' отсутствует {_DOC_PATH}. "
                "Возможно, файл повреждён."
            )
        if _CT_PATH not in b_names:
            raise ValueError(
                f"В '{base_path.name}' отсутствует {_CT_PATH}. "
                "Возможно, файл повреждён."
            )

        b_doc_xml        = bzf.read(_DOC_PATH)
        b_styles_xml     = _read_if_present(bzf, b_names, _STYLES_PATH)
        if b_styles_xml is None:
            raise ValueError(
                f"В '{base_path.name}' отсутствует word/styles.xml. "
                "Возможно, файл повреждён."
            )
        b_rels_xml       = _read_if_present(bzf, b_names, _RELS_PATH)
        b_ct_xml         = bzf.read(_CT_PATH)
        b_numbering_xml  = _read_if_present(bzf, b_names, _NUMBERING_PATH)
        b_font_xml       = _read_if_present(bzf, b_names, _FONT_TABLE_PATH)
        b_footnotes_xml  = _read_if_present(bzf, b_names, _FOOTNOTES_PATH)
        b_endnotes_xml   = _read_if_present(bzf, b_names, _ENDNOTES_PATH)
        b_footnotes_rels_xml = _read_if_present(bzf, b_names, _FOOTNOTES_RELS_PATH)
        b_endnotes_rels_xml  = _read_if_present(bzf, b_names, _ENDNOTES_RELS_PATH)

    with zipfile.ZipFile(src_path, "r") as szf:
        s_names = set(szf.namelist())

        if _DOC_PATH not in s_names:
            raise ValueError(
                f"В '{src_path.name}' отсутствует {_DOC_PATH}. "
                "Возможно, файл повреждён."
            )

        s_doc_xml        = szf.read(_DOC_PATH)
        s_styles_xml     = _read_if_present(szf, s_names, _STYLES_PATH)
        s_rels_xml       = _read_if_present(szf, s_names, _RELS_PATH)
        s_ct_xml         = _read_if_present(szf, s_names, _CT_PATH)
        s_media          = {
            n: szf.read(n) for n in s_names if n.startswith("word/media/")
        }
        s_numbering_xml  = _read_if_present(szf, s_names, _NUMBERING_PATH)
        s_font_xml       = _read_if_present(szf, s_names, _FONT_TABLE_PATH)
        s_footnotes_xml  = _read_if_present(szf, s_names, _FOOTNOTES_PATH)
        s_endnotes_xml   = _read_if_present(szf, s_names, _ENDNOTES_PATH)
        s_footnotes_rels_xml = _read_if_present(szf, s_names, _FOOTNOTES_RELS_PATH)
        s_endnotes_rels_xml  = _read_if_present(szf, s_names, _ENDNOTES_RELS_PATH)

        s_hf_raw: dict[str, bytes] = {}
        if s_rels_xml is not None:
            for _rid, (_rtype, _target, _tmode) in _parse_rels(s_rels_xml).items():
                if _tmode == "External":
                    continue
                if not (
                    _rtype.endswith(_HEADER_REL_SUFFIX)
                    or _rtype.endswith(_FOOTER_REL_SUFFIX)
                ):
                    continue
                hf_zip_path = _resolve_zip_path(_target)
                if hf_zip_path in s_names:
                    s_hf_raw[hf_zip_path] = szf.read(hf_zip_path)
                hf_rels_path = (
                    f"{posixpath.dirname(hf_zip_path)}/_rels/"
                    f"{posixpath.basename(hf_zip_path)}.rels"
                )
                if hf_rels_path in s_names:
                    s_hf_raw[hf_rels_path] = szf.read(hf_rels_path)

        s_theme_xml: bytes | None = None
        if s_rels_xml is not None:
            for _rtype, _target, _tmode in _parse_rels(s_rels_xml).values():
                if _tmode != "External" and _rtype.endswith("/theme"):
                    s_theme_xml = _read_if_present(
                        szf, s_names, _resolve_zip_path(_target),
                    )
                    break

    # Медиафайлы титула получают уникальный префикс, чтобы не перезаписать
    # одноимённые файлы документа.
    fname_map: dict[str, str] = {}
    for mf in s_media:
        prefix        = uuid.uuid4().hex
        fname_map[mf] = f"word/media/{prefix}_{Path(mf).name}"

    rid_map, new_entries, hf_fname_map, hf_content_types = _compute_rels_remap(
        fname_map, s_rels_xml, b_rels_xml, s_header_footer=s_hf_raw,
    )

    foot_rid_map, foot_new_entries, _, _ = _compute_rels_remap(
        fname_map, s_footnotes_rels_xml, b_footnotes_rels_xml,
    )
    end_rid_map, end_new_entries, _, _ = _compute_rels_remap(
        fname_map, s_endnotes_rels_xml, b_endnotes_rels_xml,
    )

    base_elems, base_sect_pr = _parse_body_elems(b_doc_xml, base_path.name)
    src_elems,  src_sect_pr  = _parse_body_elems(s_doc_xml, src_path.name)

    hf_elems: dict[str, ET._Element] = {
        zip_path: _fromstring(raw)
        for zip_path, raw in s_hf_raw.items()
        if not zip_path.endswith(".rels")
    }

    _strip_printer_settings_refs(
        src_elems,
        [src_sect_pr] if src_sect_pr is not None else [],
        hf_elems.values(),
    )

    # Связи, которые не переносятся, не должны оставлять ссылок r:id: они
    # указывали бы в никуда или, хуже, на постороннюю часть документа
    # с тем же rId.
    dropped_refs = _strip_dropped_rel_refs(
        src_elems + ([src_sect_pr] if src_sect_pr is not None else []),
        _dropped_rel_ids(s_rels_xml, rid_map),
    )
    if dropped_refs:
        logger.warning(
            "Титульный лист: %d объект(ов) неподдерживаемого типа (диаграммы, "
            "OLE-объекты, SmartArt и т.п.) удалены из итогового документа.",
            dropped_refs,
        )
    if _strip_comment_anchors(src_elems + list(hf_elems.values())):
        logger.warning(
            "Титульный лист: примечания (комментарии) Word не переносятся "
            "и удалены из итогового документа."
        )

    src_styles = _parse_styles(s_styles_xml)

    theme_fonts = _read_theme_fonts(s_theme_xml)
    title_parts = src_elems + list(hf_elems.values()) + list(src_styles.values())
    _detach_theme_formatting(title_parts, theme_fonts)
    _strip_outline_levels(title_parts)

    base_style_ids: set[str] = set()
    if b_styles_xml:
        try:
            _bsr = _fromstring(b_styles_xml)
            base_style_ids = {
                el.get(f"{{{_WNS}}}styleId", "")
                for el in _bsr.findall(f"{{{_WNS}}}style")
            }
        except Exception:
            logger.warning(
                "Не удалось разобрать styles.xml основного документа — "
                "проверка конфликтов префиксов стилей пропущена.",
                exc_info=True,
            )

    # Новый id стиля титула не должен совпасть ни со стилем документа,
    # ни с уже выданным id.
    rename_map: dict[str, str] = {}
    for sid in src_styles:
        candidate = src_prefix + sid
        suffix = 0
        while candidate in base_style_ids or candidate in rename_map.values():
            suffix += 1
            candidate = f"{src_prefix}{sid}_{suffix}"
        rename_map[sid] = candidate

    default_sid = _find_default_paragraph_style_id(src_styles)

    if rename_map:
        for elem in src_elems:
            _rename_style_refs(elem, rename_map)
        if src_sect_pr is not None:
            _rename_style_refs(src_sect_pr, rename_map)

    # Абзацы титула без явного стиля должны получать стиль абзаца
    # по умолчанию самого титула, а не Normal документа.
    if default_sid is not None:
        implicit_target = rename_map[default_sid]
    else:
        implicit_target = _add_synthetic_default_style(
            src_styles, base_style_ids, rename_map, src_prefix,
        )
    _inject_implicit_style(
        src_elems + list(hf_elems.values()), implicit_target,
        include_styled_tables=True,
    )

    char_default_sid = _find_default_style_id(
        src_styles, "character", "DefaultParagraphFont",
    )
    char_default_target = (
        rename_map.get(char_default_sid) if char_default_sid is not None else None
    )
    _inject_implicit_run_style(
        src_elems + list(hf_elems.values()), char_default_target,
    )
    table_default_sid = _find_default_style_id(src_styles, "table", "TableNormal")
    table_default_target = (
        rename_map.get(table_default_sid) if table_default_sid is not None else None
    )
    _inject_implicit_table_style(
        src_elems + list(hf_elems.values()), table_default_target,
    )

    _remap_drawing_ids(src_elems + list(hf_elems.values()), base_elems)

    bookmark_elems     = src_elems + list(hf_elems.values())
    bookmark_name_map  = _remap_bookmark_ids(bookmark_elems, base_elems)
    _rename_bookmark_field_refs(bookmark_elems, bookmark_name_map)

    if rename_map:
        for hf_root in hf_elems.values():
            _rename_style_refs(hf_root, rename_map)

    for style_el in src_styles.values():
        _rename_style_refs(style_el, rename_map)
    _resolve_undeclared_style_refs(
        list(src_elems) + list(hf_elems.values()) + list(src_styles.values()),
        declared_ids=set(rename_map.values()),
        paragraph_fallback=implicit_target,
    )

    if rid_map:
        _rename_rids(src_elems, rid_map)
        if src_sect_pr is not None:
            _rename_rids([src_sect_pr], rid_map)

    _embed_sect_pr(
        src_elems, src_sect_pr,
        break_type=break_type, filler_style_id=implicit_target,
    )

    merged_numbering, num_remap = _merge_numbering(
        b_numbering_xml, s_numbering_xml, src_elems, rename_map,
        extra_elems=list(hf_elems.values()) + list(src_styles.values()),
    )

    notes_collision_base = base_elems + src_elems + list(hf_elems.values())
    merged_footnotes = _merge_notes(
        b_footnotes_xml, s_footnotes_xml, src_elems,
        ref_tag=f"{{{_WNS}}}footnoteReference",
        rid_map=foot_rid_map,
        dropped_rids=_dropped_rel_ids(s_footnotes_rels_xml, foot_rid_map),
        rename_map=rename_map,
        existing_elems=notes_collision_base,
        style_declared_ids=set(rename_map.values()),
        style_fallback=implicit_target,
        num_remap=num_remap,
    )
    merged_endnotes = _merge_notes(
        b_endnotes_xml, s_endnotes_xml, src_elems,
        ref_tag=f"{{{_WNS}}}endnoteReference",
        rid_map=end_rid_map,
        dropped_rids=_dropped_rel_ids(s_endnotes_rels_xml, end_rid_map),
        rename_map=rename_map,
        existing_elems=notes_collision_base,
        style_declared_ids=set(rename_map.values()),
        style_fallback=implicit_target,
        num_remap=num_remap,
    )

    _isolate_root_paragraph_styles(
        src_styles,
        title_defaults=_read_doc_defaults(s_styles_xml, theme_fonts),
        document_defaults=_read_doc_defaults(b_styles_xml, None),
    )

    merged_styles_xml = _build_merged_styles(
        b_styles_xml, src_styles, rename_map, src_prefix,
    )

    merged_font = _merge_font_table(b_font_xml, s_font_xml)

    merged_doc_xml = _build_merged_document(
        b_doc_xml, base_path.name,
        leading_elems=src_elems,
        trailing_elems=base_elems,
        final_sect_pr=base_sect_pr,
        src_doc_xml=s_doc_xml,
    )

    replacements: dict[str, bytes] = {
        _DOC_PATH:    merged_doc_xml,
        _STYLES_PATH: merged_styles_xml,
    }

    if merged_numbering is not None:
        replacements[_NUMBERING_PATH] = merged_numbering
    if merged_font is not None:
        replacements[_FONT_TABLE_PATH] = merged_font
    if merged_footnotes is not None:
        replacements[_FOOTNOTES_PATH] = merged_footnotes
    if merged_endnotes is not None:
        replacements[_ENDNOTES_PATH] = merged_endnotes

    extra_files: dict[str, bytes] | None = None

    numbering_is_new  = merged_numbering is not None and b_numbering_xml is None
    footnotes_is_new  = merged_footnotes is not None and b_footnotes_xml is None
    endnotes_is_new   = merged_endnotes  is not None and b_endnotes_xml  is None
    font_table_is_new = merged_font      is not None and b_font_xml      is None
    has_media_files   = bool(fname_map)
    has_hf_files      = bool(hf_fname_map)

    # Если у документа не было numbering.xml, сносок или таблицы шрифтов,
    # а у титула они есть, для новых частей нужны связи и типы содержимого.
    if (
        new_entries or numbering_is_new or footnotes_is_new
        or endnotes_is_new or font_table_is_new
    ):
        out_rels    = _parse_rels(b_rels_xml) if b_rels_xml else {}
        rels_to_add = list(new_entries)
        used_ids    = set(out_rels) | {rid for rid, *_ in rels_to_add}
        next_n      = _max_rid(out_rels) + 1
        if numbering_is_new:
            new_rid, next_n = _next_free_rid(used_ids, next_n)
            used_ids.add(new_rid)
            rels_to_add.append((new_rid, _NUMBERING_REL_TYPE, "numbering.xml", None))
        if footnotes_is_new:
            new_rid, next_n = _next_free_rid(used_ids, next_n)
            used_ids.add(new_rid)
            rels_to_add.append((new_rid, _FOOTNOTES_REL_TYPE, "footnotes.xml", None))
        if endnotes_is_new:
            new_rid, next_n = _next_free_rid(used_ids, next_n)
            used_ids.add(new_rid)
            rels_to_add.append((new_rid, _ENDNOTES_REL_TYPE, "endnotes.xml", None))
        if font_table_is_new:
            new_rid, next_n = _next_free_rid(used_ids, next_n)
            used_ids.add(new_rid)
            rels_to_add.append((new_rid, _FONT_TABLE_REL_TYPE, "fontTable.xml", None))
        replacements[_RELS_PATH] = _build_rels_xml(out_rels, rels_to_add)

    if foot_new_entries:
        foot_out_rels = (
            _parse_rels(b_footnotes_rels_xml) if b_footnotes_rels_xml else {}
        )
        replacements[_FOOTNOTES_RELS_PATH] = _build_rels_xml(
            foot_out_rels, foot_new_entries,
        )
    if end_new_entries:
        end_out_rels = (
            _parse_rels(b_endnotes_rels_xml) if b_endnotes_rels_xml else {}
        )
        replacements[_ENDNOTES_RELS_PATH] = _build_rels_xml(
            end_out_rels, end_new_entries,
        )

    if b_ct_xml:
        ct_xml     = b_ct_xml
        ct_changed = False
        if has_media_files:
            ct_xml     = _merge_content_types(
                ct_xml, s_ct_xml, list(s_media.keys()), fname_map,
            )
            ct_changed = True
        for is_new, part_name, ct_type in (
            (numbering_is_new,  "/word/numbering.xml", _NUMBERING_CT_TYPE),
            (footnotes_is_new,  "/word/footnotes.xml", _FOOTNOTES_CT_TYPE),
            (endnotes_is_new,   "/word/endnotes.xml",  _ENDNOTES_CT_TYPE),
            (font_table_is_new, "/word/fontTable.xml", _FONT_TABLE_CT_TYPE),
        ):
            if is_new:
                ct_xml     = _add_ct_override(ct_xml, part_name, ct_type)
                ct_changed = True
        if has_hf_files:
            ct_xml = _add_hf_content_types(ct_xml, hf_fname_map, hf_content_types)
            ct_changed = True
        if ct_changed:
            replacements[_CT_PATH] = ct_xml

    if has_media_files or has_hf_files:
        extra_files = {}
        if has_media_files:
            extra_files.update(
                {fname_map.get(zp, zp): data for zp, data in s_media.items()}
            )
        if has_hf_files:
            extra_files.update(
                _build_hf_extra_files(
                    hf_fname_map, hf_elems, s_hf_raw, fname_map,
                )
            )

    _write_zip(base_path, output_path, replacements, extra_files)



def _build_hf_extra_files(
    hf_fname_map: dict[str, str],
    hf_elems: dict[str, ET._Element],
    s_hf_raw: dict[str, bytes],
    media_fname_map: dict[str, str],
) -> dict[str, bytes]:
    """
    Готовит XML колонтитулов титула под новыми именами вместе с их .rels.
    Ссылки на связи, которые не удалось перенести, удаляются из колонтитулов.
    """
    result: dict[str, bytes] = {}

    for old_xml_path, new_xml_path in hf_fname_map.items():
        old_rels_path = (
            f"{posixpath.dirname(old_xml_path)}/_rels/"
            f"{posixpath.basename(old_xml_path)}.rels"
        )
        if old_rels_path in s_hf_raw:
            new_rels_path = (
                f"{posixpath.dirname(new_xml_path)}/_rels/"
                f"{posixpath.basename(new_xml_path)}.rels"
            )
            rels_bytes, dropped_rids = _patch_hf_rels(
                s_hf_raw[old_rels_path], media_fname_map,
            )
            result[new_rels_path] = rels_bytes
            if old_xml_path in hf_elems:
                _strip_dropped_rel_refs([hf_elems[old_xml_path]], dropped_rids)

        if old_xml_path in hf_elems:
            result[new_xml_path] = _xml_bytes(hf_elems[old_xml_path])

    return result


def _patch_hf_rels(
    rels_xml: bytes,
    media_fname_map: dict[str, str],
) -> tuple[bytes, set[str]]:
    """
    Переписывает .rels колонтитула для объединённого пакета.
    Возвращает новый XML и id удалённых связей, чтобы вызывающий код
    убрал ссылки на них.
    """
    root = _fromstring(rels_xml)
    dropped: set[str] = set()

    def drop(rel_el: ET._Element) -> None:
        rel_id = rel_el.get("Id")
        if rel_id:
            dropped.add(rel_id)
        root.remove(rel_el)

    for el in list(root.findall(f"{{{_REL_NS}}}Relationship")):
        target      = el.get("Target", "")
        target_mode = el.get("TargetMode")
        rtype       = el.get("Type", "")

        if target_mode == "External":
            continue

        if any(
            rtype.endswith(suffix)
            for suffix in _SKIPPABLE_REL_SUFFIXES | _COMMENT_REL_SUFFIXES
        ):
            drop(el)
            continue

        old_zip  = _resolve_zip_path(target)
        resolved = media_fname_map.get(old_zip)
        if resolved is None:
            stripped = re.sub(r"^(?:\.\./|\./)+", "", target).lstrip("/")
            resolved = media_fname_map.get(f"word/{stripped}")
        if resolved is not None:
            el.set("Target", posixpath.relpath(resolved, "word"))
            continue

        is_media = (
            any(target.lower().endswith(f".{ext}") for ext in _MEDIA_EXTENSIONS)
            or "image" in rtype.lower()
        )
        if is_media:
            logger.warning(
                "Изображение из колонтитула присоединяемого документа (target: %s) "
                "не найдено среди скопированных медиафайлов — ссылка на него "
                "в итоговом документе может оказаться битой.",
                target,
            )
            continue

        logger.warning(
            "Неподдерживаемый встроенный объект в колонтитуле пропущен "
            "(тип связи: %s, target: %s). "
            "Удалите этот объект из документа или вставьте его как обычное "
            "изображение, чтобы он попал в итоговый файл.",
            rtype.rsplit("/", 1)[-1],
            target,
        )
        drop(el)

    return _xml_bytes(root), dropped


def _add_hf_content_types(
    ct_xml: bytes,
    hf_fname_map: dict[str, str],
    hf_content_types: dict[str, str],
) -> bytes:
    """Регистрирует перенесённые колонтитулы в [Content_Types].xml."""
    root = _fromstring(ct_xml)
    existing = {el.get("PartName") for el in root.findall(f"{{{_CT_NS}}}Override")}
    for new_path in hf_fname_map.values():
        part_name = f"/{new_path}"
        if part_name in existing:
            continue
        ET.SubElement(
            root,
            f"{{{_CT_NS}}}Override",
            PartName=part_name,
            ContentType=hf_content_types[new_path],
        )
        existing.add(part_name)
    return _xml_bytes(root)



def _parse_body_elems(
    doc_xml: bytes,
    source_name: str,
) -> tuple[list[ET._Element], ET._Element]:
    """
    Возвращает копии дочерних элементов <w:body> и отдельно — завершающий
    <w:sectPr> (если его нет, создаётся пустой с предупреждением).
    """
    root = _fromstring(doc_xml)
    body = root.find(f"{{{_WNS}}}body")

    if body is None:
        raise ValueError(
            f"В '{source_name}' отсутствует элемент <w:body>. "
            "Возможно, файл повреждён."
        )

    elems:   list[ET._Element]     = []
    sect_pr: ET._Element | None = None

    for el in body:
        if el.tag == f"{{{_WNS}}}sectPr":
            sect_pr = copy.deepcopy(el)
        else:
            elems.append(copy.deepcopy(el))

    if sect_pr is None:
        logger.warning(
            "В '%s' нет <w:sectPr> верхнего уровня: параметры раздела "
            "(размер страницы, поля, колонтитулы) не заданы. Добавлен "
            "пустой <w:sectPr> по умолчанию, чтобы итоговый файл получил "
            "корректный и независимый раздел для этого документа.",
            source_name,
        )
        sect_pr = ET.Element(f"{{{_WNS}}}sectPr", nsmap={"w": _WNS})

    return elems, sect_pr


def _parse_styles(styles_xml: bytes | None) -> dict[str, ET._Element]:
    """Стили из styles.xml по styleId; пустой словарь, если части нет."""
    if styles_xml is None:
        return {}
    root   = _fromstring(styles_xml)
    result: dict[str, ET._Element] = {}
    for el in root.findall(f"{{{_WNS}}}style"):
        sid = el.get(f"{{{_WNS}}}styleId")
        if sid:
            result[sid] = el
    return result



def _rename_style_refs(elem: ET._Element, rename_map: dict[str, str]) -> None:
    """Заменяет во всём поддереве ссылки на стили (pStyle, rStyle, basedOn и т. д.) по rename_map."""
    w_val = f"{{{_WNS}}}val"
    for node in elem.iter():
        if node.tag in _STYLE_REF_TAGS:
            val = node.get(w_val)
            if val and val in rename_map:
                node.set(w_val, rename_map[val])


def _resolve_undeclared_style_refs(
    elements: Iterable[ET._Element],
    declared_ids: set[str],
    paragraph_fallback: str | None,
) -> None:
    """
    Убирает ссылки на стили, которых нет среди declared_ids. Такие ссылки
    не переименованы и иначе привязались бы к одноимённому стилю документа.
    pStyle заменяется на paragraph_fallback (или удаляется, если его нет),
    остальные ссылки удаляются.
    """
    w_val = f"{{{_WNS}}}val"
    p_style_tag = f"{{{_WNS}}}pStyle"
    for root_el in elements:
        if root_el is None:
            continue
        for node in list(root_el.iter()):
            tag = node.tag
            if tag == p_style_tag:
                val = node.get(w_val)
                if not val or val in declared_ids:
                    continue
                if paragraph_fallback:
                    node.set(w_val, paragraph_fallback)
                else:
                    parent = node.getparent()
                    if parent is not None:
                        parent.remove(node)
            elif tag in _UNDECLARED_STYLE_DROP_TAGS:
                val = node.get(w_val)
                if val and val not in declared_ids:
                    parent = node.getparent()
                    if parent is not None:
                        parent.remove(node)


def _find_default_paragraph_style_id(src_styles: dict[str, ET._Element]) -> str | None:
    """
    Стиль абзаца по умолчанию титула: помеченный w:default="1", иначе Normal.
    Если нет ни того, ни другого, возвращает None с предупреждением.
    """
    explicit = next(
        (
            sid for sid, el in src_styles.items()
            if el.get(f"{{{_WNS}}}type") == "paragraph"
            and el.get(f"{{{_WNS}}}default") == "1"
        ),
        None,
    )
    if explicit is not None:
        return explicit
    normal = src_styles.get("Normal")
    if normal is not None and normal.get(f"{{{_WNS}}}type") == "paragraph":
        return "Normal"
    logger.warning(
        "Не удалось определить стиль абзаца по умолчанию в присоединяемом "
        "документе (ни один стиль не помечен w:default=\"1\" и нет "
        "стиля абзаца с id \"Normal\"). Абзацы без явно заданного стиля "
        "могут унаследовать параметры по умолчанию основного документа "
        "вместо собственных."
    )
    return None


def _find_default_style_id(
    src_styles: dict[str, ET._Element],
    style_type: str,
    builtin_id: str,
) -> str | None:
    """Стиль по умолчанию заданного типа: помеченный w:default="1", иначе встроенный builtin_id; None, если нет обоих."""
    explicit = next(
        (
            sid for sid, el in src_styles.items()
            if el.get(f"{{{_WNS}}}type") == style_type
            and el.get(f"{{{_WNS}}}default") == "1"
        ),
        None,
    )
    if explicit is not None:
        return explicit
    builtin = src_styles.get(builtin_id)
    if builtin is not None and builtin.get(f"{{{_WNS}}}type") == style_type:
        return builtin_id
    return None


def _add_synthetic_default_style(
    src_styles: dict[str, ET._Element],
    base_style_ids: set[str],
    rename_map: dict[str, str],
    src_prefix: str,
) -> str:
    """
    Создаёт пустой корневой стиль абзаца для титула, в котором нет стиля
    абзаца по умолчанию. Позже изоляция корневых стилей заполнит его
    docDefaults титула, и абзацы без стиля будут выглядеть как в исходном файле.
    """
    used_ids  = base_style_ids | set(rename_map.values())
    candidate = f"{src_prefix}Normal"
    suffix    = 0
    while candidate in used_ids:
        suffix   += 1
        candidate = f"{src_prefix}Normal_{suffix}"

    style_el = ET.Element(f"{{{_WNS}}}style")
    style_el.set(f"{{{_WNS}}}type", "paragraph")
    style_el.set(f"{{{_WNS}}}styleId", candidate)
    name_el = ET.SubElement(style_el, f"{{{_WNS}}}name")
    name_el.set(f"{{{_WNS}}}val", "Normal")

    src_styles[candidate] = style_el
    rename_map[candidate] = candidate
    return candidate


def _in_explicitly_styled_table(elem: ET._Element) -> bool:
    """Находится ли элемент в таблице (ближайшей охватывающей), у которой явно задан стиль."""
    tbl_tag       = f"{{{_WNS}}}tbl"
    tbl_pr_tag    = f"{{{_WNS}}}tblPr"
    tbl_style_tag = f"{{{_WNS}}}tblStyle"
    ancestor = elem.getparent()
    while ancestor is not None:
        if ancestor.tag == tbl_tag:
            tbl_pr = ancestor.find(tbl_pr_tag)
            return tbl_pr is not None and tbl_pr.find(tbl_style_tag) is not None
        ancestor = ancestor.getparent()
    return False


def _inject_implicit_style(
    src_elems: list[ET._Element],
    target_style_id: str | None,
    include_styled_tables: bool = False,
) -> None:
    """
    Явно назначает target_style_id абзацам без pStyle, чтобы после слияния
    они не получили Normal документа. Абзацы в таблицах с явным стилем
    пропускаются (их оформление задаёт стиль таблицы), если не включён
    include_styled_tables.
    """
    if target_style_id is None:
        return

    p_tag       = f"{{{_WNS}}}p"
    p_pr_tag    = f"{{{_WNS}}}pPr"
    p_style_tag = f"{{{_WNS}}}pStyle"

    for root_el in src_elems:
        for p in root_el.iter(p_tag):
            p_pr = p.find(p_pr_tag)
            if p_pr is not None and p_pr.find(p_style_tag) is not None:
                continue
            if not include_styled_tables and _in_explicitly_styled_table(p):
                continue
            if p_pr is None:
                p_pr = ET.Element(p_pr_tag)
                p.insert(0, p_pr)
            style_node = ET.Element(p_style_tag)
            style_node.set(f"{{{_WNS}}}val", target_style_id)
            p_pr.insert(0, style_node)


def _inject_implicit_run_style(
    src_elems: list[ET._Element],
    target_style_id: str | None,
) -> None:
    """Явно назначает стиль знака по умолчанию титула фрагментам без rStyle (кроме фрагментов в таблицах с явным стилем)."""
    if target_style_id is None:
        return

    r_tag       = f"{{{_WNS}}}r"
    r_pr_tag    = f"{{{_WNS}}}rPr"
    r_style_tag = f"{{{_WNS}}}rStyle"

    for root_el in src_elems:
        for r in root_el.iter(r_tag):
            r_pr = r.find(r_pr_tag)
            if r_pr is not None and r_pr.find(r_style_tag) is not None:
                continue
            if _in_explicitly_styled_table(r):
                continue
            if r_pr is None:
                r_pr = ET.Element(r_pr_tag)
                r.insert(0, r_pr)
            style_node = ET.Element(r_style_tag)
            style_node.set(f"{{{_WNS}}}val", target_style_id)
            r_pr.insert(0, style_node)


def _inject_implicit_table_style(
    src_elems: list[ET._Element],
    target_style_id: str | None,
) -> None:
    """Явно назначает стиль таблицы по умолчанию титула таблицам без tblStyle."""
    if target_style_id is None:
        return

    tbl_tag       = f"{{{_WNS}}}tbl"
    tbl_pr_tag    = f"{{{_WNS}}}tblPr"
    tbl_style_tag = f"{{{_WNS}}}tblStyle"

    for root_el in src_elems:
        for tbl in root_el.iter(tbl_tag):
            tbl_pr = tbl.find(tbl_pr_tag)
            if tbl_pr is not None and tbl_pr.find(tbl_style_tag) is not None:
                continue
            if tbl_pr is None:
                tbl_pr = ET.Element(tbl_pr_tag)
                tbl.insert(0, tbl_pr)
            style_node = ET.Element(tbl_style_tag)
            style_node.set(f"{{{_WNS}}}val", target_style_id)
            tbl_pr.insert(0, style_node)


def _collect_max_drawing_id(elems: list[ET._Element]) -> int:
    """Наибольший числовой id объектов рисования (docPr, cNvPr)."""
    max_id = 0
    for root_el in elems:
        for node in root_el.iter():
            if node.tag in _DRAWING_ID_TAGS:
                val = node.get("id")
                if val and val.isdigit():
                    max_id = max(max_id, int(val))
    return max_id


def _remap_drawing_ids(
    src_elems: list[ET._Element],
    base_elems: list[ET._Element],
) -> None:
    """
    Перенумеровывает id объектов рисования титула начиная с максимального id
    документа плюс один: повторяющиеся id Word считает повреждением документа.
    """
    max_base_id = _collect_max_drawing_id(base_elems)
    next_id     = max_base_id + 1

    old_to_new: dict[str, str] = {}
    for root_el in src_elems:
        for node in root_el.iter():
            if node.tag in _DRAWING_ID_TAGS:
                val = node.get("id")
                if val and val not in old_to_new:
                    old_to_new[val] = str(next_id)
                    next_id += 1

    if not old_to_new:
        return

    for root_el in src_elems:
        for node in root_el.iter():
            if node.tag in _DRAWING_ID_TAGS:
                val = node.get("id")
                if val and val in old_to_new:
                    node.set("id", old_to_new[val])


def _collect_max_bookmark_id(elems: list[ET._Element]) -> int:
    """Наибольший числовой w:id закладок и исправлений (ins/del/moveFrom/moveTo)."""
    max_id = 0
    w_id   = f"{{{_WNS}}}id"
    for root_el in elems:
        for node in root_el.iter():
            if node.tag in _BOOKMARK_ID_TAGS:
                val = node.get(w_id)
                if val and val.lstrip("-").isdigit():
                    max_id = max(max_id, int(val))
    return max_id


def _collect_bookmark_names(elems: Iterable[ET._Element]) -> set[str]:
    """Имена всех закладок в элементах."""
    w_name = f"{{{_WNS}}}name"
    names: set[str] = set()
    for root_el in elems:
        if root_el is None:
            continue
        for node in root_el.iter(_BOOKMARK_START_TAG):
            val = node.get(w_name)
            if val:
                names.add(val)
    return names


def _remap_bookmark_ids(
    src_elems: list[ET._Element],
    base_elems: list[ET._Element],
) -> dict[str, str]:
    """
    Перенумеровывает w:id закладок и исправлений титула, чтобы они
    не пересекались с документом, а имена закладок, уже занятые в документе,
    меняет на «имя_dupN». Возвращает карту переименования имён.
    """
    max_base_id = _collect_max_bookmark_id(base_elems)
    next_id     = max_base_id + 1
    w_id        = f"{{{_WNS}}}id"
    w_name      = f"{{{_WNS}}}name"

    old_to_new: dict[str, str] = {}
    for root_el in src_elems:
        for node in root_el.iter():
            if node.tag in _BOOKMARK_ID_TAGS:
                val = node.get(w_id)
                if val and val not in old_to_new:
                    old_to_new[val] = str(next_id)
                    next_id += 1

    for root_el in src_elems:
        for node in root_el.iter():
            if node.tag in _BOOKMARK_ID_TAGS:
                val = node.get(w_id)
                if val and val in old_to_new:
                    node.set(w_id, old_to_new[val])

    used_names: set[str] = set(_collect_bookmark_names(base_elems))
    name_map: dict[str, str] = {}
    for root_el in src_elems:
        for node in root_el.iter(_BOOKMARK_START_TAG):
            name = node.get(w_name)
            if not name or name in name_map:
                continue
            if name not in used_names:
                used_names.add(name)
                continue
            suffix    = 1
            candidate = f"{name}_dup{suffix}"
            while candidate in used_names:
                suffix    += 1
                candidate = f"{name}_dup{suffix}"
            used_names.add(candidate)
            name_map[name] = candidate

    if name_map:
        for root_el in src_elems:
            for node in root_el.iter(_BOOKMARK_START_TAG):
                name = node.get(w_name)
                if name and name in name_map:
                    node.set(w_name, name_map[name])

    return name_map


# Ссылки на закладки в кодах полей: REF/PAGEREF/NOTEREF имя и HYPERLINK \l "имя".
_FIELD_BOOKMARK_REF_RE = re.compile(r"\b(REF|PAGEREF|NOTEREF)(\s+)(\S+)")

_FIELD_HYPERLINK_L_RE = re.compile(r'(\\l\s+)(?:"([^"]*)"|(\S+))')
_HYPERLINK_KEYWORD_RE = re.compile(r"\bHYPERLINK\b")


def _patch_field_code_text(text: str, name_map: dict[str, str]) -> str:
    """Заменяет имена закладок в коде поля (REF, PAGEREF, NOTEREF, HYPERLINK \\l) по name_map."""
    def _sub_ref(match: "re.Match[str]") -> str:
        new_name = name_map.get(match.group(3))
        if new_name is None:
            return match.group(0)
        return f"{match.group(1)}{match.group(2)}{new_name}"

    text = _FIELD_BOOKMARK_REF_RE.sub(_sub_ref, text)

    if _HYPERLINK_KEYWORD_RE.search(text):
        def _sub_hyperlink(match: "re.Match[str]") -> str:
            quoted   = match.group(2) is not None
            old_name = match.group(2) if quoted else match.group(3)
            new_name = name_map.get(old_name)
            if new_name is None:
                return match.group(0)
            return f'{match.group(1)}"{new_name}"' if quoted \
                else f"{match.group(1)}{new_name}"

        text = _FIELD_HYPERLINK_L_RE.sub(_sub_hyperlink, text)

    return text


def _rename_bookmark_field_refs(
    elements: Iterable[ET._Element],
    name_map: dict[str, str],
) -> None:
    """Применяет переименование закладок к гиперссылкам (w:anchor) и кодам полей (instrText, fldSimple)."""
    if not name_map:
        return

    w_anchor = f"{{{_WNS}}}anchor"
    w_instr  = f"{{{_WNS}}}instr"

    for root_el in elements:
        if root_el is None:
            continue
        for node in root_el.iter():
            if node.tag == _HYPERLINK_TAG:
                val = node.get(w_anchor)
                if val and val in name_map:
                    node.set(w_anchor, name_map[val])
            elif node.tag == _INSTR_TEXT_TAG:
                if node.text:
                    node.text = _patch_field_code_text(node.text, name_map)
            elif node.tag == _FLD_SIMPLE_TAG:
                instr = node.get(w_instr)
                if instr:
                    node.set(w_instr, _patch_field_code_text(instr, name_map))


def _dropped_rel_ids(rels_xml: bytes | None, rid_map: dict[str, str]) -> set[str]:
    """id связей исходной части, которые не были перенесены."""
    if rels_xml is None:
        return set()
    return set(_parse_rels(rels_xml)) - set(rid_map)


def _dropped_ref_owner(node: ET._Element) -> ET._Element | None:
    """
    Возвращает элемент, который нужно удалить из-за ссылки на удалённую связь:
    саму ссылку (ссылки на колонтитулы) или охватывающий drawing/pict/object,
    расширенный до обёртки mc:AlternateContent, чтобы не осталась одна
    запасная ветка. None означает «удалить только атрибут».
    """
    if node.tag in _DROPPED_REF_SELF_TAGS:
        return node
    owner = node
    while owner is not None and owner.tag not in _DROPPED_REF_OWNER_TAGS:
        owner = owner.getparent()
    if owner is None:
        return None
    ancestor = owner.getparent()
    while ancestor is not None and ancestor.tag != f"{{{_WNS}}}r":
        if ancestor.tag == f"{{{_MC_NS}}}AlternateContent":
            return ancestor
        ancestor = ancestor.getparent()
    return owner


def _strip_dropped_rel_refs(
    elems: Iterable[ET._Element | None],
    dropped_rids: set[str],
) -> int:
    """
    Удаляет все элементы, ссылающиеся на одну из *dropped_rids* через
    атрибут r:*. Возвращает число найденных ссылок.
    """
    if not dropped_rids:
        return 0
    r_prefix = f"{{{_R_NS}}}"
    hits: list[tuple[ET._Element, str]] = []
    for root_el in elems:
        if root_el is None:
            continue
        for node in root_el.iter():
            for attr_name, value in node.attrib.items():
                if attr_name.startswith(r_prefix) and value in dropped_rids:
                    hits.append((node, attr_name))

    for node, attr_name in hits:
        owner = _dropped_ref_owner(node)
        parent = owner.getparent() if owner is not None else None
        if parent is not None:
            parent.remove(owner)
        else:
            node.attrib.pop(attr_name, None)
    return len(hits)


def _strip_comment_anchors(elems: Iterable[ET._Element | None]) -> int:
    """
    Удаляет диапазоны и ссылки примечаний (comments.xml не переносится).
    Фрагмент текста (run), в котором остались только свойства, тоже удаляется.
    Возвращает число удалённых якорей.
    """
    run_tag = f"{{{_WNS}}}r"
    rpr_tag = f"{{{_WNS}}}rPr"
    removed = 0
    for root_el in elems:
        if root_el is None:
            continue
        anchors = [node for node in root_el.iter() if node.tag in _COMMENT_ANCHOR_TAGS]
        for node in anchors:
            parent = node.getparent()
            if parent is None:
                continue
            parent.remove(node)
            removed += 1
            if parent.tag == run_tag and all(child.tag == rpr_tag for child in parent):
                grandparent = parent.getparent()
                if grandparent is not None:
                    grandparent.remove(parent)
    return removed


def _rename_rids(elems: list[ET._Element], rid_map: dict[str, str]) -> None:
    """Заменяет значения атрибутов r:* по rid_map."""
    if not rid_map:
        return
    r_prefix = f"{{{_R_NS}}}"
    for elem in elems:
        for node in elem.iter():
            for attr_name in list(node.attrib):
                if attr_name.startswith(r_prefix):
                    val = node.get(attr_name)
                    if val in rid_map:
                        node.set(attr_name, rid_map[val])




# ── Изоляция титульного листа ───────────────────────────────────────────────
#
# Объединённый пакет сохраняет docDefaults документа, поэтому любая цепочка
# стилей титула теперь упирается в умолчания документа, а не титула.
# Чтобы титул выглядел точно как в собственном файле, каждый корневой стиль
# абзаца титула получает:
#   1. docDefaults самого титула;
#   2. явное значение «по спецификации» для каждого свойства, которое задано
#      в docDefaults документа, но не в docDefaults титула (иначе в титул
#      просочились бы, например, полуторный интервал, кегль 14 и выравнивание
#      по ширине из ГОСТ).
# Значения, заданные в самом стиле, всегда имеют приоритет над обоими.

# Порядок дочерних элементов w:pPr и w:rPr по схеме OOXML: Word может счесть
# документ повреждённым, если добавленные свойства нарушат эту последовательность.

_PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
    "pPrChange",
]

_RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
    "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
    "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
    "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath", "rPrChange",
]

_STYLE_TRAILING_TAGS = ("pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr")

# Свойства-переключатели: значение, действующее при отсутствии свойства.
_PPR_ONOFF_RESETS = {
    "keepNext": "0", "keepLines": "0", "pageBreakBefore": "0",
    "widowControl": "0", "suppressLineNumbers": "0",
    "suppressAutoHyphens": "0", "kinsoku": "1", "wordWrap": "1",
    "overflowPunct": "1", "topLinePunct": "0", "autoSpaceDE": "1",
    "autoSpaceDN": "1", "bidi": "0", "adjustRightInd": "1",
    "snapToGrid": "1", "contextualSpacing": "0", "mirrorIndents": "0",
    "suppressOverlap": "0",
}
_RPR_ONOFF_RESETS = {
    tag: "0" for tag in (
        "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
        "outline", "shadow", "emboss", "imprint", "noProof", "vanish",
        "webHidden", "rtl", "cs", "specVanish", "oMath",
    )
}
_RPR_ONOFF_RESETS["snapToGrid"] = "1"

# Свойства с атрибутами: элемент → атрибут → значение при отсутствии.
_PPR_ATTR_RESETS = {
    "spacing": {
        "before": "0", "after": "0", "line": "240", "lineRule": "auto",
        "beforeAutospacing": "0", "afterAutospacing": "0",
    },
    "ind": {
        "left": "0", "start": "0", "right": "0", "end": "0",
        "firstLine": "0", "hanging": "0",
    },
    "jc":            {"val": "left"},
    "textAlignment": {"val": "auto"},
    "textDirection": {"val": "lrTb"},
}
_RPR_ATTR_RESETS = {
    "rFonts": {
        "ascii": "Times New Roman", "hAnsi": "Times New Roman",
        "eastAsia": "Times New Roman", "cs": "Times New Roman",
    },
    "sz":        {"val": "20"},
    "szCs":      {"val": "20"},
    "color":     {"val": "auto"},
    "spacing":   {"val": "0"},
    "w":         {"val": "100"},
    "kern":      {"val": "0"},
    "position":  {"val": "0"},
    "highlight": {"val": "none"},
    "u":         {"val": "none"},
    "effect":    {"val": "none"},
    "vertAlign": {"val": "baseline"},
    "em":        {"val": "none"},
}

# Группы атрибутов, из которых на одном элементе допустим только один
# (взаимоисключающие значения или синонимы, например left и start).
_EXCLUSIVE_ATTRS = {
    "ind": (
        {"firstLine", "hanging", "firstLineChars", "hangingChars"},
        {"left", "start", "leftChars", "startChars"},
        {"right", "end", "rightChars", "endChars"},
    ),
}

# Атрибут rFonts со ссылкой на шрифт темы → атрибут с конкретным шрифтом.
_THEME_FONT_ATTRS = {
    "asciiTheme":    "ascii",
    "hAnsiTheme":    "hAnsi",
    "eastAsiaTheme": "eastAsia",
    "cstheme":       "cs",
}


def _local(tag: str) -> str:
    """Локальное имя тега без пространства имён."""
    return tag.rsplit("}", 1)[-1]


def _read_theme_fonts(theme_xml: bytes | None) -> dict[str, str]:
    """Сопоставляет ссылкам на шрифты темы (например, ``minorHAnsi``) конкретные гарнитуры."""
    if theme_xml is None:
        return {}
    try:
        root = _fromstring(theme_xml)
    except ValueError:
        logger.warning("Тема титульного листа повреждена — шрифты темы не определены.")
        return {}

    fonts: dict[str, str] = {}
    for scheme, prefix in (("majorFont", "major"), ("minorFont", "minor")):
        node = root.find(f".//{{{_A_NS}}}{scheme}")
        if node is None:
            continue
        latin = node.find(f"{{{_A_NS}}}latin")
        ea    = node.find(f"{{{_A_NS}}}ea")
        cs    = node.find(f"{{{_A_NS}}}cs")
        latin_face = latin.get("typeface", "") if latin is not None else ""
        if latin_face:
            fonts[f"{prefix}HAnsi"] = latin_face
            fonts[f"{prefix}Ascii"] = latin_face
        for key, el in ((f"{prefix}EastAsia", ea), (f"{prefix}Bidi", cs)):
            face = el.get("typeface", "") if el is not None else ""
            if face or latin_face:
                fonts[key] = face or latin_face
    return fonts


def _detach_theme_formatting(
    elements: Iterable[ET._Element],
    theme_fonts: dict[str, str],
) -> None:
    """
    Заменяет ссылки на шрифты и цвета темы конкретными значениями.
    Объединённый пакет содержит тему документа, поэтому оставшаяся в титуле
    ссылка на тему разрешилась бы по чужой теме.
    """
    w = f"{{{_WNS}}}"
    rfonts_tag = f"{w}rFonts"
    for root_el in elements:
        if root_el is None:
            continue
        for node in root_el.iter():
            if node.tag == rfonts_tag:
                for theme_attr, plain_attr in _THEME_FONT_ATTRS.items():
                    ref = node.get(f"{w}{theme_attr}")
                    if ref is None:
                        continue
                    face = theme_fonts.get(ref)
                    if face:
                        node.set(f"{w}{plain_attr}", face)
                    elif node.get(f"{w}{plain_attr}") is None:
                        logger.warning(
                            "Не удалось определить шрифт темы титульного листа %r; "
                            "будет использован шрифт темы документа.", ref,
                        )
                        continue
                    del node.attrib[f"{w}{theme_attr}"]
                continue

            if node.get(f"{w}themeColor") is not None:
                concrete = "val" if _local(node.tag) == "color" else "color"
                if node.get(f"{w}{concrete}") is not None:
                    for attr in ("themeColor", "themeTint", "themeShade"):
                        node.attrib.pop(f"{w}{attr}", None)
            if node.get(f"{w}themeFill") is not None and node.get(f"{w}fill"):
                for attr in ("themeFill", "themeFillTint", "themeFillShade"):
                    node.attrib.pop(f"{w}{attr}", None)


def _strip_outline_levels(elements: Iterable[ET._Element]) -> None:
    """Удаляет все уровни структуры, чтобы ни один абзац титула не попал в оглавление."""
    tag = f"{{{_WNS}}}outlineLvl"
    for root_el in elements:
        if root_el is None:
            continue
        for node in list(root_el.iter(tag)):
            node.getparent().remove(node)


def _read_doc_defaults(
    styles_xml: bytes | None,
    theme_fonts: dict[str, str] | None,
) -> tuple[ET._Element, ET._Element]:
    """
    Возвращает копии pPr и rPr из docDefaults (пустые, если их нет).
    При переданном theme_fonts ссылки на тему заменяются конкретными
    шрифтами; уровни структуры удаляются.
    """
    ppr = ET.Element(f"{{{_WNS}}}pPr")
    rpr = ET.Element(f"{{{_WNS}}}rPr")
    if styles_xml is None:
        return ppr, rpr

    doc_defaults = _fromstring(styles_xml).find(f"{{{_WNS}}}docDefaults")
    if doc_defaults is None:
        return ppr, rpr

    src_ppr = doc_defaults.find(f"{{{_WNS}}}pPrDefault/{{{_WNS}}}pPr")
    src_rpr = doc_defaults.find(f"{{{_WNS}}}rPrDefault/{{{_WNS}}}rPr")
    if src_ppr is not None:
        ppr = copy.deepcopy(src_ppr)
    if src_rpr is not None:
        rpr = copy.deepcopy(src_rpr)

    if theme_fonts is not None:
        _detach_theme_formatting([ppr, rpr], theme_fonts)
    _strip_outline_levels([ppr])
    return ppr, rpr


def _add_spec_resets(
    target: ET._Element,
    leaking: ET._Element,
    onoff_resets: dict[str, str],
    attr_resets: dict[str, dict[str, str]],
) -> None:
    """
    Для каждого свойства из ``leaking`` (docDefaults документа), не заданного
    в ``target`` (docDefaults титула), добавляет значение, которое свойство
    имеет при отсутствии, и тем самым нейтрализует его влияние.
    """
    w = f"{{{_WNS}}}"
    existing = {child.tag: child for child in target}

    for prop in leaking:
        name = _local(prop.tag)
        own  = existing.get(prop.tag)

        if name in onoff_resets:
            if own is None:
                reset = ET.SubElement(target, prop.tag)
                reset.set(f"{w}val", onoff_resets[name])
            continue

        known = attr_resets.get(name)
        if known is None:
            if own is None:
                logger.debug(
                    "Для <w:%s> из параметров документа по умолчанию нет "
                    "значения сброса; это может повлиять на титульный лист.", name,
                )
            continue

        if own is None:
            own = ET.SubElement(target, prop.tag)
            existing[prop.tag] = own
        for attr in prop.attrib:
            plain = _local(attr)
            if plain in known and own.get(attr) is None:
                own.set(attr, known[plain])


def _merge_props_isolated(target: ET._Element, defaults: ET._Element) -> None:
    """
    Добавляет в ``target`` все свойства и атрибуты из ``defaults``, которых
    в нём нет, не сочетая взаимоисключающие атрибуты (например, firstLine
    и hanging).
    """
    existing = {child.tag: child for child in target}
    for prop in defaults:
        own = existing.get(prop.tag)
        if own is None:
            target.append(copy.deepcopy(prop))
            continue
        groups = _EXCLUSIVE_ATTRS.get(_local(prop.tag), ())
        own_locals = {_local(a) for a in own.attrib}
        for attr, value in prop.attrib.items():
            if attr in own.attrib:
                continue
            plain = _local(attr)
            if any(plain in g and own_locals & g for g in groups):
                continue
            own.set(attr, value)


def _sort_children(parent: ET._Element, order: list[str]) -> None:
    """
    Упорядочивает дочерние элементы по списку order. Неизвестные элементы
    получают ранг последнего элемента списка (…Change) и сохраняют
    взаимный порядок.
    """
    rank = {name: i for i, name in enumerate(order)}
    children = list(parent)
    children.sort(key=lambda c: rank.get(_local(c.tag), len(order) - 1))
    for child in children:
        parent.remove(child)
    for child in children:
        parent.append(child)


def _ensure_style_child(style_el: ET._Element, name: str) -> ET._Element:
    """Возвращает дочерний элемент name стиля, при необходимости создавая его в позиции, которую требует схема."""
    tag = f"{{{_WNS}}}{name}"
    found = style_el.find(tag)
    if found is not None:
        return found
    after = _STYLE_TRAILING_TAGS[_STYLE_TRAILING_TAGS.index(name) + 1:]
    child = ET.Element(tag)
    for i, existing in enumerate(style_el):
        if _local(existing.tag) in after:
            style_el.insert(i, child)
            return child
    style_el.append(child)
    return child


def _isolate_root_paragraph_styles(
    src_styles: dict[str, ET._Element],
    title_defaults: tuple[ET._Element, ET._Element],
    document_defaults: tuple[ET._Element, ET._Element],
) -> None:
    """
    Изолирует титул (см. комментарий выше): корневым стилям абзаца титула
    (без basedOn) добавляются docDefaults титула и сбросы свойств, заданных
    в docDefaults документа.
    """
    title_ppr, title_rpr = (copy.deepcopy(e) for e in title_defaults)
    doc_ppr,   doc_rpr   = document_defaults

    _add_spec_resets(title_ppr, doc_ppr, _PPR_ONOFF_RESETS, _PPR_ATTR_RESETS)
    _add_spec_resets(title_rpr, doc_rpr, _RPR_ONOFF_RESETS, _RPR_ATTR_RESETS)

    w = f"{{{_WNS}}}"
    for style_el in src_styles.values():
        if style_el.get(f"{w}type") != "paragraph":
            continue
        if style_el.find(f"{w}basedOn") is not None:
            continue
        if len(title_ppr):
            ppr = _ensure_style_child(style_el, "pPr")
            _merge_props_isolated(ppr, title_ppr)
            _sort_children(ppr, _PPR_ORDER)
        if len(title_rpr):
            rpr = _ensure_style_child(style_el, "rPr")
            _merge_props_isolated(rpr, title_rpr)
            _sort_children(rpr, _RPR_ORDER)


def _build_merged_styles(
    base_styles_xml: bytes,
    src_styles: dict[str, ET._Element],
    rename_map: dict[str, str],
    src_prefix: str,
) -> bytes:
    """
    Добавляет стили титула в styles.xml документа под новыми id.
    Отображаемые имена получают префикс «title: », а флаг w:default
    снимается, чтобы стили титула не стали стилями документа по умолчанию.
    """
    root = _fromstring(base_styles_xml)

    display_pfx = src_prefix.strip("_") + ": "

    for sid, style_el in src_styles.items():
        new_el  = copy.deepcopy(style_el)
        new_sid = rename_map[sid]
        new_el.set(f"{{{_WNS}}}styleId", new_sid)
        new_el.attrib.pop(f"{{{_WNS}}}default", None)
        name_el = new_el.find(f"{{{_WNS}}}name")
        if name_el is not None:
            orig_name = name_el.get(f"{{{_WNS}}}val", "")
            if orig_name:
                name_el.set(f"{{{_WNS}}}val", f"{display_pfx}{orig_name}")
        _rename_style_refs(new_el, rename_map)
        root.append(new_el)

    return _xml_bytes(root)



def _resolve_extra_nsmap(
    base_nsmap: dict[str | None, str],
    extra_source_nsmap: dict[str | None, str],
) -> dict[str | None, str]:
    """
    Пространства имён титула, которых нет у документа. Если префикс в документе
    уже занят другим URI, подбирается свободный (префикс0, префикс1, …).
    """
    extra_nsmap: dict[str | None, str] = {}
    occupied = set(base_nsmap.keys())
    for prefix, uri in extra_source_nsmap.items():
        if prefix in occupied:
            if base_nsmap.get(prefix) == uri:
                continue
            alt_prefix = f"{prefix}0" if prefix else "ns0"
            suffix = 1
            while alt_prefix in occupied:
                alt_prefix = f"{prefix}{suffix}" if prefix else f"ns{suffix}"
                suffix += 1
            extra_nsmap[alt_prefix] = uri
            occupied.add(alt_prefix)
        else:
            extra_nsmap[prefix] = uri
            occupied.add(prefix)
    return extra_nsmap


def _collect_full_nsmap(root: ET._Element) -> dict[str | None, str]:
    """Все объявления пространств имён во всём дереве."""
    nsmap: dict[str | None, str] = {}
    for el in root.iter():
        nsmap.update(el.nsmap)
    return nsmap


_MC_NS        = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_MC_IGNORABLE = f"{{{_MC_NS}}}Ignorable"


def _merge_ignorable(
    root: ET._Element,
    src_root: ET._Element,
) -> None:
    """
    Объединяет списки ``mc:Ignorable``, чтобы пространства имён расширений,
    используемые титулом (w14, w15, …), оставались игнорируемыми для
    программ, которые их не знают.
    """
    src_value = src_root.get(_MC_IGNORABLE)
    if not src_value:
        return
    uri_to_prefix = {uri: pfx for pfx, uri in root.nsmap.items() if pfx}
    prefixes = (root.get(_MC_IGNORABLE) or "").split()
    for src_prefix in src_value.split():
        uri = src_root.nsmap.get(src_prefix)
        local_prefix = uri_to_prefix.get(uri) if uri else None
        if local_prefix and local_prefix not in prefixes:
            prefixes.append(local_prefix)
    if prefixes:
        root.set(_MC_IGNORABLE, " ".join(prefixes))


def _build_merged_document(
    base_doc_xml: bytes,
    base_name: str,
    leading_elems: list[ET._Element],
    trailing_elems: list[ET._Element],
    final_sect_pr: ET._Element | None,
    src_doc_xml: bytes | None = None,
) -> bytes:
    """
    Собирает document.xml: корень документа, дополненный пространствами имён
    и mc:Ignorable титула; затем содержимое титула, тело документа и итоговый
    sectPr документа.
    """
    root = _fromstring(base_doc_xml)

    if src_doc_xml is not None:
        try:
            src_root    = _fromstring(src_doc_xml)
            extra_nsmap = _resolve_extra_nsmap(
                root.nsmap, _collect_full_nsmap(src_root),
            )
            if extra_nsmap:
                merged_nsmap = {**root.nsmap, **extra_nsmap}
                new_root = ET.Element(root.tag, nsmap=merged_nsmap)
                for attr, val in root.attrib.items():
                    new_root.set(attr, val)
                for child in list(root):
                    new_root.append(child)
                root = new_root
            if _MC_NS in root.nsmap.values():
                _merge_ignorable(root, src_root)
        except Exception as exc:
            logger.warning(
                "Не удалось расширить карту пространств имён из титульного листа: %s",
                exc, exc_info=True,
            )

    body_elem = root.find(f"{{{_WNS}}}body")
    if body_elem is None:
        raise ValueError(
            f"В '{base_name}' отсутствует элемент <w:body>. "
            "Возможно, файл повреждён."
        )

    for child in list(body_elem):
        body_elem.remove(child)

    for el in leading_elems:
        body_elem.append(el)
    for el in trailing_elems:
        body_elem.append(el)
    if final_sect_pr is not None:
        body_elem.append(final_sect_pr)

    return _xml_bytes(root)



def _strip_printer_settings_refs(*element_groups: Iterable[ET._Element]) -> None:
    """Удаляет ссылки на printerSettings: эти части титула не переносятся."""
    tag = f"{{{_WNS}}}printerSettings"
    for group in element_groups:
        for root_el in group:
            if root_el is None:
                continue
            for node in list(root_el.iter(tag)):
                parent = node.getparent()
                if parent is not None:
                    parent.remove(node)


def _make_filler_paragraph(style_id: str | None) -> ET._Element:
    """
    Пустой абзац минимальной высоты, несущий только разрыв раздела, —
    так он не вытолкнет часть титульного листа на лишнюю страницу.
    """
    w = f"{{{_WNS}}}"
    p     = ET.Element(f"{w}p")
    p_pr  = ET.SubElement(p, f"{w}pPr")
    if style_id:
        ET.SubElement(p_pr, f"{w}pStyle").set(f"{w}val", style_id)
    spacing = ET.SubElement(p_pr, f"{w}spacing")
    for attr, val in (("before", "0"), ("after", "0"),
                      ("line", "240"), ("lineRule", "auto")):
        spacing.set(f"{w}{attr}", val)
    mark_rpr = ET.SubElement(p_pr, f"{w}rPr")
    ET.SubElement(mark_rpr, f"{w}sz").set(f"{w}val", "2")
    ET.SubElement(mark_rpr, f"{w}szCs").set(f"{w}val", "2")
    return p


def _embed_sect_pr(
    elems: list[ET._Element],
    sect_pr: ET._Element,
    break_type: str = "continuous",
    filler_style_id: str | None = None,
) -> None:
    """
    Превращает завершающий sectPr титула в разрыв раздела: задаёт тип
    break_type и помещает sectPr в pPr последнего абзаца титула (перед
    pPrChange, как требует схема). Если последний элемент не абзац или уже
    несёт свой sectPr, добавляется пустой абзац-заполнитель.
    """
    w = f"{{{_WNS}}}"
    type_tag = f"{w}type"
    type_el  = sect_pr.find(type_tag)
    if type_el is None:
        type_el = ET.Element(type_tag)
        insert_at = 0
        for i, child in enumerate(sect_pr):
            if _local(child.tag) in ("headerReference", "footerReference",
                                     "footnotePr", "endnotePr"):
                insert_at = i + 1
        sect_pr.insert(insert_at, type_el)
    type_el.set(f"{w}val", break_type)

    last_p = elems[-1] if elems and elems[-1].tag == f"{w}p" else None
    p_pr   = last_p.find(f"{w}pPr") if last_p is not None else None

    if last_p is None or (p_pr is not None and p_pr.find(f"{w}sectPr") is not None):
        if last_p is not None:
            logger.warning(
                "Последний абзац титульного листа уже содержит собственный <w:sectPr> — "
                "для разрыва раздела после титула добавлен пустой абзац."
            )
        last_p = _make_filler_paragraph(filler_style_id)
        elems.append(last_p)
        p_pr = last_p.find(f"{w}pPr")
    elif p_pr is None:
        p_pr = ET.Element(f"{w}pPr")
        last_p.insert(0, p_pr)

    change = p_pr.find(f"{w}pPrChange")
    if change is not None:
        change.addprevious(sect_pr)
    else:
        p_pr.append(sect_pr)



def _patch_num_id_refs(
    elements: Iterable[ET._Element],
    num_remap: dict[int, int],
) -> None:
    """Заменяет ссылки numId по num_remap."""
    if not num_remap:
        return
    num_id_tag = f"{{{_WNS}}}numId"
    w_val      = f"{{{_WNS}}}val"
    for root_el in elements:
        if root_el is None:
            continue
        for node in root_el.iter(num_id_tag):
            val = node.get(w_val)
            if val is None:
                continue
            try:
                old_id = int(val)
            except ValueError:
                continue
            new_id = num_remap.get(old_id)
            if new_id is not None:
                node.set(w_val, str(new_id))


def _collect_numbering_elems_by_id(
    root: ET._Element,
    tag: str,
    id_attr: str,
) -> dict[int, ET._Element]:
    """Элементы numbering.xml с тегом tag по числовому id; нечисловые id пропускаются с предупреждением."""
    by_id: dict[int, ET._Element] = {}
    local_attr = id_attr.rsplit("}", 1)[-1]
    for el in root.findall(f"{{{_WNS}}}{tag}"):
        raw = el.get(id_attr)
        if raw is None:
            continue
        try:
            by_id[int(raw)] = el
        except ValueError:
            logger.warning(
                "Пропущен <w:%s> с нечисловым %s %r в numbering.xml.",
                tag, local_attr, raw,
            )
    return by_id


def _merge_numbering(
    base_numbering_xml: bytes | None,
    src_numbering_xml: bytes | None,
    src_elems: list[ET._Element],
    rename_map: dict[str, str],
    extra_elems: list[ET._Element] | None = None,
) -> tuple[bytes | None, dict[int, int]]:
    """
    Добавляет определения списков титула в numbering.xml документа с id
    после максимальных id документа. Порядок, требуемый схемой, соблюдается:
    numPicBullet — перед abstractNum, abstractNum — перед num. Ссылки numId
    в титуле обновляются.

    Возвращает (XML или None, если у титула нет списков; карту numId).
    """
    if src_numbering_xml is None:
        return None, {}

    src_root = _fromstring(src_numbering_xml)

    w_abstract_num_id_attr = f"{{{_WNS}}}abstractNumId"
    w_num_id_attr          = f"{{{_WNS}}}numId"
    w_pic_bullet_id_attr   = f"{{{_WNS}}}numPicBulletId"
    w_val                  = f"{{{_WNS}}}val"

    s_abstract = _collect_numbering_elems_by_id(
        src_root, "abstractNum", w_abstract_num_id_attr,
    )
    s_nums = _collect_numbering_elems_by_id(src_root, "num", w_num_id_attr)
    s_pic_bullets = _collect_numbering_elems_by_id(
        src_root, "numPicBullet", w_pic_bullet_id_attr,
    )

    if not s_abstract and not s_nums and not s_pic_bullets:
        return None, {}

    if base_numbering_xml is not None:
        base_root   = _fromstring(base_numbering_xml)
        b_abstracts = _collect_numbering_elems_by_id(
            base_root, "abstractNum", w_abstract_num_id_attr,
        )
        b_nums = _collect_numbering_elems_by_id(
            base_root, "num", w_num_id_attr,
        )
        b_pic_bullets = _collect_numbering_elems_by_id(
            base_root, "numPicBullet", w_pic_bullet_id_attr,
        )
        max_abstract   = max(b_abstracts, default=-1)
        max_num        = max((n for n in b_nums if n > 0), default=0)
        max_pic_bullet = max(b_pic_bullets, default=-1)
    else:
        base_root      = ET.Element(f"{{{_WNS}}}numbering", nsmap={"w": _WNS})
        max_abstract   = -1
        max_num        = 0
        max_pic_bullet = -1

    abstract_remap: dict[int, int] = {}
    for old_id in sorted(s_abstract):
        max_abstract += 1
        abstract_remap[old_id] = max_abstract

    num_remap: dict[int, int] = {}
    for old_id in sorted(s_nums):
        max_num += 1
        num_remap[old_id] = max_num

    pic_bullet_remap: dict[int, int] = {}
    for old_id in sorted(s_pic_bullets):
        max_pic_bullet += 1
        pic_bullet_remap[old_id] = max_pic_bullet

    w_num_tag     = f"{{{_WNS}}}num"
    first_num_pos = next(
        (i for i, child in enumerate(base_root) if child.tag == w_num_tag),
        None,
    )
    for offset, old_id in enumerate(sorted(s_abstract)):
        new_el = copy.deepcopy(s_abstract[old_id])
        new_el.set(w_abstract_num_id_attr, str(abstract_remap[old_id]))
        if rename_map:
            _rename_style_refs(new_el, rename_map)
        for pic_ref in new_el.iter(f"{{{_WNS}}}lvlPicBulletId"):
            old_pic_ref = pic_ref.get(w_val)
            if old_pic_ref is None:
                continue
            try:
                parsed_pic_ref = int(old_pic_ref)
            except ValueError:
                logger.warning(
                    "<w:abstractNum w:abstractNumId=\"%s\"> в numbering.xml "
                    "присоединяемого документа содержит нечисловой "
                    "lvlPicBulletId %r; ссылка оставлена без изменений "
                    "и после объединения может указывать не туда.",
                    old_id, old_pic_ref,
                )
                continue
            new_pic_ref = pic_bullet_remap.get(parsed_pic_ref)
            if new_pic_ref is not None:
                pic_ref.set(w_val, str(new_pic_ref))
            else:
                logger.warning(
                    "<w:abstractNum w:abstractNumId=\"%s\"> в numbering.xml "
                    "присоединяемого документа ссылается на numPicBullet %s, "
                    "который в этом документе не определён; ссылка оставлена "
                    "без изменений и после объединения может указывать не туда.",
                    old_id, old_pic_ref,
                )
        if first_num_pos is not None:
            base_root.insert(first_num_pos + offset, new_el)
        else:
            base_root.append(new_el)

    for old_id in sorted(s_nums):
        new_el = copy.deepcopy(s_nums[old_id])
        new_el.set(w_num_id_attr, str(num_remap[old_id]))
        ref = new_el.find(f"{{{_WNS}}}abstractNumId")
        if ref is not None:
            old_ref = ref.get(w_val)
            if old_ref is not None:
                try:
                    parsed_ref = int(old_ref)
                except ValueError:
                    logger.warning(
                        "<w:num w:numId=\"%s\"> в numbering.xml присоединяемого "
                        "документа содержит нечисловой abstractNumId "
                        "%r; ссылка оставлена без изменений и после "
                        "объединения может указывать не туда.",
                        old_id, old_ref,
                    )
                else:
                    new_ref = abstract_remap.get(parsed_ref)
                    if new_ref is not None:
                        ref.set(w_val, str(new_ref))
                    else:
                        logger.warning(
                            "<w:num w:numId=\"%s\"> в numbering.xml "
                            "присоединяемого документа ссылается на "
                            "abstractNumId %s, который в этом документе не "
                            "определён; ссылка оставлена без изменений и "
                            "после объединения может указывать не туда.",
                            old_id, old_ref,
                        )
        base_root.append(new_el)

    if s_pic_bullets:
        w_abstract_num_tag = f"{{{_WNS}}}abstractNum"
        structural_pos = next(
            (
                i for i, child in enumerate(base_root)
                if child.tag in (w_abstract_num_tag, w_num_tag)
            ),
            None,
        )
        for offset, old_id in enumerate(sorted(s_pic_bullets)):
            new_el = copy.deepcopy(s_pic_bullets[old_id])
            new_el.set(w_pic_bullet_id_attr, str(pic_bullet_remap[old_id]))
            if structural_pos is not None:
                base_root.insert(structural_pos + offset, new_el)
            else:
                base_root.append(new_el)

    _patch_num_id_refs(list(src_elems) + list(extra_elems or []), num_remap)

    return _xml_bytes(base_root), num_remap


def _merge_font_table(
    base_font_xml: bytes | None,
    src_font_xml: bytes | None,
) -> bytes | None:
    """
    Добавляет в таблицу шрифтов документа шрифты титула, которых в ней нет
    (имена сравниваются без учёта регистра). Если у документа таблицы нет,
    используется таблица титула. None — таблицу менять не нужно.
    """
    if src_font_xml is None:
        return None
    if base_font_xml is None:
        return src_font_xml

    base_root  = _fromstring(base_font_xml)
    src_root   = _fromstring(src_font_xml)

    w_name = f"{{{_WNS}}}name"
    existing: set[str] = {
        el.get(w_name, "").lower()
        for el in base_root.findall(f"{{{_WNS}}}font")
    }

    added = 0
    for font_el in src_root.findall(f"{{{_WNS}}}font"):
        name = font_el.get(w_name, "").lower()
        if name and name not in existing:
            base_root.append(copy.deepcopy(font_el))
            existing.add(name)
            added += 1

    return _xml_bytes(base_root) if added else None


def _merge_notes(
    base_notes_xml: bytes | None,
    src_notes_xml: bytes | None,
    src_elems: list[ET._Element],
    ref_tag: str,
    rid_map: dict[str, str] | None = None,
    rename_map: dict[str, str] | None = None,
    existing_elems: list[ET._Element] | None = None,
    style_declared_ids: set[str] | None = None,
    style_fallback: str | None = None,
    num_remap: dict[int, int] | None = None,
    dropped_rids: set[str] | None = None,
) -> bytes | None:
    """
    Переносит сноски (или концевые сноски) титула в соответствующую часть
    документа с id после максимального. Если у документа такой части нет,
    она создаётся с обязательными разделителями (id -1 и 0).

    К перенесённым сноскам применяются те же очистки и перенумерации, что
    и к телу титула; ссылки на сноски в титуле обновляются. existing_elems
    пополняется новыми сносками, чтобы следующий вызов учитывал уже
    занятые id. Возвращает XML или None, если у титула нет сносок.
    """
    if src_notes_xml is None:
        return None

    src_root = _fromstring(src_notes_xml)
    note_tag = src_root.tag
    w_id     = f"{{{_WNS}}}id"

    s_notes: dict[int, ET._Element] = {}
    for el in src_root:
        nid_str = el.get(w_id)
        if nid_str is None:
            continue
        try:
            nid = int(nid_str)
        except ValueError:
            continue
        if nid <= 0:
            continue
        s_notes[nid] = el

    if not s_notes:
        return None

    if base_notes_xml is not None:
        base_root = _fromstring(base_notes_xml)
        b_ids: set[int] = set()
        for el in base_root:
            s = el.get(w_id)
            if s is not None:
                with contextlib.suppress(ValueError):
                    b_ids.add(int(s))
        max_id = max((i for i in b_ids if i > 0), default=0)
    else:
        is_footnotes = "footnote" in note_tag.lower()
        item_tag     = f"{{{_WNS}}}{'footnote' if is_footnotes else 'endnote'}"
        base_root    = ET.Element(note_tag, nsmap={"w": _WNS})
        for sep_id, sep_type, mark_local in (
            ("-1", "separator",             "separator"),
            ("0",  "continuationSeparator", "continuationSeparator"),
        ):
            sep_el = ET.SubElement(base_root, item_tag)
            sep_el.set(f"{{{_WNS}}}type", sep_type)
            sep_el.set(w_id, sep_id)
            p_el = ET.SubElement(sep_el, f"{{{_WNS}}}p")
            r_el = ET.SubElement(p_el,   f"{{{_WNS}}}r")
            ET.SubElement(r_el, f"{{{_WNS}}}{mark_local}")
        max_id = 0

    id_remap: dict[int, int] = {}
    for old_id in sorted(s_notes):
        max_id += 1
        id_remap[old_id] = max_id

    new_els: list[ET._Element] = []
    for old_id, el in s_notes.items():
        new_el = copy.deepcopy(el)
        new_el.set(w_id, str(id_remap[old_id]))
        _strip_dropped_rel_refs([new_el], dropped_rids or set())
        _strip_comment_anchors([new_el])
        if rid_map:
            _rename_rids([new_el], rid_map)
        new_els.append(new_el)

    if rename_map:
        for new_el in new_els:
            _rename_style_refs(new_el, rename_map)
    base_for_ids = existing_elems if existing_elems is not None else []
    _remap_drawing_ids(new_els, base_for_ids)
    note_name_map = _remap_bookmark_ids(new_els, base_for_ids)
    _rename_bookmark_field_refs(new_els, note_name_map)
    _resolve_undeclared_style_refs(
        new_els,
        declared_ids=style_declared_ids or set(),
        paragraph_fallback=style_fallback,
    )
    if num_remap:
        _patch_num_id_refs(new_els, num_remap)

    for new_el in new_els:
        base_root.append(new_el)

    if existing_elems is not None:
        existing_elems.extend(new_els)

    for root_el in src_elems:
        for node in root_el.iter(ref_tag):
            val = node.get(w_id)
            if val is not None:
                try:
                    new_n = id_remap.get(int(val))
                    if new_n is not None:
                        node.set(w_id, str(new_n))
                except ValueError:
                    pass

    return _xml_bytes(base_root)



def _write_zip(
    base_path: Path,
    output_path: Path,
    replacements: dict[str, bytes],
    extra_files: dict[str, bytes] | None = None,
) -> None:
    """
    Записывает архив: части base_path (с заменами из replacements), затем
    новые части и extra_files. Запись идёт во временный файл, который
    затем заменяет output_path; при любой ошибке временный файл удаляется.
    """
    tmp = output_path.with_name(f"{output_path.name}.{uuid.uuid4().hex}.tmp")
    try:
        with (
            zipfile.ZipFile(base_path,  "r")                          as base_zip,
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)           as out_zip,
        ):
            written: set[str] = set()
            for item in base_zip.infolist():
                if item.filename in replacements:
                    out_zip.writestr(item, replacements[item.filename])
                else:
                    out_zip.writestr(item, base_zip.read(item.filename))
                written.add(item.filename)
            for zip_path, data in replacements.items():
                if zip_path not in written:
                    out_zip.writestr(zip_path, data)
            if extra_files:
                for zip_path, data in extra_files.items():
                    out_zip.writestr(zip_path, data)
        tmp.replace(output_path)
    except BaseException:
        if tmp.exists():
            tmp.unlink()
        raise


def _resolve_zip_path(target: str) -> str:
    """Путь части в архиве по атрибуту Target связи (относительно word/ или абсолютный)."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(f"word/{target}")


def _compute_rels_remap(
    fname_map: dict[str, str],
    src_rels_xml: bytes | None,
    base_rels_xml: bytes | None,
    s_header_footer: dict[str, bytes] | None = None,
) -> tuple[
    dict[str, str],
    list[tuple[str, str, str, str | None]],
    dict[str, str],
    dict[str, str],
]:
    """
    Определяет, какие связи титула переносятся и с какими новыми rId:
    внешние ссылки и медиафайлы — всегда, колонтитулы — под новыми
    уникальными именами файлов. Общие части (_SKIPPABLE_REL_SUFFIXES)
    и примечания пропускаются, прочие встроенные объекты — с предупреждением.

    Возвращает (карту rId, новые записи .rels, карту имён колонтитулов,
    типы содержимого колонтитулов).
    """
    if src_rels_xml is None:
        return {}, [], {}, {}

    src_rels = _parse_rels(src_rels_xml)
    if not src_rels:
        return {}, [], {}, {}

    s_header_footer = s_header_footer or {}

    out_rels     = _parse_rels(base_rels_xml) if base_rels_xml else {}
    used_ids: set[str] = set(out_rels)
    next_rid_num = _max_rid(out_rels) + 1

    hf_fname_map:      dict[str, str] = {}
    hf_content_types:  dict[str, str] = {}
    rid_map:      dict[str, str] = {}
    new_entries:  list[tuple[str, str, str, str | None]] = []

    for old_rid, (rtype, target, target_mode) in src_rels.items():
        if target_mode == "External":
            new_rid, next_rid_num = _next_free_rid(used_ids, next_rid_num)
            used_ids.add(new_rid)
            rid_map[old_rid] = new_rid
            new_entries.append((new_rid, rtype, target, target_mode))
            continue

        if any(rtype.endswith(suffix) for suffix in _SKIPPABLE_REL_SUFFIXES):
            continue

        if (
            rtype.endswith(_HEADER_REL_SUFFIX)
            or rtype.endswith(_FOOTER_REL_SUFFIX)
        ):
            old_zip_path = _resolve_zip_path(target)
            if old_zip_path not in s_header_footer:
                logger.warning(
                    "Цель связи колонтитула '%s' отсутствует "
                    "в исходном архиве — пропущено.",
                    target,
                )
                continue
            new_zip_path = hf_fname_map.get(old_zip_path)
            if new_zip_path is None:
                new_basename = (
                    f"{Path(old_zip_path).stem}_{uuid.uuid4().hex}.xml"
                )
                new_zip_path = f"word/{new_basename}"
                hf_fname_map[old_zip_path] = new_zip_path
                hf_content_types[new_zip_path] = (
                    _HEADER_CT_TYPE if rtype.endswith(_HEADER_REL_SUFFIX)
                    else _FOOTER_CT_TYPE
                )
            new_basename = new_zip_path[len("word/"):]

            new_rid, next_rid_num = _next_free_rid(used_ids, next_rid_num)
            used_ids.add(new_rid)
            rid_map[old_rid] = new_rid
            new_entries.append((new_rid, rtype, new_basename, None))
            continue

        is_media = (
            any(target.lower().endswith(f".{ext}") for ext in _MEDIA_EXTENSIONS)
            or "image" in rtype.lower()
        )
        if is_media:
            old_zip_path = _resolve_zip_path(target)
            resolved = (
                fname_map.get(old_zip_path)
                or fname_map.get("word/" + old_zip_path)
            )
            if resolved is None:
                logger.warning(
                    "Изображение из присоединяемого документа (target: %s) не найдено "
                    "среди скопированных медиафайлов — ссылка на него в итоговом "
                    "документе может оказаться битой.",
                    target,
                )
            new_zip_path = resolved or old_zip_path

            new_rid, next_rid_num = _next_free_rid(used_ids, next_rid_num)
            used_ids.add(new_rid)
            rid_map[old_rid] = new_rid
            new_entries.append(
                (new_rid, rtype, new_zip_path[len("word/"):], None)
            )
            continue

        if any(rtype.endswith(suffix) for suffix in _COMMENT_REL_SUFFIXES):
            continue

        logger.warning(
            "Неподдерживаемый встроенный объект пропущен "
            "(тип связи: %s, target: %s). "
            "Удалите этот объект из документа или вставьте его как обычное "
            "изображение, чтобы он попал в итоговый файл.",
            rtype.rsplit("/", 1)[-1],
            target,
        )

    return rid_map, new_entries, hf_fname_map, hf_content_types



def _parse_rels(data: bytes) -> dict[str, tuple[str, str, str | None]]:
    """Связи из .rels: rId → (тип, цель, TargetMode). Неполные записи пропускаются с предупреждением."""
    root   = _fromstring(data)
    result: dict[str, tuple[str, str, str | None]] = {}
    for el in root.findall(f"{{{_REL_NS}}}Relationship"):
        rid, rtype, target = el.get("Id"), el.get("Type"), el.get("Target")
        if None in (rid, rtype, target):
            logger.warning(
                "Пропущена повреждённая запись Relationship: %s",
                ET.tostring(el, encoding="unicode"),
            )
            continue
        target_mode = el.get("TargetMode")
        result[rid] = (rtype, target, target_mode)
    return result


def _max_rid(rels: dict[str, tuple[str, str, str | None]]) -> int:
    """Наибольший номер среди rId вида rIdN."""
    max_n = 0
    for rid in rels:
        m = re.match(r"^rId(\d+)$", rid)
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n


def _next_free_rid(existing_ids: set[str], next_n: int) -> tuple[str, int]:
    """Первый свободный rIdN начиная с next_n; возвращает его и номер, с которого продолжать поиск."""
    candidate = f"rId{next_n}"
    while candidate in existing_ids:
        next_n += 1
        candidate = f"rId{next_n}"
    return candidate, next_n + 1


def _build_rels_xml(
    existing: dict[str, tuple[str, str, str | None]],
    new_entries: list[tuple[str, str, str, str | None]],
) -> bytes:
    """Собирает .rels из существующих и новых записей."""
    root = ET.Element("Relationships", nsmap={None: _REL_NS})
    for rid, (rtype, target, target_mode) in existing.items():
        attrs: dict[str, str] = {"Id": rid, "Type": rtype, "Target": target}
        if target_mode:
            attrs["TargetMode"] = target_mode
        ET.SubElement(root, "Relationship", **attrs)
    for new_rid, rtype, target, target_mode in new_entries:
        attrs = {"Id": new_rid, "Type": rtype, "Target": target}
        if target_mode:
            attrs["TargetMode"] = target_mode
        ET.SubElement(root, "Relationship", **attrs)
    return _xml_bytes(root)


def _add_ct_override(ct_xml: bytes, part_name: str, content_type: str) -> bytes:
    """Добавляет в [Content_Types].xml Override для части, если его ещё нет."""
    root     = _fromstring(ct_xml)
    existing = {el.get("PartName") for el in root.findall(f"{{{_CT_NS}}}Override")}
    if part_name not in existing:
        ET.SubElement(
            root,
            f"{{{_CT_NS}}}Override",
            PartName=part_name,
            ContentType=content_type,
        )
    return _xml_bytes(root)


def _merge_content_types(
    out_ct_xml: bytes,
    src_ct_xml: bytes | None,
    src_media_names: list[str],
    fname_map: dict[str, str],
) -> bytes:
    """
    Регистрирует в [Content_Types].xml медиафайлы титула: Default для их
    расширений (из [Content_Types].xml титула или по таблице _MIME_FALLBACKS)
    и Override для медиафайлов под новыми именами.
    """
    root = _fromstring(out_ct_xml)
    existing_exts: set[str] = {
        el.get("Extension", "").lower()
        for el in root.findall(f"{{{_CT_NS}}}Default")
    }
    existing_parts: set[str] = {
        el.get("PartName", "")
        for el in root.findall(f"{{{_CT_NS}}}Override")
    }

    media_exts: set[str] = {
        posixpath.splitext(p)[1].lstrip(".").lower()
        for p in src_media_names
        if posixpath.splitext(p)[1]
    }

    if src_ct_xml is None:
        logger.warning("В присоединяемом документе нет [Content_Types].xml.")
    else:
        src_ct = _fromstring(src_ct_xml)

        for default in src_ct.findall(f"{{{_CT_NS}}}Default"):
            ext = default.get("Extension", "").lower()
            if (
                (ext in _IMAGE_EXTENSIONS or ext in media_exts)
                and ext not in existing_exts
            ):
                new_node = copy.deepcopy(default)
                new_node.set("Extension", ext)
                root.append(new_node)
                existing_exts.add(ext)

        for override in src_ct.findall(f"{{{_CT_NS}}}Override"):
            part = override.get("PartName", "")
            if part.startswith("/word/media/"):
                old_key  = part.lstrip("/")
                new_part = (
                    f"/{fname_map[old_key]}" if old_key in fname_map else part
                )
                if new_part not in existing_parts:
                    new_node = copy.deepcopy(override)
                    new_node.set("PartName", new_part)
                    root.append(new_node)
                    existing_parts.add(new_part)

    for src_path in src_media_names:
        ext = posixpath.splitext(src_path)[1].lstrip(".").lower()
        if ext and ext not in existing_exts:
            mime = _MIME_FALLBACKS.get(ext, _GENERIC_MEDIA_CT_TYPE)
            ET.SubElement(
                root,
                f"{{{_CT_NS}}}Default",
                Extension=ext,
                ContentType=mime,
            )
            existing_exts.add(ext)

    return _xml_bytes(root)
  

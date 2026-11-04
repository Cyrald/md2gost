#!/usr/bin/env python3
"""
Генератор GOST_7.32-2017.docx — эталонного документа (reference-doc) для pandoc.

Собирает .docx с нуля и сохраняет в папку профиля
(resources/styles/GOST_7.32-2017/GOST_7.32-2017.docx) — это файл, на который
ссылается reference_docx в config.cfg. Параметры оформления взяты
из ГОСТ 7.32-2017; отступления от него отмечены в комментариях.

Запуск: python3 resources/styles/GOST_7.32-2017/tools/build_gost.py
"""
import zipfile
from pathlib import Path

# ── Перевод единиц в единицы WordprocessingML ────────────────────────────────
def mm(v):   return round(v * 1440 / 25.4)   # мм → твипы (1/20 пункта)
def hp(v):   return round(v * 2)              # пт → полупункты (размер шрифта, w:sz)
def twp(v):  return round(v * 20)             # пт → твипы (интервалы и отступы)

# ── ГОСТ 7.32-2017: размеры страницы и поля (§4.2) ──────────────────────────
PG_W  = mm(210)   # A4
PG_H  = mm(297)
ML    = mm(30)    # левое  ≥ 30 мм
MR    = mm(15)    # правое поле — 15 мм по инструкции проекта
MT    = mm(20)    # верхнее ≥ 20 мм
MB    = mm(20)    # нижнее ≥ 20 мм
TEXT_W = PG_W - ML - MR   # ширина области текста (165 мм) — позиция правой табуляции в TOC


# ── ГОСТ 7.32-2017: шрифт и интервалы (§4.1) ─────────────────────────────────
FONT  = "Times New Roman"   # рекомендован
MONO  = "Courier New"       # для кода
SZ14  = hp(14)              # кегль 14 → 28 полупунктов
SZ12  = hp(12)              # для подписей и сносок → 24
LS    = 360                 # полуторный межстрочный: 1.5 × 240 = 360
IND   = mm(12.5)            # абзацный отступ 1.25 см → ≈ 709 twips

# ── Заголовочные отступы (§7.2) ───────────────────────────────────────────────
# «не менее двух межстрочных интервалов» до, «не менее одного» после
# 14 пт × 1.5 ≈ 21 пт на строку
SP_B2 = twp(42)   # 2 строки перед заголовком
SP_A1 = twp(21)   # 1 строка после заголовка

# ── Отступ перечислений (§7.3) ────────────────────────────────────────────────
# Перечисления записываются с абзацным отступом
LIST_IND  = IND           # левый отступ = абзацный
LIST_HANG = mm(6.25)      # выступ (половина абзацного для вложенных)


# ─────────────────────────────────────────────────────────────────────────────
# Части пакета .docx (XML)
# ─────────────────────────────────────────────────────────────────────────────

RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
    Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
    Target="docProps/app.xml"/>
</Relationships>
"""

CORE = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>ГОСТ 7.32-2017 Reference Document</dc:title>
  <dc:description>Pandoc reference doc. ГОСТ 7.32-2017.</dc:description>
</cp:coreProperties>
"""

APP = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>md2gost GOST Template Builder</Application>
</Properties>
"""

FONT_TABLE = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="{FONT}">
    <w:panose1 w:val="02020603050405020304"/>
    <w:charset w:val="CC"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="variable"/>
  </w:font>
  <w:font w:name="{MONO}">
    <w:panose1 w:val="02070309020205020404"/>
    <w:charset w:val="CC"/>
    <w:family w:val="modern"/>
    <w:pitch w:val="fixed"/>
  </w:font>
</w:fonts>
"""

FOOTNOTES = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:footnote>
</w:footnotes>
"""

ENDNOTES = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:endnote>
</w:endnotes>
"""

SETTINGS = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="709"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
      <w:doNotExpandShiftReturn/>
      <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
      <w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
      <w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
      <w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
  </w:compat>
</w:settings>
"""

# ── Numbering: §7.3 перечисления ─────────────────────────────────────────────
# abstractNumId=1 — маркированный список: уровень 0 — тире (—), уровень 1 —
#   строчная русская буква со скобкой (russianLower), уровень 2 — цифра со скобкой;
# abstractNumId=2 — нумерованный список: уровень 0 — цифра со скобкой,
#   уровень 1 — строчная русская буква со скобкой.
NUMBERING = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <!-- abstractNum 1: маркированный список (§7.3 — дефис) -->
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2014;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="{IND + LIST_HANG}" w:hanging="{LIST_HANG}"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>
        <w:sz w:val="{SZ14}"/>
        <w:szCs w:val="{SZ14}"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="russianLower"/>
      <w:lvlText w:val="%2)"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="{IND * 2 + LIST_HANG}" w:hanging="{LIST_HANG}"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>
        <w:sz w:val="{SZ14}"/>
        <w:szCs w:val="{SZ14}"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%3)"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="{IND * 3 + LIST_HANG}" w:hanging="{LIST_HANG}"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>
        <w:sz w:val="{SZ14}"/>
        <w:szCs w:val="{SZ14}"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>

  <!-- abstractNum 2: нумерованный список (§7.3 — цифры со скобкой) -->
  <w:abstractNum w:abstractNumId="2">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1)"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="{IND + LIST_HANG}" w:hanging="{LIST_HANG}"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>
        <w:sz w:val="{SZ14}"/>
        <w:szCs w:val="{SZ14}"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="russianLower"/>
      <w:lvlText w:val="%2)"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="{IND * 2 + LIST_HANG}" w:hanging="{LIST_HANG}"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>
        <w:sz w:val="{SZ14}"/>
        <w:szCs w:val="{SZ14}"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>

  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>
"""


# ── styles.xml ────────────────────────────────────────────────────────────────
def make_styles():
    """Возвращает word/styles.xml со всеми стилями, которые использует pandoc и Lua-фильтр профиля."""
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns = f'xmlns:w="{W}"'

    # Вспомогательные функции собирают фрагменты w:rPr, w:spacing и w:ind.
    def rpr(sz=None, bold=False, italic=False, caps=False, mono=False, underline=False,
            color="auto", sz_cs=None, kern=None):
        f = MONO if mono else FONT
        s = sz or SZ14
        sc = sz_cs or s
        out = f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>'
        if bold:      out += "<w:b/><w:bCs/>"
        if italic:    out += "<w:i/><w:iCs/>"
        if caps:      out += "<w:caps/>"
        if underline: out += '<w:u w:val="single"/>'
        if kern:      out += f'<w:kern w:val="{kern}"/>'
        if color != "auto": out += f'<w:color w:val="{color}"/>'
        out += f'<w:sz w:val="{s}"/><w:szCs w:val="{sc}"/>'
        return out

    def spacing(before=0, after=0, ls=LS):
        parts = []
        if before: parts.append(f'w:before="{before}"')
        if after:  parts.append(f'w:after="{after}"')
        parts.append(f'w:line="{ls}" w:lineRule="auto"')
        return f'<w:spacing {" ".join(parts)}/>'

    def ind(left=None, right=None, first=None, hanging=None):
        parts = []
        if left    is not None: parts.append(f'w:left="{left}"')
        if right   is not None: parts.append(f'w:right="{right}"')
        if first   is not None: parts.append(f'w:firstLine="{first}"')
        if hanging is not None: parts.append(f'w:hanging="{hanging}"')
        return f'<w:ind {" ".join(parts)}/>' if parts else ""

    # ── §4.1: Normal ──────────────────────────────────────────────────────────
    # Times New Roman 14пт, полуторный, выравнивание по ширине, отступ 1.25 см
    s_normal = f"""
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr>
      <w:widowControl/>
      {spacing()}
      {ind(first=IND)}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── Default Paragraph Font (character) ────────────────────────────────────
    s_dpf = f"""
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/>
    <w:uiPriority w:val="1"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>"""

    # ── §7.2: заголовки разделов (Heading 1) ──────────────────────────────────
    # «с абзацным отступом с прописной буквы», полужирный
    # §7.2: до — ≥2 межстрочных интервала, после — ≥1
    # §4.4: каждый раздел рекомендуется начинать с новой страницы
    s_h1 = f"""
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:pageBreakBefore/>
      {spacing(before=SP_B2, after=SP_A1)}
      {ind(first=IND)}
      <w:jc w:val="both"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>{rpr(bold=True)}</w:rPr>
  </w:style>"""

    # ── §7.2: заголовки подразделов (Heading 2) ───────────────────────────────
    s_h2 = f"""
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      {spacing(before=SP_B2, after=SP_A1)}
      {ind(first=IND)}
      <w:jc w:val="both"/>
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>{rpr(bold=True)}</w:rPr>
  </w:style>"""

    # ── §7.2: заголовки пунктов (Heading 3) ───────────────────────────────────
    s_h3 = f"""
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      {spacing(before=SP_B2, after=SP_A1)}
      {ind(first=IND)}
      <w:jc w:val="both"/>
      <w:outlineLvl w:val="2"/>
    </w:pPr>
    <w:rPr>{rpr(bold=True, italic=True)}</w:rPr>
  </w:style>"""

    # ── §7.2: заголовки подпунктов (Heading 4) ────────────────────────────────
    s_h4 = f"""
  <w:style w:type="paragraph" w:styleId="Heading4">
    <w:name w:val="heading 4"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      {spacing(before=SP_B2, after=SP_A1)}
      {ind(first=IND)}
      <w:jc w:val="both"/>
      <w:outlineLvl w:val="3"/>
    </w:pPr>
    <w:rPr>{rpr(italic=True)}</w:rPr>
  </w:style>"""

    # Heading 5-6 (резервные)
    s_h5 = f"""
  <w:style w:type="paragraph" w:styleId="Heading5">
    <w:name w:val="heading 5"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      {spacing(before=SP_B2, after=SP_A1)}
      {ind(first=IND)}
      <w:jc w:val="both"/>
      <w:outlineLvl w:val="4"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    s_h6 = f"""
  <w:style w:type="paragraph" w:styleId="Heading6">
    <w:name w:val="heading 6"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      {spacing(before=SP_B2, after=SP_A1)}
      {ind(first=IND)}
      <w:jc w:val="both"/>
      <w:outlineLvl w:val="5"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── §6.1: Title (наименование НИР на титульном листе) ────────────────────
    # Титульный лист: наименование НИР — по центру прописными буквами
    s_title = f"""
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      {spacing(before=0, after=twp(14))}
      {ind(first=0)}
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>{rpr(bold=True, caps=True)}</w:rPr>
  </w:style>"""

    # Подзаголовок (Subtitle) — по центру, курсив
    s_subtitle = f"""
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      {spacing(before=0, after=twp(14))}
      {ind(first=0)}
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>{rpr(italic=True)}</w:rPr>
  </w:style>"""

    # Автор и дата (Author, Date) — по центру
    s_author = f"""
  <w:style w:type="paragraph" w:styleId="Author">
    <w:name w:val="Author"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      {spacing(before=0, after=twp(7))}
      {ind(first=0)}
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    s_date = f"""
  <w:style w:type="paragraph" w:styleId="Date">
    <w:name w:val="Date"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      {spacing(before=0, after=twp(7))}
      {ind(first=0)}
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── §6.3: Реферат (Abstract) ──────────────────────────────────────────────
    # §6.3: реферат — не более одной страницы, ключевые слова прописными
    s_abstract = f"""
  <w:style w:type="paragraph" w:styleId="Abstract">
    <w:name w:val="Abstract"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      {spacing()}
      {ind(first=IND)}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    s_abstract_title = f"""
  <w:style w:type="paragraph" w:styleId="AbstractTitle">
    <w:name w:val="AbstractTitle"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Abstract"/>
    <w:pPr>
      {spacing(before=0, after=SP_A1)}
      {ind(first=0)}
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>{rpr(bold=True, caps=True)}</w:rPr>
  </w:style>"""

    # ── First Paragraph (первый абзац после заголовка) ───────────────────────
    # pandoc генерирует этот стиль для первого абзаца после заголовка;
    # §4.1: абзацный отступ единый по всему тексту, поэтому как у Normal
    s_first_par = f"""
  <w:style w:type="paragraph" w:styleId="FirstParagraph">
    <w:name w:val="First Paragraph"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      {spacing()}
      {ind(first=IND)}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # Body Text — синоним Normal. pandoc применяет его ко всем обычным абзацам,
    # поэтому дополнительных интервалов до/после быть не должно
    s_body_text = f"""
  <w:style w:type="paragraph" w:styleId="BodyText">
    <w:name w:val="Body Text"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      {spacing()}
      {ind(first=IND)}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # Compact — плотный список (без отступа первой строки, без межабзацного)
    s_compact = f"""
  <w:style w:type="paragraph" w:styleId="Compact">
    <w:name w:val="Compact"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Compact"/>
    <w:pPr>
      {spacing(ls=LS)}
      {ind(first=0)}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── §7.3: List Paragraph ──────────────────────────────────────────────────
    # Перечисления с абзацным отступом
    s_list_par = f"""
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="ListParagraph"/>
    <w:uiPriority w:val="34"/>
    <w:qFormat/>
    <w:pPr>
      {spacing()}
      {ind(left=IND, first=0)}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── Block Text (цитата / выделенный блок) ────────────────────────────────
    # ГОСТ 7.32 не регламентирует блок-цитаты; отступы слева и справа
    s_block = f"""
  <w:style w:type="paragraph" w:styleId="BlockText">
    <w:name w:val="Block Text"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      {spacing()}
      {ind(left=IND, right=IND, first=0)}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── Source Code (листинг кода) ────────────────────────────────────────────
    # Моноширинный, кегль 12, одинарный интервал, без отступа
    s_code = f"""
  <w:style w:type="paragraph" w:styleId="SourceCode">
    <w:name w:val="Source Code"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="SourceCode"/>
    <w:pPr>
      {spacing(ls=240)}
      {ind(first=0)}
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>{rpr(mono=True, sz=SZ12, sz_cs=SZ12)}</w:rPr>
  </w:style>"""

    # Verbatim Char — встроенный код (стиль знака)
    s_verbatim_char = f"""
  <w:style w:type="character" w:styleId="VerbatimChar">
    <w:name w:val="Verbatim Char"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:rPr>{rpr(mono=True, sz=SZ12, sz_cs=SZ12)}</w:rPr>
  </w:style>"""

    # ── §8: абзац с иллюстрацией ─────────────────────────────────────────────
    # pandoc 3.x пишет pStyle w:val="CaptionedFigure" для абзаца с картинкой
    # §8: без абзацного отступа, по центру
    s_captioned_figure = f"""
  <w:style w:type="paragraph" w:styleId="CaptionedFigure">
    <w:name w:val="CaptionedFigure"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>

      {spacing(before=twp(6), after=0)}
      {ind(left=0, first=0)}
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # FigureParagraph — стиль, который Lua-фильтр профиля назначает абзацу
    # с изображением без подписи (у таких абзацев нет CaptionedFigure)
    s_figure_par = f"""
  <w:style w:type="paragraph" w:styleId="FigureParagraph">
    <w:name w:val="FigureParagraph"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>

      {spacing(before=twp(6), after=twp(6))}
      {ind(left=0, first=0)}
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── §8: подписи к иллюстрациям (Caption / Image Caption) ─────────────────
    # §8: «Слово "Рисунок" и его наименование помещаются после пояснительных
    # данных по центру»; кегль можно оставить основной — ГОСТ не уменьшает
    s_caption = f"""
  <w:style w:type="paragraph" w:styleId="Caption">
    <w:name w:val="Caption"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      {spacing(before=twp(6), after=twp(6))}
      {ind(left=0, first=0)}
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    s_img_caption = f"""
  <w:style w:type="paragraph" w:styleId="ImageCaption">
    <w:name w:val="Image Caption"/>
    <w:basedOn w:val="Caption"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      {spacing(before=twp(6), after=twp(6))}
      {ind(left=0, first=0)}
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── §9: подписи к таблицам (Table Caption) ────────────────────────────────
    # §9: «наименование таблицы... располагается над таблицей слева без
    # абзацного отступа»; перед наименованием таблицы — отступ в одну строку
    s_tbl_caption = f"""
  <w:style w:type="paragraph" w:styleId="TableCaption">
    <w:name w:val="Table Caption"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>

      {spacing(before=SP_A1, after=twp(6))}
      {ind(first=0)}
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── §9: Table / Table Paragraph ───────────────────────────────────────────
    # §9: «Таблицы ограничиваются линиями» — все 6 границ одинарные;
    # текст в ячейках — тот же шрифт и кегль, что и основной текст.
    # Минимальная высота строки ГОСТ 7.32-2017 не регламентируется
    # (требование «не менее 8 мм» относится к ГОСТ 2.105), поэтому не задаётся
    s_table = f"""
  <w:style w:type="table" w:styleId="Table">
    <w:name w:val="Table"/>
    <w:uiPriority w:val="99"/>
    <w:unhideWhenUsed/>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblBorders>
        <w:top     w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:left    w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:bottom  w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:right   w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      </w:tblBorders>
      <w:tblCellMar>
        <w:top    w:w="{twp(1)}" w:type="dxa"/>
        <w:left   w:w="{mm(1.9)}" w:type="dxa"/>
        <w:bottom w:w="{twp(1)}" w:type="dxa"/>
        <w:right  w:w="{mm(1.9)}" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>"""

    s_table_par = f"""
  <w:style w:type="paragraph" w:styleId="TableParagraph">
    <w:name w:val="Table Paragraph"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="TableParagraph"/>
    <w:pPr>
      {spacing(ls=240)}
      {ind(first=0)}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── §10: нумерованная формула (Equation) ──────────────────────────────────
    # Применяется через eqnBlockTemplate в GOST_7.32-2017.yaml: строка
    # «<tab>формула<tab>(номер)». Формула — по центру, номер — на уровне
    # формулы справа; выше и ниже — не менее одной свободной строки
    s_equation = f"""
  <w:style w:type="paragraph" w:styleId="Equation">
    <w:name w:val="Equation"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      <w:keepLines/>
      <w:tabs>
        <w:tab w:val="center" w:pos="{TEXT_W // 2}"/>
        <w:tab w:val="right" w:pos="{TEXT_W}"/>
      </w:tabs>
      {spacing(before=SP_A1, after=SP_A1)}
      {ind(left=0, right=0, first=0)}
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── §7.4: сноски (Footnote Text) ─────────────────────────────────────────
    # §7.4: «сноски к тексту печатаются с абзацного отступа»; отделяются линией
    # Кегль 12 (не менее 12, §4.1)
    s_fn_text = f"""
  <w:style w:type="paragraph" w:styleId="FootnoteText">
    <w:name w:val="footnote text"/>
    <w:basedOn w:val="Normal"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:pPr>
      {spacing(ls=240)}
      {ind(first=IND)}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>{rpr(sz=SZ12, sz_cs=SZ12)}</w:rPr>
  </w:style>"""

    s_fn_text_char = f"""
  <w:style w:type="character" w:styleId="FootnoteTextChar">
    <w:name w:val="Footnote Text Char"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:semiHidden/>
    <w:rPr>{rpr(sz=SZ12, sz_cs=SZ12)}</w:rPr>
  </w:style>"""

    # §7.4: «знак сноски — арабская цифра или звёздочка»
    s_fn_ref = f"""
  <w:style w:type="character" w:styleId="FootnoteReference">
    <w:name w:val="footnote reference"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:rPr>
      <w:vertAlign w:val="superscript"/>
    </w:rPr>
  </w:style>"""

    # ── §6.11: список использованных источников (Bibliography) ───────────────
    # Источники записываются по ГОСТ 7.1 / ГОСТ Р 7.0.5
    s_bib = f"""
  <w:style w:type="paragraph" w:styleId="Bibliography">
    <w:name w:val="Bibliography"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Bibliography"/>
    <w:qFormat/>
    <w:pPr>
      {spacing()}
      {ind(left=IND, first=-IND)}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── §6.7: Определения / Definition ───────────────────────────────────────
    s_def_term = f"""
  <w:style w:type="paragraph" w:styleId="DefinitionTerm">
    <w:name w:val="Definition Term"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Definition"/>
    <w:pPr>
      {spacing()}
      {ind(first=IND)}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>{rpr(bold=True)}</w:rPr>
  </w:style>"""

    s_def = f"""
  <w:style w:type="paragraph" w:styleId="Definition">
    <w:name w:val="Definition"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      {spacing()}
      {ind(left=IND, first=0)}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── Заголовок «СОДЕРЖАНИЕ» (pandoc ставит pStyle "TOCHeading") ───────────
    # Структурный элемент: по центру, прописными, полужирный, без отступа,
    # без номера и без попадания в само содержание (outlineLvl 9 = «текст»)
    s_toc_heading = f"""
  <w:style w:type="paragraph" w:styleId="TOCHeading">
    <w:name w:val="TOC Heading"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="39"/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      {spacing(before=0, after=SP_A1)}
      {ind(left=0, first=0)}
      <w:jc w:val="center"/>
      <w:outlineLvl w:val="9"/>
    </w:pPr>
    <w:rPr>{rpr(bold=True, caps=True)}</w:rPr>
  </w:style>"""

    # ── §6.4: Содержание (TOC) ────────────────────────────────────────────────
    # «содержание... с указанием номеров страниц»
    s_toc1 = f"""
  <w:style w:type="paragraph" w:styleId="TOC1">
  <w:name w:val="toc 1"/>
  <w:basedOn w:val="Normal"/>
  <w:next w:val="Normal"/>
  <w:uiPriority w:val="39"/>
  <w:unhideWhenUsed/>
  <w:pPr>
    <w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>
    <w:ind w:left="0" w:firstLine="0"/>
    <w:jc w:val="both"/>
    <w:tabs>
      <w:tab w:val="right" w:leader="dot" w:pos="{TEXT_W}"/>
    </w:tabs>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
    <w:sz w:val="28"/>
    <w:szCs w:val="28"/>
    <w:color w:val="000000"/>
  </w:rPr>
</w:style>"""

    s_toc2 = f"""
  <w:style w:type="paragraph" w:styleId="TOC2">
  <w:name w:val="toc 2"/>
  <w:basedOn w:val="Normal"/>
  <w:next w:val="Normal"/>
  <w:uiPriority w:val="39"/>
  <w:unhideWhenUsed/>
  <w:pPr>
    <w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>
    <w:ind w:left="709" w:firstLine="0"/>
    <w:jc w:val="both"/>
    <w:tabs>
      <w:tab w:val="right" w:leader="dot" w:pos="{TEXT_W}"/>
    </w:tabs>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
    <w:sz w:val="28"/>
    <w:szCs w:val="28"/>
    <w:color w:val="000000"/>
  </w:rPr>
</w:style>"""

    s_toc3 = f"""
  <w:style w:type="paragraph" w:styleId="TOC3">
  <w:name w:val="toc 3"/>
  <w:basedOn w:val="Normal"/>
  <w:next w:val="Normal"/>
  <w:uiPriority w:val="39"/>
  <w:unhideWhenUsed/>
  <w:pPr>
    <w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>
    <w:ind w:left="1418" w:firstLine="0"/>
    <w:jc w:val="both"/>
    <w:tabs>
      <w:tab w:val="right" w:leader="dot" w:pos="{TEXT_W}"/>
    </w:tabs>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
    <w:sz w:val="28"/>
    <w:szCs w:val="28"/>
    <w:color w:val="000000"/>
  </w:rPr>
</w:style>"""

    # ── Hyperlink (стиль знака): чёрный с подчёркиванием ─────────────────────
    s_hyperlink = f"""
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="99"/>
    <w:unhideWhenUsed/>
    <w:rPr>
      <w:color w:val="000000"/>
      <w:u w:val="single"/>
    </w:rPr>
  </w:style>"""

    # ── §4.4: нумерация страниц — нижний центр (Header / Footer) ─────────────
    s_header = f"""
  <w:style w:type="paragraph" w:styleId="Header">
    <w:name w:val="header"/>
    <w:basedOn w:val="Normal"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:pPr>
      {spacing(ls=240)}
      {ind(first=0)}
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    s_footer = f"""
  <w:style w:type="paragraph" w:styleId="Footer">
    <w:name w:val="footer"/>
    <w:basedOn w:val="Normal"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:pPr>
      {spacing(ls=240)}
      {ind(first=0)}
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>{rpr()}</w:rPr>
  </w:style>"""

    # ── docDefaults — база для всех стилей ────────────────────────────────────
    doc_defaults = f"""
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}" w:eastAsia="{FONT}"/>
        <w:sz w:val="{SZ14}"/>
        <w:szCs w:val="{SZ14}"/>
        <w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:widowControl/>
        <w:spacing w:line="{LS}" w:lineRule="auto"/>
        <w:jc w:val="both"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>"""

    all_styles = "".join([
        s_normal, s_dpf,
        s_h1, s_h2, s_h3, s_h4, s_h5, s_h6,
        s_title, s_subtitle, s_author, s_date,
        s_abstract, s_abstract_title,
        s_first_par, s_body_text, s_compact,
        s_list_par, s_block,
        s_code, s_verbatim_char,
        s_captioned_figure, s_figure_par,
        s_caption, s_img_caption, s_tbl_caption,
        s_table, s_table_par,
        s_equation,
        s_fn_text, s_fn_text_char, s_fn_ref,
        s_bib,
        s_def_term, s_def,
        s_toc_heading, s_toc1, s_toc2, s_toc3,
        s_hyperlink,
        s_header, s_footer,
    ])

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {ns}
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  mc:Ignorable="">
{doc_defaults}
{all_styles}
</w:styles>
"""


# ── document.xml — страница с разметкой ГОСТ ─────────────────────────────────
def make_document():
    """Возвращает (word/document.xml, word/footer1.xml): пустое тело с параметрами страницы и колонтитул с номером."""
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # §4.4: страницы нумеруются арабскими цифрами в центре нижнего поля
    # Реализуется нижним колонтитулом из одного поля PAGE, подключённым в sectPr
    page_number_footer = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:ftr xmlns:w="{W}">\n'
        '  <w:p>\n'
        '    <w:pPr>\n'
        '      <w:pStyle w:val="Footer"/>\n'
        '      <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>\n'
        '      <w:ind w:left="0" w:right="0" w:firstLine="0"/>\n'
        '      <w:jc w:val="center"/>\n'
        '    </w:pPr>\n'
        '    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n'
        '    <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>\n'
        '    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n'
        '  </w:p>\n'
        '</w:ftr>'
    )

    # §4.2: поля страницы
    # §16: формат А4
    sect = f"""
    <w:sectPr>
      <w:footerReference w:type="default" r:id="rId7"/>
      <w:pgSz w:w="{PG_W}" w:h="{PG_H}"/>
      <w:pgMar w:top="{MT}" w:right="{MR}" w:bottom="{MB}" w:left="{ML}"
               w:header="{mm(10)}" w:footer="{mm(10)}" w:gutter="0"/>
      <w:cols w:num="1"/>
      <w:docGrid w:linePitch="{twp(14 * 1.5)}"/>
    </w:sectPr>"""

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  mc:Ignorable="wpc">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
      <w:r><w:t></w:t></w:r>
    </w:p>
    {sect}
  </w:body>
</w:document>
""", page_number_footer


DOC_RELS_WITH_FOOTER = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    Target="numbering.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
    Target="fontTable.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
    Target="footnotes.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
    Target="endnotes.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    Target="footer1.xml"/>
</Relationships>
"""

CONTENT_TYPES_WITH_FOOTER = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/numbering.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/fontTable.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/footnotes.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  <Override PartName="/word/footer1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/core.xml"
    ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml"
    ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


# ─────────────────────────────────────────────────────────────────────────────
# Сборка .docx
# ─────────────────────────────────────────────────────────────────────────────
def build():
    """Записывает GOST_7.32-2017.docx в папку профиля, перезаписывая существующий файл."""
    out_dir = Path(__file__).resolve().parent.parent
    out_path = out_dir / "GOST_7.32-2017.docx"

    document_xml, footer_xml = make_document()
    styles_xml = make_styles()

    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",          CONTENT_TYPES_WITH_FOOTER)
        z.writestr("_rels/.rels",                  RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS_WITH_FOOTER)
        z.writestr("word/document.xml",            document_xml)
        z.writestr("word/styles.xml",              styles_xml)
        z.writestr("word/settings.xml",            SETTINGS)
        z.writestr("word/numbering.xml",           NUMBERING)
        z.writestr("word/fontTable.xml",           FONT_TABLE)
        z.writestr("word/footnotes.xml",           FOOTNOTES)
        z.writestr("word/endnotes.xml",            ENDNOTES)
        z.writestr("word/footer1.xml",             footer_xml)
        z.writestr("docProps/core.xml",            CORE)
        z.writestr("docProps/app.xml",             APP)

    print(f"Создан: {out_path}")
    print(f"  Страница:  A4  {PG_W}×{PG_H} twips  (§16)")
    print(f"  Поля:  Л={ML}  П={MR}  В={MT}  Н={MB} twips  (§4.2)")
    print(f"  Шрифт: {FONT}  {SZ14//2}пт  (§4.1)")
    print(f"  Межстрочный: {LS}/240 = {LS/240:.1f}  (§4.1)")
    print(f"  Абзацный отступ: {IND} twips = {IND/1440*25.4:.2f} мм  (§4.1)")
    print(f"  Отступ до заголовка: {SP_B2} twips = {SP_B2/20:.0f}пт  (§7.2)")
    print(f"  Отступ после заголовка: {SP_A1} twips = {SP_A1/20:.0f}пт  (§7.2)")
    print(f"  Стилей: 42  (Normal, DefaultParagraphFont,")
    print(f"              Heading1-6, Title, Subtitle, Author, Date,")
    print(f"              Abstract, AbstractTitle, FirstParagraph, BodyText,")
    print(f"              Compact, ListParagraph, BlockText, SourceCode,")
    print(f"              VerbatimChar, CaptionedFigure, FigureParagraph,")
    print(f"              Caption, ImageCaption, TableCaption,")
    print(f"              Table, TableParagraph, Equation, FootnoteText,")
    print(f"              FootnoteTextChar, FootnoteReference, Bibliography,")
    print(f"              DefinitionTerm, Definition, TOCHeading, TOC1-3,")
    print(f"              Hyperlink, Header, Footer)")
    print(f"  Нумерация перечислений: §7.3 — дефис (numId=1), цифры со скобкой (numId=2)")
    print(f"  Нижний колонтитул: номер страницы по центру  (§4.4)")


if __name__ == "__main__":
    build()

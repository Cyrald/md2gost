-- Lua-фильтр pandoc профиля ГОСТ 7.32-2017: преобразования AST во время
-- конвертации в DOCX. Диаграммы рендерит препроцессор tools/preprocess.py
-- ещё до запуска pandoc, поэтому здесь их обработки нет.
--
-- Основной фильтр обходит дерево сверху вниз (topdown): узел Figure
-- обрабатывается раньше своих дочерних элементов. Это позволяет применить
-- замену SoftBreak→LineBreak к содержимому рисунка вручную, а его подпись
-- (Figure.caption.long) этой замене не подвергать; все прочие абзацы
-- обрабатывает общий обработчик Para.

local function _softbreak_to_linebreak()
    return pandoc.LineBreak()
end

local _SOFTBREAK_FILTER = {
    Para = function(p)
        p.content = p.content:walk({ SoftBreak = _softbreak_to_linebreak })
        return p
    end
}

-- Горизонтальная линия «---» означает «начать следующий абзац с новой страницы».
--
-- Символ <w:br w:type="page"/> не используется: он оказывается в отдельном
-- абзаце, который даёт пустую строку вверху новой страницы, а если текст
-- перед ним заканчивается ровно на границе страницы — целую пустую страницу.
-- Вместо этого следующий абзац должен получить свойство <w:pageBreakBefore/>
-- («С новой страницы»), которое Word и LibreOffice игнорируют, если абзац
-- и так начинается в начале страницы, — тот же механизм использует стиль
-- Heading1.
--
-- Из фильтра нельзя задать свойства обычному абзацу через DOCX-писатель
-- pandoc, поэтому линия превращается в абзац-носитель с собственным
-- <w:pageBreakBefore/>. В нём нет текста, высота строки — ровно 1 пт,
-- интервалы и отступы нулевые, а keepNext привязывает его к следующему
-- блоку, так что он фактически невидим.
local _PAGE_BREAK_MARKER = pandoc.RawBlock(
    "openxml",
    '<w:p><w:pPr><w:keepNext/><w:pageBreakBefore/>'
    .. '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
    .. '<w:ind w:left="0" w:right="0" w:firstLine="0"/>'
    .. '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>'
)

local function _page_break_marker()
    return _PAGE_BREAK_MARKER
end

-- Выполняется отдельным проходом после основного фильтра, чтобы маркер
-- не мешал обработке изображений и рисунков:
--   * несколько линий подряд дают один разрыв страницы;
--   * линия перед заголовком первого уровня отбрасывается (у Heading1
--     в эталонном документе уже есть pageBreakBefore);
--   * линия в самом конце, после которой ничего нет, отбрасывается.
local _PAGE_BREAK_FILTER = {
    Blocks = function(blocks)
        local result  = pandoc.Blocks({})
        local pending = false
        for _, block in ipairs(blocks) do
            if block.t == "HorizontalRule" then
                pending = true
            else
                if pending then
                    if not (block.t == "Header" and block.level == 1) then
                        result:insert(_page_break_marker())
                    end
                    pending = false
                end
                result:insert(block)
            end
        end
        return result
    end,
}

local _MAIN_FILTER = {
    traverse = "topdown",


    -- pandoc 3 оборачивает отдельное изображение с подписью (alt-текстом)
    -- в узел Figure. Изображение без подписи остаётся обычным
    -- Para [Image ...] и обрабатывается обработчиком Para ниже.
    --
    -- Замена SoftBreak→LineBreak применяется к содержимому рисунка, но
    -- намеренно не к подписи, чтобы многострочная подпись выводилась
    -- одной сплошной строкой, как требует ГОСТ.
    -- Возврат (fig, false) запрещает pandoc обходить дочерние элементы,
    -- поэтому общий обработчик Para внутри Figure не срабатывает.
    Figure = function(fig)
        local new_content = {}
        for _, block in ipairs(fig.content) do
            new_content[#new_content + 1] = pandoc.walk_block(block, _SOFTBREAK_FILTER)
        end
        fig.content = new_content
        return fig, false
    end,

    -- Назначаем каждой таблице стиль «Table», чтобы pandoc записал
    -- <w:tblStyle w:val="Table"/> в <w:tblPr>. Без этого pandoc создаёт
    -- таблицы без w:tblStyle, и Word показывает их без границ, даже если
    -- в эталонном документе есть стиль Table с границами.
    Table = function(t)
        t.attr.attributes["custom-style"] = "Table"
        return t
    end,

    -- Обработчик Para решает две задачи:
    --   1. Отдельное изображение без подписи (Para [Image ...]): pandoc
    --      не создаёт для него Figure, и изображение попадает в абзац
    --      BodyText с абзацным отступом и выравниванием по ширине. Абзац
    --      оборачивается в Div{custom-style="FigureParagraph"}, чтобы
    --      DOCX-писатель применил стиль по центру без отступа из эталонного
    --      документа. У Para нет поля .attr, поэтому custom-style задаётся
    --      на обёртке Div. Метка «data-gost-fig» в Image.attr не даёт
    --      обходу сверху вниз повторно обернуть внутренний Para.
    --   2. Все остальные абзацы: SoftBreak → LineBreak, чтобы одиночный
    --      перевод строки в Markdown становился разрывом строки в DOCX.
    --      Внутри Figure не срабатывает (см. обработчик Figure выше).
    Para = function(el)
        if #el.content == 1 and el.content[1].t == "Image" then
            local img = el.content[1]
            -- Абзац уже обёрнут: мы внутри Div, повторно не оборачиваем.
            if img.attr.attributes["data-gost-fig"] then
                return el
            end
            img.attr.attributes["data-gost-fig"] = "1"
            el.content[1] = img
            return pandoc.Div(
                {el},
                pandoc.Attr("", {}, {["custom-style"] = "FigureParagraph"})
            )
        end
        el.content = el.content:walk({ SoftBreak = _softbreak_to_linebreak })
        return el
    end,
}

return { _MAIN_FILTER, _PAGE_BREAK_FILTER }

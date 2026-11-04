"""
Виджеты интерфейса: всплывающая подсказка, выбор файла, выбор шаблона,
консоль журнала, строка состояния pandoc, диалог установки pandoc
и обработчик logging, пишущий в консоль.
"""
import queue
import shutil
import logging
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox
from typing import Callable, List, Optional
from core.logging_utils import mirror_log_message

_NO_SELECTION_LABEL = "— Не использовать —"


class ToolTip:
    """
    Всплывающая подсказка при наведении указателя на виджет.
    Появляется с задержкой 600 мс, чтобы при быстром движении мыши окна
    не мелькали; скрывается при уходе указателя или нажатии кнопки мыши.
    """

    def __init__(self, widget: tk.Widget, text_fn) -> None:
        """
        text_fn — функция без аргументов, возвращающая текст подсказки.
        Функция вместо строки нужна, чтобы подсказка отражала текущее
        состояние виджета (например, выбранный элемент списка).
        Пустой текст — подсказка не показывается.
        """
        self._widget   = widget
        self._text_fn  = text_fn
        self._tip: Optional[tk.Toplevel] = None
        self._after_id: Optional[str]   = None
        widget.bind("<Enter>",       self._schedule)
        widget.bind("<Leave>",       self._cancel)
        widget.bind("<ButtonPress>", self._cancel)

    def _schedule(self, event=None) -> None:
        self._cancel()
        self._after_id = self._widget.after(600, self._show)

    def _cancel(self, event=None) -> None:
        if self._after_id:
            self._widget.after_cancel(self._after_id)
            self._after_id = None
        if self._tip:
            self._tip.destroy()
            self._tip = None

    def _show(self) -> None:
        """Создаёт окно подсказки под виджетом; при пустом тексте ничего не показывает."""
        text = self._text_fn()
        if not text:
            return
        x = self._widget.winfo_rootx() + 10
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 4
        self._tip = tk.Toplevel(self._widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self._tip,
            text=text,
            justify=tk.LEFT,
            bg="#ffffe0",
            relief="solid",
            borderwidth=1,
            font=("TkDefaultFont", 9),
            wraplength=400,
        ).pack()


class FilePickerFrame(ttk.Frame):
    """
    Подпись, поле пути, кнопка очистки и кнопка «Обзор…» в одном виджете.
    Поле только для чтения: путь виден, но случайно его не изменить.
    Кнопка «✕» очищает путь и неактивна, когда путь пуст.
    """

    def __init__(
        self,
        master: tk.Misc,
        label_text: str,
        initial: str = "",
        ftypes=(("Все файлы", "*.*"),),
        label_width: int = 22,
    ) -> None:
        super().__init__(master)
        self.ftypes = ftypes
        self.var    = tk.StringVar(value=initial)

        ttk.Label(self, text=label_text, width=label_width, anchor="w").pack(side=tk.LEFT)
        self.entry = ttk.Entry(self, textvariable=self.var, state="readonly")
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 3))

        self._clear_btn = ttk.Button(self, text="✕", command=self._clear, width=2)
        self._clear_btn.pack(side=tk.LEFT, padx=(0, 3))

        ttk.Button(self, text="Обзор…", command=self._browse, width=9).pack(side=tk.LEFT)

        self._update_clear_btn()

    def get_path(self) -> str:
        return self.var.get()

    def set_path(self, p: str) -> None:
        self.var.set(p)
        self._update_clear_btn()

    def _clear(self) -> None:
        self.set_path("")

    def _update_clear_btn(self) -> None:
        self._clear_btn.config(state="normal" if self.var.get() else "disabled")

    def _browse(self) -> None:
        p = filedialog.askopenfilename(filetypes=self.ftypes, parent=self)
        if p:
            self.set_path(p)


class TemplateSelectorFrame(ttk.Frame):
    """
    Подпись и выпадающий список для выбора шаблона по отображаемому имени.
    Первый элемент означает «не использовать»; при его выборе get_value()
    возвращает None, чтобы вызывающему коду не приходилось сравнивать строки.

    Необязательная кнопка импорта («📂+») копирует выбранный пользователем
    файл в target_dir, после чего список обновляется и новый элемент
    выбирается автоматически. Если файл с таким именем уже есть,
    пользователь подтверждает перезапись.

    Параметры
    ---------
    target_dir : Path или None
        Папка, куда копируются импортируемые файлы. Если None, кнопка
        импорта не показывается.
    ftypes : tuple
        Фильтры типов файлов для диалога открытия, например
        ``(("Word DOCX", "*.docx"),)`` или ``(("CSL", "*.csl"),)``.
    scan_ext : str или None
        Расширение (с точкой), по которому target_dir пересканируется после
        импорта, например ``".docx"`` или ``".csl"``. Должно совпадать с ftypes:
        только файлы с этим расширением попадают в список и принимаются к импорту.
    hide_ext : bool
        Если True, элементы показываются без расширения (тогда и opts нужно
        передавать без расширений, и искать выбранный файл по имени без него).
    """

    def __init__(
        self,
        master: tk.Misc,
        label_text: str,
        opts: List[str],
        label_width: int = 22,
        target_dir: Optional[Path] = None,
        ftypes: Optional[tuple] = None,
        scan_ext: Optional[str] = None,
        hide_ext: bool = False,
    ) -> None:
        super().__init__(master)
        self._target_dir = Path(target_dir) if target_dir else None
        self._scan_ext   = scan_ext.lower() if scan_ext else None
        self._hide_ext   = hide_ext

        ttk.Label(self, text=label_text, width=label_width, anchor="w").pack(side=tk.LEFT)

        self.combo = ttk.Combobox(self, state="readonly")
        self.combo["values"] = [_NO_SELECTION_LABEL] + opts
        self.combo.current(0)
        self.combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 3))

        ToolTip(
            self.combo,
            lambda: (self.combo.get()
                     if self.combo.get() != _NO_SELECTION_LABEL else ""),
        )

        if self._target_dir is not None and ftypes:
            self._ftypes = ftypes
            import_btn = ttk.Button(
                self, text="📂+", command=self._import_file, width=4
            )
            import_btn.pack(side=tk.LEFT)
            ToolTip(import_btn, lambda: "Добавить файл в папку шаблонов")

    def get_value(self) -> Optional[str]:
        """Возвращает имя выбранного шаблона или None, если выбрано «не использовать»."""
        val = self.combo.get()
        return None if val == _NO_SELECTION_LABEL else val

    def _import_file(self) -> None:
        """Копирует выбранный пользователем файл в папку шаблонов и выбирает его в списке."""
        p = filedialog.askopenfilename(filetypes=self._ftypes, parent=self)
        if not p:
            return

        src = Path(p)
        if self._scan_ext and src.suffix.lower() != self._scan_ext:
            messagebox.showerror(
                "Ошибка импорта",
                f"Можно добавить только файлы {self._scan_ext}.\n\nВыбран: {src.name}",
                parent=self,
            )
            return
        dst = self._target_dir / src.name

        if dst.exists():
            if not messagebox.askyesno(
                "Файл уже существует",
                f"Файл «{src.name}» уже есть в папке шаблонов.\nПерезаписать?",
                parent=self,
            ):
                return

        try:
            self._target_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(str(src), str(dst))
        except OSError as exc:
            messagebox.showerror(
                "Ошибка импорта",
                f"Не удалось скопировать файл в папку шаблонов:\n\n"
                f"{src}\n\nПричина: {exc}",
                parent=self,
            )
            return
        self._refresh(self._display_name(src))

    def _display_name(self, path: Path) -> str:
        return path.stem if self._hide_ext else path.name

    def _refresh(self, select_name: Optional[str] = None) -> None:
        """Пересканирует target_dir, обновляет список и выбирает select_name, если он есть в списке."""
        if self._target_dir is None or self._scan_ext is None:
            return
        files = sorted(
            self._display_name(f)
            for f in self._target_dir.iterdir()
            if f.is_file() and f.suffix.lower() == self._scan_ext
        )
        self.combo["values"] = [_NO_SELECTION_LABEL] + files
        if select_name and select_name in files:
            self.combo.set(select_name)
        else:
            # Возвращаемся к «не использовать», даже если список не пуст:
            # иначе в поле мог бы остаться уже несуществующий файл.
            self.combo.current(0)


class LogConsole(tk.Text):
    """
    Текстовый виджет с тёмной темой для журнала конвертации.
    Каждый уровень журнала выделяется своим цветом, чтобы ошибки
    и предупреждения были заметны.

    Для пользователя виджет только для чтения: текст добавляется через
    append(). Исключение — режим ввода (enable_input), в котором консоль
    работает как простой терминал для ручной установки pandoc.

    Объём хранимого текста ограничен примерно 2000 строками.
    """

    _LEVEL_TAGS = {
        "DEBUG":    "#858585",
        "INFO":     "#9cdcfe",
        "WARNING":  "#dcdcaa",
        "ERROR":    "#f44747",
        "CRITICAL": "#f44747",
    }

    def __init__(self, master: tk.Misc, height: int = 12) -> None:
        import tkinter.font as _tkfont
        _mono = "Consolas" if "Consolas" in _tkfont.families() else "TkFixedFont"
        super().__init__(
            master,
            height=height,
            state="disabled",
            bg="#1e1e1e",
            fg="#d4d4d4",
            font=(_mono, 10),
            relief="flat",
            wrap=tk.WORD,
        )
        for level, colour in self._LEVEL_TAGS.items():
            self.tag_config(level, foreground=colour)
        self.bind("<Control-c>", self._on_copy)
        self._ctx_menu = tk.Menu(self, tearoff=0)
        self._ctx_menu.add_command(label="Копировать", command=self._on_copy)
        self.bind("<Button-3>", self._show_context_menu)
        # Счётчик логических строк (переводов строки в добавленных сообщениях),
        # а не визуальных строк Tk, зависящих от переноса и ширины виджета.
        # Запрос индекса у виджета при каждой вставке стоил бы лишнего обращения
        # к Tcl, а для ограничения объёма приблизительного счёта достаточно.
        self._line_count: int = 0
        # Tkinter не потокобезопасен: фоновые потоки только кладут сообщения
        # в эту очередь, а главный поток выгружает её в _drain_loop().
        self._pending: "queue.SimpleQueue[tuple[str, str]]" = queue.SimpleQueue()
        self._drain_after_id: Optional[str] = self.after(
            self._DRAIN_INTERVAL_MS, self._drain_loop,
        )
        # Режим ввода (enable_input): текст после метки «input_start»
        # редактируется, всё до неё — защищённый от изменений вывод.
        self._input_cb: Optional[Callable[[str], None]] = None
        self.bind("<Key>", self._on_input_key)
        for sequence in ("<<Paste>>", "<<Cut>>", "<<PasteSelection>>", "<<Clear>>"):
            self.bind(sequence, lambda e, s=sequence: self._on_input_virtual(s))

    _DRAIN_INTERVAL_MS = 50
    # Клавиши, которые не изменяют текст и поэтому разрешены в любом месте консоли.
    _NAV_KEYS = frozenset({
        "Left", "Right", "Up", "Down", "Home", "End", "Prior", "Next",
        "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R",
        "Caps_Lock", "Escape", "Tab",
    })

    @property
    def input_enabled(self) -> bool:
        return self._input_cb is not None

    def enable_input(self, on_line: Callable[[str], None]) -> None:
        """
        Превращает консоль в простой терминал: пользователь печатает после
        последнего вывода, а Enter передаёт строку в on_line (в главном потоке).
        """
        self.flush_pending()
        self._input_cb = on_line
        self.config(state="normal", insertbackground="#d4d4d4")
        self.mark_set("input_start", "end-1c")
        self.mark_gravity("input_start", tk.LEFT)
        self.mark_set("insert", "end-1c")
        self.see("end")
        self.focus_set()

    def disable_input(self) -> None:
        """Выключает режим ввода и возвращает консоль в режим только для чтения."""
        if self._input_cb is None:
            return
        self.flush_pending()
        self._input_cb = None
        try:
            self.mark_unset("input_start")
            self.config(state="disabled")
        except tk.TclError:
            pass

    def _selection_touches_output(self) -> bool:
        return bool(self.tag_ranges("sel")) and self.compare("sel.first", "<", "input_start")

    def _move_insert_to_input(self) -> None:
        if self._selection_touches_output():
            self.tag_remove("sel", "1.0", "end")
        if self.compare("insert", "<", "input_start"):
            self.mark_set("insert", "end-1c")

    def _on_input_key(self, event) -> Optional[str]:
        """
        Фильтрует нажатия в режиме ввода: «break» отменяет нажатие, None
        пропускает его стандартному обработчику Text. Изменять можно только
        текст после «input_start».
        """
        if self._input_cb is None:
            return None
        if event.keysym in ("Return", "KP_Enter"):
            self._submit_input()
            return "break"
        if event.keysym in self._NAV_KEYS:
            return None
        # С Ctrl разрешена только вставка; копирование привязано отдельно.
        if event.state & 0x4:
            return None if event.keysym.lower() == "v" else "break"
        if event.keysym in ("BackSpace", "Delete"):
            if self._selection_touches_output():
                return "break"
            if self.tag_ranges("sel"):
                return None
            # BackSpace удаляет символ слева от курсора, поэтому запрещён уже
            # на самой границе ввода; Delete — только до неё.
            limit = "<=" if event.keysym == "BackSpace" else "<"
            if self.compare("insert", limit, "input_start"):
                return "break"
            return None
        self._move_insert_to_input()
        return None

    def _on_input_virtual(self, sequence: str) -> Optional[str]:
        """Не даёт вставке, вырезанию и очистке изменить защищённый вывод."""
        if self._input_cb is None:
            return None
        if sequence == "<<Paste>>":
            self._move_insert_to_input()
            return None
        if sequence in ("<<Cut>>", "<<Clear>>") and not self._selection_touches_output():
            return None
        return "break"

    def _submit_input(self) -> None:
        """Фиксирует введённый текст как вывод и передаёт его в обработчик построчно."""
        callback = self._input_cb
        if callback is None:
            return
        text = self.get("input_start", "end-1c")
        self.insert("end", "\n")
        self.mark_set("input_start", "end-1c")
        self.mark_set("insert", "end-1c")
        self._line_count += 1
        self.see("end")
        for line in text.split("\n"):
            callback(line)

    def append(self, msg: str, level: str = "INFO") -> None:
        """
        Ставит сообщение в очередь на вывод с цветом его уровня и сразу
        дублирует его в консоль процесса. Tk здесь не вызывается, поэтому
        метод безопасно вызывать из любого потока.
        """
        tag = level if level in self._LEVEL_TAGS else "INFO"
        self._pending.put((msg, tag))
        mirror_log_message(msg, tag)

    def flush_pending(self) -> None:
        """Немедленно выводит все сообщения из очереди. Вызывать только в главном потоке."""
        batch: list[tuple[str, str]] = []
        while True:
            try:
                batch.append(self._pending.get_nowait())
            except queue.Empty:
                break
        if not batch:
            return
        try:
            if not self.winfo_exists():
                return
            self.config(state="normal")
            for msg, tag in batch:
                if self._input_cb is not None:
                    # Вывод вставляется перед набираемой строкой. У «input_start»
                    # левая гравитация, и сама она за вставкой не сдвинется,
                    # поэтому её переносим за новый текст вручную.
                    self.mark_set("output_end", "input_start")
                    self.mark_gravity("output_end", tk.RIGHT)
                    self.insert("output_end", msg, tag)
                    self.mark_set("input_start", "output_end")
                    self.mark_unset("output_end")
                else:
                    self.insert("end", msg, tag)
                self._line_count += msg.count("\n")
            if self._line_count > 2000:
                excess = self._line_count - 2000
                self.delete("1.0", f"{excess + 1}.0")
                self._line_count = 2000
            if self._input_cb is None:
                self.config(state="disabled")
            self.see("end")
        except tk.TclError:
            pass

    def _drain_loop(self) -> None:
        """Каждые _DRAIN_INTERVAL_MS мс выводит накопившиеся сообщения."""
        self.flush_pending()
        try:
            self._drain_after_id = self.after(self._DRAIN_INTERVAL_MS, self._drain_loop)
        except tk.TclError:
            self._drain_after_id = None

    def destroy(self) -> None:
        if self._drain_after_id is not None:
            try:
                self.after_cancel(self._drain_after_id)
            except tk.TclError:
                pass
            self._drain_after_id = None
        super().destroy()

    def _on_copy(self, event=None) -> str:
        """Обработчик Ctrl+C и пункта меню: копирует выделение, а если его нет — весь текст."""
        try:
            text = self.get("sel.first", "sel.last")
        except tk.TclError:
            text = self.get("1.0", "end-1c")
        if text:
            self.clipboard_clear()
            self.clipboard_append(text)
        return "break"

    def _show_context_menu(self, event) -> None:
        """Показывает контекстное меню по правому щелчку в позиции указателя."""
        try:
            self._ctx_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self._ctx_menu.grab_release()

    def clear(self) -> None:
        """Очищает консоль вместе с ещё не выведенной очередью и прокручивает её в начало."""
        while True:
            try:
                self._pending.get_nowait()
            except queue.Empty:
                break
        self.config(state="normal")
        self.delete("1.0", "end")
        self._line_count = 0
        if self._input_cb is not None:
            self.mark_set("input_start", "end-1c")
            self.mark_set("insert", "end-1c")
        else:
            self.config(state="disabled")
        self.see("1.0")


class PandocStatusBar(ttk.Frame):
    """
    Строка состояния pandoc внизу главного окна. Скрыта, пока всё в порядке
    (проверка при запуске, pandoc готов), и появляется только во время
    установки или после неудачной/отменённой установки.
    Размещается через dock(), а не pack(). Все методы вызываются только
    в главном потоке.
    """

    _DOT_COLOURS = {
        "checking":   "#9aa0a6",
        "installing": "#1a73e8",
        "ready":      "#1e8e3e",
        "error":      "#d93025",
        "cancelled":  "#e37400",
    }
    _HIDDEN_KINDS = frozenset({"checking", "ready"})
    _SECONDARY_FG = "#5f6368"
    _ANIM_MS = 30

    def __init__(self, master: tk.Misc, on_retry: Callable[[], None]) -> None:
        import tkinter.font as _tkfont
        super().__init__(master)
        style = ttk.Style(self)
        bg = style.lookup("TFrame", "background") or self.winfo_toplevel().cget("background")

        self._title_font = _tkfont.nametofont("TkDefaultFont").copy()
        self._title_font.configure(weight="bold")

        ttk.Separator(self, orient=tk.HORIZONTAL).pack(side=tk.TOP, fill=tk.X)
        body = ttk.Frame(self, padding=(10, 5))
        body.pack(side=tk.TOP, fill=tk.X)

        self._dot = tk.Canvas(body, width=10, height=10, highlightthickness=0, bd=0, bg=bg)
        self._dot_id = self._dot.create_oval(1, 1, 9, 9, outline="", fill=self._DOT_COLOURS["checking"])
        self._dot.pack(side=tk.LEFT, padx=(0, 8))

        self._title = ttk.Label(body, font=self._title_font)
        self._title.pack(side=tk.LEFT)
        self._detail = ttk.Label(body, foreground=self._SECONDARY_FG)
        self._detail.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)

        self._right = ttk.Frame(body)
        self._right.pack(side=tk.RIGHT)
        self._bar = ttk.Progressbar(self._right, length=170, mode="determinate", maximum=100)
        self._percent = ttk.Label(self._right, width=5, anchor=tk.E, foreground=self._SECONDARY_FG)
        self._retry_btn = ttk.Button(self._right, text="Повторить", command=on_retry)

        self._target = 0.0
        self._shown = 0.0
        self._anim_id: Optional[str] = None
        self._dock_opts: Optional[dict] = None
        self._visible = False
        self.show_checking()

    def dock(self, **pack_opts) -> None:
        """Запоминает параметры pack(); строка размещается по ним, только когда ей есть что показать."""
        self._dock_opts = pack_opts
        self._apply_visibility()

    def show_checking(self) -> None:
        self._render("checking", "Проверка pandoc…", "")

    def show_installing(self, detail: str, percent: float) -> None:
        self._render("installing", "Установка pandoc", detail, percent=percent)

    def show_ready(self, detail: str) -> None:
        self._render("ready", "pandoc готов к работе", detail)

    def show_error(self, detail: str = "Подробности — в логе") -> None:
        self._render("error", "Не удалось установить pandoc", detail, retry=True)

    def show_cancelled(self) -> None:
        self._render(
            "cancelled", "Установка pandoc отменена",
            "Нажмите «Повторить», чтобы установить pandoc", retry=True,
        )

    def _render(
        self, kind: str, title: str, detail: str,
        percent: Optional[float] = None, retry: bool = False,
    ) -> None:
        """
        Перерисовывает строку под состояние kind: цвет индикатора, тексты,
        полоса прогресса (если передан percent) и кнопка «Повторить» (если retry).
        """
        try:
            self._dot.itemconfigure(self._dot_id, fill=self._DOT_COLOURS[kind])
            self._title.config(text=title)
            self._detail.config(text=detail)
            for child in (self._bar, self._percent, self._retry_btn):
                child.pack_forget()
            if percent is not None:
                self._bar.pack(side=tk.LEFT)
                self._percent.pack(side=tk.LEFT, padx=(6, 0))
                self._set_percent(percent)
            else:
                self._set_percent(0.0, immediate=True)
            if retry:
                self._retry_btn.pack(side=tk.LEFT, padx=(10, 0))
            self._visible = kind not in self._HIDDEN_KINDS
            self._apply_visibility()
        except tk.TclError:
            pass

    def _apply_visibility(self) -> None:
        if self._dock_opts is None:
            return
        is_packed = bool(self.winfo_manager())
        if self._visible and not is_packed:
            # pack распределяет место в порядке размещения, поэтому строку
            # ставим перед остальными виджетами — так она не сожмётся,
            # когда окно маленькое.
            opts = dict(self._dock_opts)
            siblings = [w for w in self.master.pack_slaves() if w is not self]
            if siblings and "before" not in opts and "after" not in opts:
                opts["before"] = siblings[0]
            self.pack(**opts)
        elif not self._visible and is_packed:
            self.pack_forget()

    def _set_percent(self, value: float, immediate: bool = False) -> None:
        """
        Задаёт целевой процент. Рост анимируется плавно, а уменьшение
        (или immediate=True) применяется сразу.
        """
        self._target = max(0.0, min(100.0, value))
        if immediate or self._target < self._shown:
            self._shown = self._target
        self._paint()
        if self._anim_id is None and self._shown != self._target:
            self._anim_id = self.after(self._ANIM_MS, self._animate)

    def _animate(self) -> None:
        # Установщик сообщает прогресс загрузки крупными шагами; каждый кадр
        # сокращает разрыв до цели на 20 %, чтобы полоса двигалась плавно.
        self._anim_id = None
        diff = self._target - self._shown
        self._shown = self._target if abs(diff) < 0.5 else self._shown + diff * 0.2
        self._paint()
        if self._shown != self._target:
            try:
                self._anim_id = self.after(self._ANIM_MS, self._animate)
            except tk.TclError:
                self._anim_id = None

    def _paint(self) -> None:
        try:
            self._bar["value"] = self._shown
            self._percent.config(text=f"{round(self._shown)} %")
        except tk.TclError:
            pass

    def destroy(self) -> None:
        if self._anim_id is not None:
            try:
                self.after_cancel(self._anim_id)
            except tk.TclError:
                pass
            self._anim_id = None
        super().destroy()


def ask_pandoc_action(parent: tk.Misc, details: str) -> Optional[str]:
    """
    Модальный диалог, показываемый, когда `install_pandoc.py --check` сообщает
    о проблеме. details — вывод самого установщика.
    Возвращает "auto", "manual" или None (отложить).

    Сейчас не вызывается: GUI запускает автоматическую установку без вопросов
    (см. gui/pandoc_setup.py).
    """
    import tkinter.font as _tkfont
    mono = "Consolas" if "Consolas" in _tkfont.families() else "TkFixedFont"

    result: dict = {"value": None}
    dlg = tk.Toplevel(parent)
    dlg.title("pandoc не готов к работе")
    dlg.transient(parent)
    dlg.minsize(480, 300)

    frame = ttk.Frame(dlg, padding=12)
    frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(
        frame, text="Проверка установщика pandoc сообщила о проблеме:", anchor="w",
    ).pack(fill=tk.X, pady=(0, 6))

    text_row = ttk.Frame(frame)
    text_row.pack(fill=tk.BOTH, expand=True)
    details_box = tk.Text(text_row, height=12, width=86, wrap=tk.WORD, font=(mono, 9), relief="solid", borderwidth=1)
    details_box.insert("1.0", details)
    details_box.config(state="disabled")
    scroll = ttk.Scrollbar(text_row, orient=tk.VERTICAL, command=details_box.yview)
    details_box.config(yscrollcommand=scroll.set)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)
    details_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    ttk.Label(
        frame,
        text=(
            "Без pandoc и pandoc-crossref конвертация невозможна.\n"
            "Автоматически — установщик сам подберёт и установит совместимые версии.\n"
            "Вручную — меню установщика откроется в консоли программы, ответы вводятся там же."
        ),
        justify=tk.LEFT,
        anchor="w",
    ).pack(fill=tk.X, pady=(10, 10))

    def choose(value: Optional[str]) -> None:
        result["value"] = value
        dlg.destroy()

    buttons = ttk.Frame(frame)
    buttons.pack(fill=tk.X)
    ttk.Button(buttons, text="Позже", command=lambda: choose(None)).pack(side=tk.RIGHT)
    ttk.Button(
        buttons, text="Установить вручную", command=lambda: choose("manual"),
    ).pack(side=tk.RIGHT, padx=(0, 6))
    auto_btn = ttk.Button(
        buttons, text="Установить автоматически", command=lambda: choose("auto"),
    )
    auto_btn.pack(side=tk.RIGHT, padx=(0, 6))

    dlg.protocol("WM_DELETE_WINDOW", lambda: choose(None))
    dlg.bind("<Escape>", lambda e: choose(None))
    dlg.bind("<Return>", lambda e: choose("auto"))
    try:
        dlg.wait_visibility()
        dlg.grab_set()
    except tk.TclError:
        pass
    auto_btn.focus_set()
    parent.wait_window(dlg)
    return result["value"]


class LogHandler(logging.Handler):
    """
    Обработчик logging, выводящий отформатированные записи в LogConsole.
    Уровень записи передаётся в консоль, чтобы она раскрасила сообщение.
    """

    def __init__(self, console: LogConsole) -> None:
        super().__init__()
        self.console = console

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record) + "\n"
            self.console.append(msg, record.levelname)
        except Exception:
            self.handleError(record)

    def flush(self) -> None:
        # Сбрасывать нечего: консоль сама выводит свою очередь в _drain_loop().
        pass

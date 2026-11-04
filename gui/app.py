"""
Главное окно приложения: компоновка виджетов и сбор параметров задачи.
"""
import tkinter as tk
from tkinter import ttk
from pathlib import Path
from typing import List, Optional, TYPE_CHECKING
from core.types import ConversionParams
from .components import FilePickerFrame, TemplateSelectorFrame, LogConsole, PandocStatusBar
from .events import EventHandler
from .pandoc_setup import PandocSetup

if TYPE_CHECKING:
    from core.config import ProjectConfig


class MainWindow(tk.Tk):
    """
    Корневое окно: владеет деревом виджетов, EventHandler и PandocSetup.
    Выбранные пользователем значения (пути, шаблоны, флажки) не дублируются
    в полях окна, а читаются из виджетов в get_task_parameters().
    """

    def __init__(
        self,
        titles: Optional[List[str]] = None,
        styles: Optional[List[str]] = None,
        cfg: Optional["ProjectConfig"] = None,
    ) -> None:
        super().__init__()
        self.title("md2gost")
        self.geometry("700x600")
        self.minsize(600, 500)
        self.resizable(True, True)

        self.titles = titles or []
        self.styles = styles or []
        self._cfg = cfg
        self.handler: Optional[EventHandler] = None
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        """Создаёт виджеты окна, затем PandocSetup и EventHandler, которым нужны готовые виджеты."""
        cfg = self._cfg

        _docx_ftypes = (("Word DOCX", "*.docx"),)

        ttl_dir = Path(cfg.resources_dir) / "titles" if cfg else None

        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Строка состояния pandoc скрыта, пока всё в порядке; она появляется
        # только во время установки или после неудачной/отменённой установки.
        self.pandoc_status = PandocStatusBar(self, on_retry=self._on_pandoc_retry)
        self.pandoc_status.dock(side=tk.BOTTOM, fill=tk.X)

        input_frame = ttk.LabelFrame(main_frame, text="Входные данные", padding=10)
        input_frame.pack(fill=tk.X, pady=(0, 10))

        md_bib_row = ttk.Frame(input_frame)
        md_bib_row.pack(fill=tk.X, pady=5)

        self.md_picker = FilePickerFrame(
            md_bib_row, "Markdown (.md):",
            ftypes=(("Markdown", "*.md"),),
            label_width=15,
        )
        self.md_picker.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        self.bib_picker = FilePickerFrame(
            md_bib_row, ".bib файл:",
            ftypes=(("BibTeX", "*.bib"),),
            label_width=10,
        )
        self.bib_picker.pack(side=tk.LEFT, fill=tk.X, expand=True)

        style_frame = ttk.LabelFrame(main_frame, text="Оформление (Шаблоны)", padding=10)
        style_frame.pack(fill=tk.X, pady=(0, 10))

        self.title_sel = TemplateSelectorFrame(
            style_frame, "Титульный лист:", self.titles,
            target_dir=ttl_dir, ftypes=_docx_ftypes, scan_ext=".docx",
        )
        self.title_sel.pack(fill=tk.X, pady=5)

        self.style_sel = TemplateSelectorFrame(
            style_frame, "Стили тела:", self.styles,
        )
        self.style_sel.pack(fill=tk.X, pady=5)

        checks_frame = ttk.Frame(style_frame)
        checks_frame.pack(anchor=tk.W, pady=5)

        self.toc_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            checks_frame, text="Авто-оглавление (TOC)", variable=self.toc_var
        ).pack(side=tk.LEFT, padx=(0, 20))

        self.num_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            checks_frame, text="Нумерация разделов", variable=self.num_var
        ).pack(side=tk.LEFT)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 6))

        self.convert_btn = ttk.Button(
            btn_frame,
            text="🚀 КОНВЕРТИРОВАТЬ",
            command=self._on_convert,
            width=20,
        )
        self.convert_btn.pack(side=tk.RIGHT)

        self.cancel_btn = ttk.Button(
            btn_frame,
            text="✕ Отмена",
            command=self._on_cancel,
            width=12,
            state="disabled",
        )
        self.cancel_btn.pack(side=tk.RIGHT, padx=(0, 5))

        self.clear_btn = ttk.Button(
            btn_frame,
            text="🗑 Очистить лог",
            command=self._on_clear_log,
            width=15,
        )
        self.clear_btn.pack(side=tk.LEFT)

        log_frame = ttk.LabelFrame(main_frame, text="Лог конвертации", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.console = LogConsole(log_frame)

        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.console.yview)
        self.console.config(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.console.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Обработчики создаются последними: им нужны уже готовые консоль
        # и кнопки, к которым они обращаются.
        self.pandoc_setup = PandocSetup(
            self, self.console, self._on_pandoc_setup_busy, status=self.pandoc_status,
        )
        self.handler = EventHandler(self, self.console, self.pandoc_setup, cfg=self._cfg)

    def _on_pandoc_retry(self) -> None:
        self.pandoc_setup.start_check()

    def _on_pandoc_setup_busy(self, busy: bool) -> None:
        """Пока работает установщик pandoc, конвертация заблокирована, а «Отмена» останавливает установщик."""
        try:
            self.convert_btn.config(state="disabled" if busy else "normal")
            self.cancel_btn.config(state="normal" if busy else "disabled")
        except tk.TclError:
            pass

    def _on_convert(self) -> None:
        if self.handler is None:
            raise RuntimeError("EventHandler не инициализирован — _build_ui() не вызван")
        self.handler.on_convert_click()

    def _on_cancel(self) -> None:
        if self.handler is None:
            raise RuntimeError("EventHandler не инициализирован — _build_ui() не вызван")
        if self.pandoc_setup.busy:
            self.pandoc_setup.cancel()
        else:
            self.handler.request_cancel()

    def _on_clear_log(self) -> None:
        self.console.clear()

    def _on_close(self) -> None:
        self.pandoc_setup.shutdown()
        if self.handler is not None:
            # Должно завершиться до destroy(): после него цикл событий
            # останавливается и запланированная очистка уже не выполнится.
            self.handler.shutdown()
        self.destroy()

    def get_task_parameters(self) -> ConversionParams:
        """
        Собирает текущее состояние виджетов в типизированный словарь.
        Пустой bib_path заменяется на None, чтобы дальше по конвейеру
        не приходилось отдельно обрабатывать пустую строку.
        """
        return ConversionParams(
            md_path=self.md_picker.get_path(),
            bib_path=self.bib_picker.get_path() or None,
            title_name=self.title_sel.get_value(),
            style_name=self.style_sel.get_value(),
            use_toc=self.toc_var.get(),
            number_headings=self.num_var.get(),
        )

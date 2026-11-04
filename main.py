"""
Точка входа md2gost: настраивает журналирование, сканирует ресурсы
(титульные листы и профили ГОСТ) и запускает главное окно.
"""
import logging
from core.config import ProjectConfig
from core.logging_utils import StartupLogBuffer
from gui.components import LogHandler
from gui.app import MainWindow


def main() -> None:
    """
    Запускает приложение.

    Пока окно не создано, предупреждения копятся в буфере; после создания окна
    они переносятся в консоль журнала GUI, а дальнейшие сообщения уровня
    WARNING и выше направляются туда напрямую. Консоль GUI дублирует каждое
    сообщение в stdout/stderr, поэтому вывод виден и в окне CMD.

    Наличие pandoc/pandoc-crossref здесь не проверяется: окно делегирует это
    tools/install_pandoc.py (см. gui/pandoc_setup.py).
    """
    startup_logs = StartupLogBuffer()
    logging.basicConfig(
        level=logging.WARNING,
        handlers=[startup_logs],
        force=True,
    )
    log_format = logging.Formatter("%(levelname)s: %(message)s")

    cfg = ProjectConfig()

    titles = [t.name  for t in cfg.scan_templates("titles")]
    styles = [profile.name for profile in cfg.scan_gost_profiles()]

    app = MainWindow(titles=titles, styles=styles, cfg=cfg)
    gui_log_handler = LogHandler(app.console)
    gui_log_handler.setLevel(logging.WARNING)
    gui_log_handler.setFormatter(log_format)

    # Подменяем буферный обработчик на консоль GUI и выводим в неё всё,
    # что успело накопиться во время сканирования ресурсов.
    root_logger = logging.getLogger()
    root_logger.removeHandler(startup_logs)
    root_logger.addHandler(gui_log_handler)
    for record in startup_logs.records:
        gui_log_handler.handle(record)
    startup_logs.records.clear()

    # Проверка pandoc запускается после отрисовки окна, чтобы при ошибке
    # пользователю сразу можно было предложить автоматическую или ручную установку.
    app.after(100, app.pandoc_setup.start_check)
    app.mainloop()


if __name__ == "__main__":
    main()

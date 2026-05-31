import sys
from pathlib import Path

import typer
from loguru import logger

from src.utils.logger import setup_logger

app = typer.Typer(add_completion=False, no_args_is_help=True)


@app.command()
def cli(
    input_file: Path = typer.Argument(..., exists=True, readable=True, help="Путь к входному .xlsm файлу"),
    output: Path | None = typer.Option(None, "--output", "-o", help="Путь для сохранения результата"),
    log_level: str = typer.Option("INFO", "--log-level", "-l"),
) -> None:
    """
    Запуск приложения в консольном режиме.
    """

    setup_logger(log_level=log_level, log_to_file=False)
    from src.core.processor import process_file
    result = process_file(input_file, output)
    typer.echo(result.summary())
    if not result.success:
        raise typer.Exit(code=1)


def run_gui() -> None:
    """
    Запуск графического интерфейса приложения.
    """

    from PyQt6.QtGui import QIcon
    from PyQt6.QtWidgets import QApplication
    from src.gui.main_window import MainWindow

    setup_logger(log_level="INFO", log_to_file=True)
    qt_app = QApplication(sys.argv)
    qt_app.setApplicationName("ЯКНО Spec Generator")
    icon_path = Path(__file__).parent / "assets" / "icon.ico"
    if icon_path.exists():
        qt_app.setWindowIcon(QIcon(str(icon_path)))
    window = MainWindow()
    window.show()
    sys.exit(qt_app.exec())


if __name__ == "__main__":
    if "--gui" in sys.argv:
        sys.argv.remove("--gui")
        run_gui()
    else:
        app()

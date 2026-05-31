import sys
from pathlib import Path
from loguru import logger

LOG_DIR = Path.home() / "AppData" / "Local" / "YaknoSpecGenerator" / "logs"
LOG_FILE = LOG_DIR / "app.log"

_configured = False


def setup_logger(log_level: str = "DEBUG", log_to_file: bool = True, force: bool = False) -> None:
    """
    Настройка системы логирования приложения.
    """

    global _configured
    if _configured and not force:
        return

    logger.remove()

    logger.add(
        sys.stderr,
        level=log_level,
        format="<green>{time:HH:mm:ss}</green> | <level>{level:<8}</level> | {message}",
        colorize=True,
    )

    if log_to_file:
        try:
            LOG_DIR.mkdir(parents=True, exist_ok=True)
            logger.add(
                str(LOG_FILE),
                level="DEBUG",
                format="{time:YYYY-MM-DD HH:mm:ss} | {level:<8} | {name}:{line} | {message}",
                rotation="5 MB",
                retention=5,
                encoding="utf-8",
            )
            logger.debug(f"Лог-файл: {LOG_FILE}")
        except Exception as e:
            logger.warning(f"Не удалось открыть лог-файл: {e}")

    _configured = True

import sys
import os
from pathlib import Path
from loguru import logger

def get_app_data_dir() -> Path:
    """Возвращает правильную папку для хранения логов"""
    if getattr(sys, 'frozen', False):
        # .exe режим
        return Path(sys.executable).parent / "logs"
    else:
        # Python режим
        if sys.platform == "win32":
            # Windows
            app_data = os.environ.get('APPDATA', str(Path.home() / 'AppData' / 'Roaming'))
            return Path(app_data) / "EMSpecGenerator" / "logs"
        else:
            # Linux/Mac
            return Path.home() / ".local" / "share" / "EMSpecGenerator" / "logs"

LOG_DIR = get_app_data_dir()
LOG_FILE = LOG_DIR / "app.log"

_configured = False


def setup_logger(log_level: str = "DEBUG", log_to_file: bool = True, force: bool = False) -> None:
    """
    Настройка системы логирования приложения.
    """
    global _configured
    if _configured and not force:
        return

    # Удаляем все существующие обработчики
    logger.remove()

    # Добавляем вывод в консоль (всегда)
    logger.add(
        sys.stderr,
        level=log_level,
        format="<green>{time:HH:mm:ss}</green> | <level>{level:<8}</level> | {message}",
        colorize=True,
    )

    # Добавляем файловый лог
    if log_to_file:
        try:
            LOG_DIR.mkdir(parents=True, exist_ok=True)
            logger.add(
                str(LOG_FILE),
                level="DEBUG",
                format="{time:YYYY-MM-DD HH:mm:ss} | {level:<8} | {name}:{line} | {message}",
                rotation="5 MB",
                retention=3,
                encoding="utf-8",
                enqueue=True,
            )
            logger.debug(f"Лог-файл: {LOG_FILE}")
        except Exception as e:
            logger.warning(f"Не удалось создать лог-файл: {e}")

    _configured = True
    logger.info(f"Логирование настроено. Уровень: {log_level}")

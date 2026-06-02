import sys
from pathlib import Path
from loguru import logger

# Определяем путь для логов в зависимости от окружения
if getattr(sys, 'frozen', False):
    # Запущено из .exe
    APP_DIR = Path(sys.executable).parent
    LOG_DIR = APP_DIR / "logs"
else:
    # Запущено из Python
    LOG_DIR = Path.home() / "AppData" / "Local" / "EMSpecGenerator" / "logs"

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

    # Добавляем вывод в консоль (для отладки)
    logger.add(
        sys.stderr,
        level=log_level,
        format="<green>{time:HH:mm:ss}</green> | <level>{level:<8}</level> | {message}",
        colorize=True,
    )

    # Добавляем файловый лог только если включено и папка доступна
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
                enqueue=True,  # Для работы в многопоточном режиме
            )
            logger.debug(f"Лог-файл: {LOG_FILE}")
        except Exception as e:
            # Если не можем писать в лог-файл, просто выводим предупреждение
            logger.warning(f"Не удалось создать лог-файл: {e}")

    _configured = True
    logger.info(f"Логирование настроено. Уровень: {log_level}")

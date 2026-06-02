import sys
import os
from pathlib import Path
from loguru import logger


def get_log_dir() -> Path:
    """Получить правильную директорию для логов в зависимости от окружения"""
    if getattr(sys, 'frozen', False):
        # Запущено как exe - пишем логи в папку рядом с exe
        exe_dir = Path(sys.executable).parent
        log_dir = exe_dir / "logs"
    else:
        # Режим разработки
        log_dir = Path.home() / "AppData" / "Local" / "EMSpecGenerator" / "logs"

    # Создаем директорию если её нет
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        # Если не можем создать, используем временную папку
        import tempfile
        log_dir = Path(tempfile.gettempdir()) / "EMSpecGenerator_logs"
        log_dir.mkdir(parents=True, exist_ok=True)

    return log_dir


LOG_DIR = get_log_dir()
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

    # Проверяем, доступен ли stderr
    if sys.stderr is not None:
        try:
            logger.add(
                sys.stderr,
                level=log_level,
                format="<green>{time:HH:mm:ss}</green> | <level>{level:<8}</level> | {message}",
                colorize=True,
            )
        except Exception:
            # Если stderr не работает, используем stdout
            if sys.stdout is not None:
                logger.add(
                    sys.stdout,
                    level=log_level,
                    format="{time:HH:mm:ss} | {level:<8} | {message}",
                )
    elif sys.stdout is not None:
        logger.add(
            sys.stdout,
            level=log_level,
            format="{time:HH:mm:ss} | {level:<8} | {message}",
        )
    else:
        # Если ни stderr ни stdout не доступны, создаем файл лога в текущей директории
        fallback_log = Path("fallback_log.txt")
        logger.add(
            str(fallback_log),
            level=log_level,
            format="{time:HH:mm:ss} | {level:<8} | {message}",
        )

    if log_to_file:
        try:
            # Убеждаемся, что директория существует
            LOG_DIR.mkdir(parents=True, exist_ok=True)

            logger.add(
                str(LOG_FILE),
                level="DEBUG",
                format="{time:YYYY-MM-DD HH:mm:ss} | {level:<8} | {name}:{line} | {message}",
                rotation="5 MB",
                retention=5,
                encoding="utf-8",
                backtrace=True,
                diagnose=True,
                enqueue=True,  # Добавляем для потокобезопасности
            )
            # Пробуем записать тестовое сообщение
            logger.debug(f"Лог-файл: {LOG_FILE}")
        except Exception as e:
            # Если не можем писать в файл, просто игнорируем
            pass

    _configured = True

    # Логируем информацию о окружении (если удалось)
    try:
        logger.info(f"Логгер настроен. Режим: {'Frozen (exe)' if getattr(sys, 'frozen', False) else 'Development'}")
        logger.info(f"Директория логов: {LOG_DIR}")
        logger.info(f"Текущая директория: {os.getcwd()}")
    except Exception:
        pass

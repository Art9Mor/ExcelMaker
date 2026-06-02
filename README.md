# ExcelMaker

Утилита на Python для автоматической обработки Excel-файлов `.xlsm` и генерации спецификации по данным первой вкладки.

## Возможности

- Читает первый лист Excel-книги.
- Парсит структуру спецификации.
- Формирует результат без ручного вмешательства.
- Поддерживает CLI и GUI.
- Сохраняет `.xlsm` через `openpyxl` с `keep_vba=True`.

## Требования

- Python 3.12+
- `openpyxl`
- `loguru`
- `typer`
- `PyQt6`

## Установка

```bash
pip install -r requirements.txt
```

Или через `uv`:

```bash
uv sync
```

## Запуск

### Build

```bash
pyinstaller --onefile --windowed --name ExcelMaker --add-data "src:src" main.py
```

или через скрипт:

```bash
python scripts/build.py
```

### CLI

```bash
python main.py путь/к/файлу.xlsm
```

С указанием выходного файла:

```bash
python main.py путь/к/файлу.xlsm --output result.xlsm
```

Или:

```bash
./dist/ExcelMaker --help
```

### GUI

```bash
python main.py --gui
```

Или:

```bash
./dist/ExcelMaker --gui
```

## Структура проекта

```text
src/
  core/
    models.py
    parser.py
    processor.py
    writer.py
  gui/
    main_window.py
  utils/
    logger.py
tests/
  test_parser.py
  test_writer.py
```

## Как это работает

1. Приложение открывает первую вкладку Excel-файла.
2. Находит таблицу спецификации.
3. Делит данные на разделы и позиции.
4. Записывает результат обратно в книгу.
5. Сохраняет файл.

## Тесты

Запуск тестов:

```bash
pytest
```

## Примечания

- Для `.xlsm` файл сохраняется с `keep_vba=True`.
- Проект не создаёт лишних листов.
- Первая вкладка остаётся основным источником данных.


Ошибка сборки:
Traceback (most recent call last):
  File "main.py", line 55, in <module>
    run_gui()
  File "main.py", line 38, in run_gui
    setup_logger(log_level="INFO", log_to_file=True)
  File "src\utils\logger.py", line 22, in setup_logger
    logger.add(
  File "loguru\_logger.py", line 872, in add
    raise TypeError("Cannot log to objects of type '%s'" % type(sink).__name__)
TypeError: Cannot log to objects of type 'NoneType'
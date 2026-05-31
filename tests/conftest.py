from pathlib import Path

import pytest
from openpyxl import Workbook


@pytest.fixture
def sample_workbook_path(tmp_path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист1"

    ws.append([
        "Структура", "Скрыть строку, символы /*", "Опция", "Значение",
        "Примечание 1", "Примечание 2", "Примечание 3", "Примечание 4"
    ])
    ws.append(["ДАННЫЕ", "/*", "Пункт 1"])
    ws.append(["Проект", "", "2055"])
    ws.append(["Тип оборудования", "", "ЯКНО"])
    ws.append([
        "Структура", "Скрыть строку, символы /*", "Наименование",
        "Цена", "Количество", "Стоиомсть, руб", "Стоимость с наценкой, руб", "Ед. изм."
    ])
    ws.append(["1", "", "Корпус"])
    ws.append(["", "", "1.1", "Каркас ЯКНО", "", "шт", 1])
    ws.append(["", "", "1.2", "Рама опорная", "", "шт", 1])
    ws.append(["2", "", "Отсек"])
    ws.append(["", "", "2.1", "Вакуумный выключатель", "", "шт", 1])
    ws.append(["ИТОГ", "", "", "", "", "", "", ""])

    path = tmp_path / "sample.xlsm"
    wb.save(path)
    return path

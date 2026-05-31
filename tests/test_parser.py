from openpyxl import Workbook

from src.core.parser import parse_workbook
from src.core.models import SpecDocument


def make_sample_workbook() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист1"

    ws.append(["Структура", "Скрыть строку, символы /*", "Опция", "Значение", "Примечание 1", "Примечание 2", "Примечание 3", "Примечание 4"])
    ws.append(["ДАННЫЕ", "/*", "Пункт 1"])
    ws.append(["Проект", "", "2055"])
    ws.append(["Тип оборудования", "", "ЯКНО"])
    ws.append(["Структура", "Скрыть строку, символы /*", "Наименование", "Цена", "Количество", "Стоиомсть, руб", "Стоимость с наценкой, руб", "Ед. изм."])
    ws.append(["1", "", "Корпус"])
    ws.append(["", "", "1.1", "Каркас ЯКНО", "", "шт", 1])
    ws.append(["", "", "1.2", "Рама опорная", "", "шт", 1])
    ws.append(["2", "", "Отсек"])
    ws.append(["", "", "2.1", "Вакуумный выключатель", "", "шт", 1])
    ws.append(["ИТОГ", "", "", "", "", "", "", ""])

    return wb


def test_parse_workbook_extracts_sections_and_items():
    wb = make_sample_workbook()
    doc = parse_workbook(wb, "Лист1")

    assert isinstance(doc, SpecDocument)
    assert doc.source_sheet == "Лист1"
    assert doc.target_sheet == "Лист1"
    assert len(doc.sections) == 2
    assert doc.total_items == 3

    assert doc.sections[0].title == "Корпус"
    assert doc.sections[0].items[0].number == "1.1"
    assert doc.sections[0].items[0].name == "Каркас ЯКНО"

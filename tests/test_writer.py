"""
Модуль с тестами для записи спецификации.
"""

import pytest
from src.core.models import SpecDocument
from src.core.writer import create_result_sheet, write_spec_to_sheet


class TestCreateResultSheet:
    """
    Тесты для функции create_result_sheet.
    """

    def test_create_result_sheet(self, empty_workbook):
        """
        Тест создания листа с результатом по умолчанию.
        """

        ws = create_result_sheet(empty_workbook)
        assert ws.title == "Спецификация"

    @pytest.mark.parametrize("sheet_name", [
        "Тест",
        "Мой лист",
        "Spec",
    ])
    def test_create_result_sheet_custom_name(self, empty_workbook, sheet_name):
        """
        Параметризованный тест создания листа с пользовательским названием.
        """

        ws = create_result_sheet(empty_workbook, sheet_name)
        assert ws.title == sheet_name


class TestWriteSpecToSheet:
    """
    Тесты для функции write_spec_to_sheet.
    """

    def test_write_empty_spec(self, empty_workbook):
        """
        Тест записи пустой спецификации.
        """

        doc = SpecDocument(source_sheet="ВВ", target_sheet="Спецификация")
        write_spec_to_sheet(empty_workbook, doc)
        assert "Спецификация" in empty_workbook.sheetnames

    def test_write_spec_with_section(self, empty_workbook, sample_spec_document):
        """
        Тест записи спецификации с разделом с использованием фикстуры.
        """

        write_spec_to_sheet(empty_workbook, sample_spec_document)

        ws = empty_workbook["Спецификация"]

        assert ws.cell(1, 1).value == "Техническая спецификация"

        equipment_row = None
        for row in range(1, 10):
            if ws.cell(row, 1).value == "ЯКНО-ВВ-6кВ":
                equipment_row = row
                break
        assert equipment_row is not None, "Название изделия не найдено"

        header_row = None
        for row in range(1, 20):
            if ws.cell(row, 1).value == "№ п/п":
                header_row = row
                break
        assert header_row is not None, "Заголовки таблицы не найдены"

        assert ws.cell(header_row, 1).value == "№ п/п"
        assert ws.cell(header_row, 2).value == "Наименование"
        assert ws.cell(header_row, 3).value == "Ед. изм."
        assert ws.cell(header_row, 4).value == "Кол-во"
        assert ws.cell(header_row, 5).value == "Сумма, руб"

        section_row = header_row + 1
        assert ws.cell(section_row, 1).value == "1"
        assert ws.cell(section_row, 2).value == "Корпус"

        assert ws.cell(section_row + 1, 1).value == "1.1"
        assert ws.cell(section_row + 1, 2).value == "Каркас ЯКНО"
        assert ws.cell(section_row + 1, 3).value == "шт"
        assert ws.cell(section_row + 1, 4).value == 1
        assert ws.cell(section_row + 1, 5).value == 80000

        assert ws.cell(section_row + 2, 1).value == "1.2"
        assert ws.cell(section_row + 2, 2).value == "Рама опорная"
        assert ws.cell(section_row + 2, 3).value == "шт"
        assert ws.cell(section_row + 2, 4).value == 1
        assert ws.cell(section_row + 2, 5).value == 5000
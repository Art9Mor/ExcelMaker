"""
Модуль с тестами для моделей данных.
"""

import pytest
from src.core.models import SpecItem, Section, SpecHeader, SpecDocument


class TestSpecItem:
    """
    Тесты для класса SpecItem.
    """

    @pytest.mark.parametrize("number,name,unit,quantity", [
        ("1.1", "Тестовая позиция", "шт", 2),
        ("2.3", "Другая позиция", "м", 5),
        ("3.0", "Позиция", "кг", 1),
    ])
    def test_spec_item_creation(self, number, name, unit, quantity):
        """
        Параметризованный тест создания позиции спецификации.
        """

        item = SpecItem(number=number, name=name, unit=unit, quantity=quantity)
        assert item.number == number
        assert item.name == name
        assert item.unit == unit
        assert item.quantity == quantity
        assert item.price == 0.0
        assert item.total == 0.0
        assert item.is_hidden is False

    def test_spec_item_with_price(self):
        """
        Тест создания позиции с ценой.
        """

        item = SpecItem(number="1.1", name="Тест", price=100, quantity=2)
        assert item.price == 100
        assert item.quantity == 2


class TestSection:
    """
    Тесты для класса Section.
    """

    def test_section_creation(self):
        """
        Тест создания раздела.
        """

        section = Section(number=1, title="Корпус")
        assert section.number == 1
        assert section.title == "Корпус"
        assert section.items == []
        assert section.section_total == 0.0

    def test_section_add_item(self):
        """
        Тест добавления позиции в раздел.
        """

        section = Section(number=1, title="Корпус")
        item = SpecItem(number="1.1", name="Каркас", unit="шт", quantity=1)
        section.items.append(item)
        section.section_total = 80000.0
        assert len(section.items) == 1
        assert section.section_total == 80000.0


class TestSpecHeader:
    """
    Тесты для класса SpecHeader.
    """

    def test_header_creation(self):
        """
        Тест создания заголовка без данных.
        """

        header = SpecHeader()
        assert header.project is None
        assert header.equipment_type is None
        assert header.doc_number is None
        assert header.customer is None
        assert header.calculation_date is None

    @pytest.mark.parametrize("project,equipment,customer", [
        ("2055", "ЯКНО-ВВ-6кВ", "ООО Соврудник"),
        ("3000", "КТП-1000", "ООО Тест"),
        ("1234", "Щит", "ИП Иванов"),
    ])
    def test_header_with_data(self, project, equipment, customer):
        """
        Параметризованный тест создания заголовка с данными.
        """

        header = SpecHeader(
            project=project,
            equipment_type=equipment,
            customer=customer
        )
        assert header.project == project
        assert header.equipment_type == equipment
        assert header.customer == customer


class TestSpecDocument:
    """
    Тесты для класса SpecDocument.
    """

    def test_document_creation(self):
        """
        Тест создания документа.
        """

        doc = SpecDocument(source_sheet="ВВ", target_sheet="Спецификация")
        assert doc.source_sheet == "ВВ"
        assert doc.target_sheet == "Спецификация"
        assert doc.sections == []
        assert doc.grand_total == 0.0

    def test_total_items_property(self, sample_spec_document):
        """
        Тест подсчёта общего количества позиций с использованием фикстуры.
        """

        assert sample_spec_document.total_items == 2
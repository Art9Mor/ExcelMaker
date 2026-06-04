"""
Модуль с тестами для парсера Excel-файлов.
"""

import pytest
from src.core.parser import _clean, _as_number, _find_header_row


class TestClean:
    """
    Тесты для функции _clean.
    """

    @pytest.mark.parametrize("input_value,expected", [
        (None, ""),
        ("  test  ", "test"),
        ("test\xa0string", "test string"),
        ("", ""),
        ("   ", ""),
    ])
    def test_clean(self, input_value, expected):
        """
        Параметризованный тест очистки строк.
        """

        assert _clean(input_value) == expected


class TestAsNumber:
    """
    Тесты для функции _as_number.
    """

    @pytest.mark.parametrize("input_value,expected", [
        (None, 0.0),
        (100, 100.0),
        (100.5, 100.5),
        ("100", 100.0),
        ("100.5", 100.5),
        ("100 ₽", 100.0),
        ("abc", 0.0),
        ("", 0.0),
    ])
    def test_as_number(self, input_value, expected):
        """
        Параметризованный тест преобразования в число.
        """

        assert _as_number(input_value) == expected


class TestFindHeaderRow:
    """
    Тесты для функции _find_header_row.
    """

    def test_find_header_row(self, workbook_with_headers):
        """
        Тест поиска строки с заголовками с использованием фикстуры.
        """

        ws = workbook_with_headers.active
        row_idx, col_map = _find_header_row(ws)
        assert row_idx == 1
        assert col_map['struct'] == 0
        assert col_map['name'] == 2
        assert col_map['price'] == 3
        assert col_map['qty'] == 4
        assert col_map['unit'] == 7

    def test_find_header_row_not_found(self, empty_workbook):
        """
        Тест поиска строки с заголовками когда её нет.
        """

        ws = empty_workbook.active
        with pytest.raises(ValueError, match="Не найден заголовок таблицы"):
            _find_header_row(ws)
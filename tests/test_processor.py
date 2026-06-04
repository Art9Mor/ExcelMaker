"""
Модуль с тестами для процессора.
"""

import pytest
from pathlib import Path
from src.core.models import SpecDocument
from src.core.processor import ProcessingResult


class TestProcessingResult:
    """
    Тесты для класса ProcessingResult.
    """

    @pytest.mark.parametrize("output_path", [
        Path("test.xlsx"),
        Path("/tmp/result.xlsx"),
        Path("C:\\test.xlsx"),
    ])
    def test_success_result(self, output_path):
        """
        Параметризованный тест успешного результата обработки.
        """

        doc = SpecDocument(source_sheet="Тест", target_sheet="Спецификация")

        result = ProcessingResult(
            success=True,
            doc=doc,
            output_path=output_path
        )
        assert result.success is True
        assert "✓" in result.summary()
        assert str(output_path) in result.summary()

    @pytest.mark.parametrize("error_message", [
        "Тестовая ошибка",
        "Файл не найден",
        "Ошибка доступа",
    ])
    def test_error_result(self, error_message):
        """
        Параметризованный тест результата с ошибкой.
        """

        result = ProcessingResult(success=False, error=error_message)
        assert result.success is False
        assert "Ошибка" in result.summary()
        assert error_message in result.summary()
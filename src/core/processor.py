from pathlib import Path

from loguru import logger

from .models import SpecDocument
from .parser import load_and_parse
from .writer import save_workbook, write_spec_to_sheet


class ProcessingResult:
    """
    Результат обработки Excel-файла.
    """

    def __init__(
        self,
        success: bool,
        doc: SpecDocument | None = None,
        output_path: Path | None = None,
        word_path: Path | None = None,
        error: str = "",
    ) -> None:
        """
        Инициализация результата обработки.
        """

        self.success = success
        self.doc = doc
        self.output_path = output_path
        self.word_path = word_path
        self.error = error

    def summary(self) -> str:
        """
        Формирование текстовой сводки результата обработки.
        """

        if not self.success:
            return f"Ошибка: {self.error}"

        d = self.doc
        lines = [
            "✓ Обработка завершена успешно",
            f"Исходный лист: {d.source_sheet}",
            f"Целевой лист: {d.target_sheet}",
            f"Разделов: {len(d.sections)}",
            f"Позиций: {d.total_items}",
        ]
        if self.output_path:
            lines.append(f"Файл Excel: {self.output_path}")
        if self.word_path:
            lines.append(f"Файл Word: {self.word_path}")
        return "\n".join(lines)


def process_file(input_path: Path, output_path: Path | None = None) -> ProcessingResult:
    """
    Обработка Excel-файла и создание спецификации.
    """

    if output_path is None:
        output_path = input_path.parent / f"{input_path.stem}_specification{input_path.suffix}"

    try:
        wb, doc = load_and_parse(input_path)
        write_spec_to_sheet(wb, doc)
        save_workbook(wb, output_path)
        return ProcessingResult(success=True, doc=doc, output_path=output_path)
    except Exception as e:
        logger.exception("Ошибка обработки")
        return ProcessingResult(success=False, error=str(e))

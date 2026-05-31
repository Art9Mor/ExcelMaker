from pathlib import Path

from loguru import logger

from .models import SpecDocument
from .parser import load_and_parse
from .writer import save_workbook, write_spec_to_sheet


class ProcessingResult:
    """
    Описание результата обработки Excel-файла.
    """

    def __init__(
        self,
        success: bool,
        doc: SpecDocument | None = None,
        output_path: Path | None = None,
        error: str = "",
    ) -> None:
        self.success = success
        self.doc = doc
        self.output_path = output_path
        self.error = error

    def summary(self) -> str:
        """
        Формирование текстовой сводки результата обработки.
        """

        if not self.success:
            return f"Ошибка: {self.error}"

        d = self.doc

        return "\n".join(
            [
                "✓ Обработка завершена успешно",
                f"Исходный лист: {d.source_sheet}",
                f"Целевой лист: {d.target_sheet}",
                f"Разделов: {len(d.sections)}",
                f"Позиций: {d.total_items}",
                f"Файл сохранён: {self.output_path}",
            ]
        )


def build_output_path(
    input_path: Path,
) -> Path:
    """
    Формирование пути выходного файла.
    """

    return input_path.with_name(
        f"{input_path.stem}_specification.xlsx"
    )


def process_file(
    input_path: Path,
    output_path: Path | None = None,
) -> ProcessingResult:
    """
    Обработка Excel-файла и сохранение результата.
    """

    if input_path.suffix.lower() != ".xlsm":
        raise ValueError(
            "Поддерживаются только файлы .xlsm"
        )

    output_path = (
        output_path
        or build_output_path(input_path)
    )

    try:
        wb, doc = load_and_parse(input_path)

        write_spec_to_sheet(
            wb,
            doc,
        )

        save_workbook(
            wb,
            output_path,
        )

        return ProcessingResult(
            success=True,
            doc=doc,
            output_path=output_path,
        )

    except Exception as e:
        logger.exception(
            "Ошибка обработки"
        )

        return ProcessingResult(
            success=False,
            error=str(e),
        )

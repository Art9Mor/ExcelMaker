from pathlib import Path

from loguru import logger
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .models import SpecDocument


def get_or_create_result_sheet(
    wb: Workbook,
    sheet_name: str,
) -> Worksheet:
    """
    Создание листа результата.
    """

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    return wb.create_sheet(sheet_name)


def _autosize_columns(
    ws: Worksheet,
) -> None:
    """
    Автоматический подбор ширины колонок.
    """

    for column in ws.columns:

        max_length = 0

        for cell in column:

            if cell.value is None:
                continue

            max_length = max(
                max_length,
                len(str(cell.value)),
            )

        column_letter = get_column_letter(
            column[0].column
        )

        ws.column_dimensions[
            column_letter
        ].width = max_length + 3


def write_spec_to_sheet(
    wb: Workbook,
    doc: SpecDocument,
) -> None:
    """
    Запись спецификации на отдельный лист.
    """

    ws = get_or_create_result_sheet(
        wb,
        doc.target_sheet,
    )

    row = 1

    ws.cell(
        row,
        1,
    ).value = "Техническая спецификация"

    row += 2

    if doc.header.project:

        ws.cell(
            row,
            1,
        ).value = "Проект"

        ws.cell(
            row,
            2,
        ).value = doc.header.project

        row += 1

    if doc.header.equipment_type:

        ws.cell(
            row,
            1,
        ).value = "Оборудование"

        ws.cell(
            row,
            2,
        ).value = doc.header.equipment_type

        row += 1

    row += 1

    ws.cell(row, 1).value = "№ п/п"
    ws.cell(row, 2).value = "Наименование"
    ws.cell(row, 3).value = "Ед. изм."
    ws.cell(row, 4).value = "Кол-во"

    row += 1

    for section in doc.sections:

        ws.cell(
            row,
            1,
        ).value = str(section.number)

        ws.cell(
            row,
            2,
        ).value = section.title

        row += 1

        for item in section.items:

            ws.cell(
                row,
                1,
            ).value = item.number

            ws.cell(
                row,
                2,
            ).value = item.name

            ws.cell(
                row,
                3,
            ).value = item.unit

            ws.cell(
                row,
                4,
            ).value = item.quantity

            row += 1

        row += 1

    _autosize_columns(ws)

    logger.info(
        f"Записано разделов: "
        f"{len(doc.sections)}"
    )


def save_workbook(
    wb: Workbook,
    file_path: Path,
) -> None:
    """
    Сохранение Excel-файла на диск.
    """

    wb.save(str(file_path))

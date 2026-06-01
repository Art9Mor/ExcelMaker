from pathlib import Path
from loguru import logger
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Alignment, Border, Side
from .models import SpecDocument, Section, SpecItem

# Константы колонок для выходного листа
COL_NUM = 3  # C
COL_NAME = 4  # D
COL_UNIT = 6  # F
COL_QTY = 7  # G
COL_TOTAL = 8  # H


def write_spec_to_sheet(wb: Workbook, doc: SpecDocument) -> None:
    target_name = doc.target_sheet
    if target_name not in wb.sheetnames:
        raise ValueError(f"Лист «{target_name}» не найден")

    ws = wb[target_name]

    # 1. Очистка старых данных (от строки 28 и ниже, чтобы не задеть шапку)
    # Находим строку с "ИТОГ" или "Структура" для определения конца
    last_row = 27
    for r in ws.iter_rows(values_only=True):
        if any(cell in ("Структура", "ИТОГ") for cell in r if cell):
            # Находим индекс этой строки
            pass

            # Для простоты очистим всё с 28 строки до конца
    for row in ws.iter_rows(min_row=28):
        for cell in row:
            cell.value = None
            cell.font = Font()
            cell.alignment = Alignment()
            cell.border = Border()

    # Стили
    bold_center_underline = Font(bold=True, underline="single")
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    curr_row = 28

    for section in doc.sections:
        # Заголовок раздела
        ws.cell(row=curr_row, column=COL_NUM).value = str(section.number)
        ws.cell(row=curr_row, column=COL_NAME).value = section.title

        # Форматирование заголовка как в VBA
        ws.cell(row=curr_row, column=COL_NUM).font = bold_center_underline
        ws.cell(row=curr_row, column=COL_NAME).font = bold_center_underline
        ws.cell(row=curr_row, column=COL_NUM).alignment = center_align
        ws.cell(row=curr_row, column=COL_NAME).alignment = center_align

        curr_row += 1

        for item in section.items:
            ws.cell(row=curr_row, column=COL_NUM).value = item.number
            ws.cell(row=curr_row, column=COL_NAME).value = item.name
            ws.cell(row=curr_row, column=COL_UNIT).value = item.unit
            ws.cell(row=curr_row, column=COL_QTY).value = item.quantity
            ws.cell(row=curr_row, column=COL_TOTAL).value = item.total

            # Если строка должна быть скрыта
            if item.is_hidden:
                ws.row_dimensions[curr_row].hidden = True

            curr_row += 1

        # ИТОГ ПО РАЗДЕЛУ
        ws.cell(row=curr_row, column=COL_NAME).value = "ИТОГ"
        ws.cell(row=curr_row, column=COL_TOTAL).value = section.section_total
        ws.cell(row=curr_row, column=COL_NAME).font = Font(bold=True)
        ws.cell(row=curr_row, column=COL_TOTAL).font = Font(bold=True)

        curr_row += 2  # Пропуск строки между разделами

    # ОБЩИЙ ИТОГ
    ws.cell(row=curr_row, column=COL_NAME).value = "ОБЩИЙ ИТОГ"
    ws.cell(row=curr_row, column=COL_TOTAL).value = doc.grand_total
    ws.cell(row=curr_row, column=COL_NAME).font = Font(bold=True)
    ws.cell(row=curr_row, column=COL_TOTAL).font = Font(bold=True)

    logger.info(f"Спецификация записана: {doc.total_items} поз. Сумма: {doc.grand_total}")


def save_workbook(wb: Workbook, file_path: Path) -> None:
    wb.save(str(file_path))
    logger.info(f"Файл сохранён: {file_path}")

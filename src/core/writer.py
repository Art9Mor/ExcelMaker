from pathlib import Path
from loguru import logger
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from .models import SpecDocument


def create_result_sheet(wb: Workbook, sheet_name: str = "Спецификация") -> Worksheet:
    """Создаёт новый лист для результата"""
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    return wb.create_sheet(sheet_name)


def write_spec_to_sheet(wb: Workbook, doc: SpecDocument) -> None:
    """
    Записывает спецификацию на новый отдельный лист
    """
    ws = create_result_sheet(wb, "Спецификация")

    # Стили
    title_font = Font(bold=True, size=14, name='Arial')
    header_font = Font(bold=True, size=11, name='Arial')
    section_font = Font(bold=True, size=11, name='Arial')
    total_font = Font(bold=True, size=11, name='Arial')
    normal_font = Font(size=10, name='Arial')

    center_align = Alignment(horizontal='center', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    right_align = Alignment(horizontal='right', vertical='center')

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    row = 1

    # Заголовок
    ws.merge_cells(f'A{row}:E{row}')
    ws.cell(row, 1, "Техническая спецификация")
    ws.cell(row, 1).font = title_font
    ws.cell(row, 1).alignment = center_align
    row += 2

    # Название изделия
    if doc.header and doc.header.equipment_type:
        ws.merge_cells(f'A{row}:E{row}')
        ws.cell(row, 1, doc.header.equipment_type)
        ws.cell(row, 1).font = Font(bold=True, size=12)
        ws.cell(row, 1).alignment = center_align
        row += 2

    # Шапка таблицы
    headers = ['№ п/п', 'Наименование', 'Ед. изм.', 'Кол-во', 'Сумма, руб']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row, col, header)
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    row += 1

    # Заполнение данных
    for section in doc.sections:
        # Заголовок секции
        ws.merge_cells(f'A{row}:E{row}')
        cell = ws.cell(row, 1, f"{section.number}  {section.title}")
        cell.font = section_font
        cell.alignment = left_align
        cell.border = thin_border
        row += 1

        # Позиции секции
        for item in section.items:
            ws.cell(row, 1, item.number).border = thin_border
            ws.cell(row, 1).alignment = center_align
            ws.cell(row, 1).font = normal_font

            ws.cell(row, 2, item.name).border = thin_border
            ws.cell(row, 2).alignment = left_align
            ws.cell(row, 2).font = normal_font

            ws.cell(row, 3, item.unit).border = thin_border
            ws.cell(row, 3).alignment = center_align
            ws.cell(row, 3).font = normal_font

            ws.cell(row, 4, item.quantity).border = thin_border
            ws.cell(row, 4).alignment = center_align
            ws.cell(row, 4).font = normal_font

            ws.cell(row, 5, item.total).border = thin_border
            ws.cell(row, 5).alignment = right_align
            ws.cell(row, 5).font = normal_font
            ws.cell(row, 5).number_format = '#,##0.00'

            row += 1

        # Итог по разделу (если есть позиции)
        if section.items and section.section_total > 0:
            ws.merge_cells(f'A{row}:D{row}')
            cell = ws.cell(row, 1, "ИТОГО по разделу")
            cell.font = total_font
            cell.alignment = right_align
            cell.border = thin_border

            ws.cell(row, 5, section.section_total).border = thin_border
            ws.cell(row, 5).font = total_font
            ws.cell(row, 5).alignment = right_align
            ws.cell(row, 5).number_format = '#,##0.00'
            row += 1

        row += 1  # Отступ после секции

    # Общий итог
    if doc.grand_total > 0:
        ws.merge_cells(f'A{row}:D{row}')
        cell = ws.cell(row, 1, "ОБЩИЙ ИТОГ")
        cell.font = Font(bold=True, size=12)
        cell.alignment = right_align
        cell.border = thin_border

        ws.cell(row, 5, doc.grand_total).border = thin_border
        ws.cell(row, 5).font = Font(bold=True, size=12)
        ws.cell(row, 5).alignment = right_align
        ws.cell(row, 5).number_format = '#,##0.00'

    # Автоширина
    for col in range(1, 6):
        max_len = 0
        for r in range(1, row + 5):
            val = ws.cell(r, col).value
            if val:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[get_column_letter(col)].width = min(max_len + 3, 50)

    logger.info(f"Создан лист 'Спецификация' с {doc.total_items} позициями, итого: {doc.grand_total:.2f}")


def save_workbook(wb: Workbook, file_path: Path) -> None:
    wb.save(str(file_path))
    logger.info(f"Файл сохранён: {file_path}")

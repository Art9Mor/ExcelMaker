from pathlib import Path
from loguru import logger
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Alignment, Border, Side

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
    left_top_align = Alignment(horizontal='left', vertical='top', wrap_text=True)
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
    ws.cell(row, 1).value = "Техническая спецификация"
    ws.cell(row, 1).font = title_font
    ws.cell(row, 1).alignment = center_align
    row += 2

    # Название изделия
    if doc.header and doc.header.equipment_type:
        ws.merge_cells(f'A{row}:E{row}')
        ws.cell(row, 1).value = doc.header.equipment_type
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
        ws.cell(row, 1).value = str(section.number)
        ws.cell(row, 2).value = section.title
        ws.cell(row, 1).font = section_font
        ws.cell(row, 2).font = section_font
        ws.cell(row, 1).alignment = center_align
        ws.cell(row, 2).alignment = center_align
        for col in [1, 2, 3, 4, 5]:
            ws.cell(row, col).border = thin_border
        row += 1

        # Позиции секции
        for item in section.items:
            # Записываем номер
            cell_num = ws.cell(row, 1, item.number)
            cell_num.border = thin_border
            cell_num.alignment = center_align
            cell_num.font = normal_font

            # Записываем наименование с СОХРАНЕНИЕМ переносов
            cell_name = ws.cell(row, 2, item.name)
            cell_name.border = thin_border
            cell_name.font = normal_font
            # ВАЖНО: включаем перенос текста
            cell_name.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

            # Если есть переносы строк, принудительно устанавливаем высоту строки
            if item.name and '\n' in item.name:
                lines_count = item.name.count('\n') + 1
                ws.row_dimensions[row].height = 12.75 * lines_count
                logger.debug(f"Установлена высота строки {row}: {ws.row_dimensions[row].height} для {lines_count} строк")

            # Ед. изм
            cell_unit = ws.cell(row, 3, item.unit)
            cell_unit.border = thin_border
            cell_unit.alignment = center_align
            cell_unit.font = normal_font

            # Количество
            cell_qty = ws.cell(row, 4, item.quantity)
            cell_qty.border = thin_border
            cell_qty.alignment = center_align
            cell_qty.font = normal_font

            # Сумма
            cell_total = ws.cell(row, 5, item.total)
            cell_total.border = thin_border
            cell_total.alignment = right_align
            cell_total.font = normal_font
            cell_total.number_format = '#,##0.00'

            row += 1

        # Итог по разделу
        if section.section_total > 0:
            ws.merge_cells(f'A{row}:D{row}')
            cell = ws.cell(row, 1, "ИТОГО по разделу")
            cell.border = thin_border
            cell.font = total_font
            cell.alignment = right_align

            cell_total = ws.cell(row, 5, section.section_total)
            cell_total.border = thin_border
            cell_total.font = total_font
            cell_total.alignment = right_align
            cell_total.number_format = '#,##0.00'
            row += 1

        row += 1

    # Общий итог
    if doc.grand_total > 0:
        ws.merge_cells(f'A{row}:D{row}')
        cell = ws.cell(row, 1, "ОБЩИЙ ИТОГ")
        cell.border = thin_border
        cell.font = Font(bold=True, size=12)
        cell.alignment = right_align

        cell_total = ws.cell(row, 5, doc.grand_total)
        cell_total.border = thin_border
        cell_total.font = Font(bold=True, size=12)
        cell_total.alignment = right_align
        cell_total.number_format = '#,##0.00'

    # Ширина колонок
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 80
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 18

    # Дополнительно: для всех ячеек в колонке B включаем перенос текста
    for r in range(1, row + 1):
        cell = ws.cell(r, 2)
        if cell.value:
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

    logger.info(f"Создан лист 'Спецификация' с {doc.total_items} позициями, итого: {doc.grand_total:.2f}")


def save_workbook(wb: Workbook, file_path: Path) -> None:
    """Сохраняет workbook"""
    try:
        file_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(str(file_path))
        if file_path.exists():
            logger.info(f"Файл сохранён: {file_path} (размер: {file_path.stat().st_size} байт)")
        else:
            logger.error(f"Файл НЕ создан: {file_path}")
    except Exception as e:
        logger.error(f"Ошибка сохранения: {e}")
        raise

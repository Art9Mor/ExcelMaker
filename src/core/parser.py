from pathlib import Path
from loguru import logger
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from .models import SpecDocument, Section, SpecItem, SpecHeader

STRUCTURE_HEADER = "Структура"
RESULT_MARKER = "ИТОГ"

# Колонки в исходном листе (0-based)
COL_STRUCT = 0  # A — Структура
COL_HIDE = 1  # B — Скрыть строку (/*)
COL_NAME = 2  # C — Наименование
COL_PRICE = 3  # D — Цена, руб
COL_QTY = 4  # E — Количество
COL_UNIT = 7  # H — Ед. изм.


def _clean(v) -> str:
    return str(v).replace("\xa0", " ").strip() if v is not None else ""


def _as_number(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = _clean(v).replace(" ", "").replace(",", ".").replace("₽", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def _find_data_start(ws) -> int:
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx > 100:
            break
        if row and len(row) > COL_NAME:
            if row[COL_STRUCT] == STRUCTURE_HEADER and row[COL_NAME] == "Наименование":
                logger.debug(f"Найден заголовок таблицы в строке {row_idx}")
                return row_idx + 1
    raise ValueError("Не найден заголовок таблицы спецификации")


def _find_project_info(ws) -> SpecHeader:
    header = SpecHeader()

    for row in ws.iter_rows(values_only=True, max_row=30):
        if not row or len(row) < 4:
            continue

        param = _clean(row[2]) if len(row) > 2 else ""
        value = _clean(row[3]) if len(row) > 3 else ""

        if param == "Проект":
            header.project = value
        elif param == "Наименование изделия":
            header.equipment_type = value
        elif param == "Заказчик":
            header.customer = value
        elif param == "Шифр документации":
            header.doc_number = value

    logger.info(f"Проект: {header.project}, Оборудование: {header.equipment_type}")
    return header


def parse_workbook(wb: Workbook, source_name: str) -> SpecDocument:
    ws = wb[source_name]
    header = _find_project_info(ws)

    doc = SpecDocument(
        source_sheet=source_name,
        target_sheet="Спецификация",
        header=header
    )

    start_row = _find_data_start(ws)
    logger.info(f"Начало данных: строка {start_row}")

    current_section = None
    section_counter = 0
    item_counter = 0

    section_names = ["Корпус", "Отсек высоковольтного выключателя", "Отсек РЗА", "Прочее"]

    for row_idx in range(start_row, min(ws.max_row + 1, 200)):
        row = ws[row_idx]
        if not row or len(row) < 5:
            continue

        struct_val = _clean(row[COL_STRUCT].value) if row[COL_STRUCT] and row[COL_STRUCT].value else ""
        hide_val = _clean(row[COL_HIDE].value) if row[COL_HIDE] and row[COL_HIDE].value else ""
        name_val = _clean(row[COL_NAME].value) if row[COL_NAME] and row[COL_NAME].value else ""
        price_val = row[COL_PRICE].value if row[COL_PRICE] and row[COL_PRICE].value else None
        qty_val = row[COL_QTY].value if row[COL_QTY] and row[COL_QTY].value else None
        unit_val = _clean(row[COL_UNIT].value) if len(row) > COL_UNIT and row[COL_UNIT] and row[
            COL_UNIT].value else "шт"

        if hide_val == "/*":
            continue

        # Пропускаем служебную строку со "Структура"
        if struct_val == STRUCTURE_HEADER:
            continue

        if struct_val == RESULT_MARKER or name_val == RESULT_MARKER:
            break

        if struct_val and struct_val in section_names:
            section_counter += 1
            current_section = Section(number=section_counter, title=struct_val)
            doc.sections.append(current_section)
            item_counter = 0
            logger.info(f"Секция {section_counter}: {struct_val} (строка {row_idx})")
            continue

        if name_val and current_section:
            has_qty = qty_val is not None and qty_val != 0 and qty_val != ""
            if has_qty:
                qty = _as_number(qty_val)
                price = _as_number(price_val)
                total = price * qty

                item_counter += 1
                number = f"{section_counter}.{item_counter}"
                name_clean = ' '.join(name_val.split())
                if len(name_clean) > 100:
                    name_clean = name_clean[:97] + "..."

                item = SpecItem(
                    number=number,
                    name=name_clean,
                    unit=unit_val if unit_val and unit_val != "None" else "шт",
                    quantity=qty,
                    price=price,
                    total=total,
                    is_hidden=False
                )

                current_section.items.append(item)
                current_section.section_total += total
                doc.grand_total += total

    total_items = sum(len(s.items) for s in doc.sections)
    logger.info(f"Парсинг завершён: {len(doc.sections)} разделов, {total_items} позиций, итого: {doc.grand_total:.2f}")

    for section in doc.sections:
        logger.info(
            f"  {section.number}. {section.title}: {len(section.items)} позиций, сумма: {section.section_total:.2f}")

    return doc


def load_and_parse(file_path: Path) -> tuple[Workbook, SpecDocument]:
    wb = load_workbook(str(file_path), data_only=True, keep_vba=True)
    source_name = wb.sheetnames[0]
    logger.info(f"Загружен лист: {source_name}")
    doc = parse_workbook(wb, source_name)
    return wb, doc

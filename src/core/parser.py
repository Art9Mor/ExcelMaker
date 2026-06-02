from pathlib import Path
from loguru import logger
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from .models import SpecDocument, Section, SpecItem, SpecHeader

STRUCTURE_HEADER = "Структура"
RESULT_MARKER = "ИТОГ"

# Настройки
SECTION_TITLES = {
    "Корпус": "Корпус",
    "Отсек высоковольтного выключателя": "Отсек высоковольтного выключателя",
    "Отсек РЗА": "Отсек РЗА",
    "Прочее": "Прочее",
    "Дополнительно": "Прочее",
}

SECTION_ORDER = ["Корпус", "Отсек высоковольтного выключателя", "Отсек РЗА", "Прочее"]


def _clean(v) -> str:
    if v is None:
        return ""
    return str(v).replace("\xa0", " ").strip()


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


def _find_header_row(ws) -> tuple[int, dict]:
    """Находит строку с заголовками и определяет колонки"""
    for row_idx in range(1, min(50, ws.max_row + 1)):
        row = ws[row_idx]
        if not row:
            continue

        col_map = {}
        for col_idx in range(min(10, len(row))):
            cell = row[col_idx]
            val = cell.value if hasattr(cell, 'value') else cell
            if val is None:
                continue
            val_str = _clean(val)

            if val_str == "Структура":
                col_map['struct'] = col_idx
            elif val_str == "Наименование":
                col_map['name'] = col_idx
            elif "Цена" in val_str and "руб" in val_str:
                col_map['price'] = col_idx
            elif val_str == "Количество":
                col_map['qty'] = col_idx
            elif val_str == "Ед. изм.":
                col_map['unit'] = col_idx

        if 'name' in col_map and 'qty' in col_map:
            logger.info(f"Найдены колонки в строке {row_idx}: {col_map}")
            return row_idx, col_map

    raise ValueError("Не найден заголовок таблицы")


def _find_project_info(ws) -> SpecHeader:
    header = SpecHeader()
    for row in ws.iter_rows(values_only=True, max_row=40):
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
    return header


def parse_workbook(ws) -> SpecDocument:
    """
    Парсит лист Excel и извлекает спецификацию (без загрузки всего workbook)
    """
    header = _find_project_info(ws)

    doc = SpecDocument(
        source_sheet=ws.title,
        target_sheet="Спецификация",
        header=header
    )

    header_row, col_map = _find_header_row(ws)
    logger.info(f"Заголовок в строке {header_row}, колонки: {col_map}")

    current_section = None
    section_counter = 0
    item_counter = 0

    struct_col = col_map.get('struct', 0)
    name_col = col_map.get('name', 2)
    price_col = col_map.get('price')
    qty_col = col_map.get('qty', 4)
    unit_col = col_map.get('unit')

    for row_idx in range(header_row + 1, min(ws.max_row + 1, 500)):
        row = ws[row_idx]
        if not row:
            continue

        struct_val = _clean(row[struct_col].value) if struct_col < len(row) and row[struct_col] and row[
            struct_col].value else ""
        name_val = _clean(row[name_col].value) if name_col < len(row) and row[name_col] and row[name_col].value else ""
        price_val = row[price_col].value if price_col and price_col < len(row) and row[price_col] else None
        qty_val = row[qty_col].value if qty_col < len(row) and row[qty_col] else None
        unit_val = _clean(row[unit_col].value) if unit_col and unit_col < len(row) and row[unit_col] and row[
            unit_col].value else "шт"

        if struct_val == STRUCTURE_HEADER:
            continue

        if struct_val == RESULT_MARKER or name_val == RESULT_MARKER:
            break

        # Определение секции
        is_section = struct_val and struct_val not in ["", "None"] and not name_val
        if is_section:
            section_counter += 1
            section_title = SECTION_TITLES.get(struct_val, struct_val)
            current_section = Section(number=section_counter, title=section_title)
            doc.sections.append(current_section)
            item_counter = 0
            logger.info(f"Секция {section_counter}: {section_title}")
            continue

        # Позиция
        if name_val and current_section:
            has_qty = qty_val is not None and qty_val != 0 and qty_val != ""
            if has_qty:
                qty = _as_number(qty_val)
                price = _as_number(price_val) if price_val else 0
                total = price * qty
                item_counter += 1
                number = f"{section_counter}.{item_counter}"
                name_clean = name_val.strip()

                item = SpecItem(
                    number=number,
                    name=name_clean,
                    unit=unit_val,
                    quantity=qty,
                    price=price,
                    total=total,
                )
                current_section.items.append(item)
                current_section.section_total += total
                doc.grand_total += total

    # Сортировка секций
    if SECTION_ORDER:
        doc.sections.sort(
            key=lambda s: SECTION_ORDER.index(s.title) if s.title in SECTION_ORDER else len(SECTION_ORDER))
        for idx, section in enumerate(doc.sections, 1):
            section.number = idx
            for item_idx, item in enumerate(section.items, 1):
                item.number = f"{idx}.{item_idx}"

    logger.info(
        f"Парсинг завершён: {len(doc.sections)} разделов, {doc.total_items} позиций, итого: {doc.grand_total:.2f}")
    return doc


def load_and_parse(file_path: Path) -> tuple[Workbook, SpecDocument]:
    """Загружает только первый лист для чтения данных, возвращает новый workbook для записи"""
    # Загружаем исходный файл только для чтения данных
    wb_read = load_workbook(str(file_path), data_only=True)
    source_ws = wb_read.worksheets[0]

    # Парсим данные
    doc = parse_workbook(source_ws)

    # Создаём НОВЫЙ workbook (чистый, без лишних листов)
    wb_write = Workbook()
    # Удаляем стандартный пустой лист (он будет создан заново при записи)
    default_sheet = wb_write.active
    wb_write.remove(default_sheet)

    wb_read.close()

    return wb_write, doc

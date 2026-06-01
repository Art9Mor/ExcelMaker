from pathlib import Path
from loguru import logger
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from .models import SpecDocument, Section, SpecItem, SpecHeader

STRUCTURE_HEADER = "Структура"
RESULT_MARKER = "ИТОГ"

# Централизованный словарь соответствия названий разделов
SECTION_TITLES = {
    "Корпус": "Корпус",
    "Отсек высоковольтного выключателя": "Отсек высоковольтного выключателя",
    "Отсек РЗА": "Отсек РЗА",
    "Прочее": "Прочее",
    "Дополнительно": "Прочее",
}

SECTION_ORDER = ["Корпус", "Отсек высоковольтного выключателя", "Отсек РЗА", "Прочее"]


def _clean(v) -> str:
    """Очищает значение, но СОХРАНЯЕТ переносы строк"""
    if v is None:
        return ""
    # Не удаляем \n и не заменяем их на пробелы
    return str(v).replace("\xa0", " ").strip() if isinstance(v, str) else str(v)


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
    """Автоматически определяет колонки по заголовкам"""
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx > 50:
            break
        if not row:
            continue

        col_map = {}
        for col_idx, val in enumerate(row):
            if val is None:
                continue
            val_str = _clean(val)
            if val_str == "Структура":
                col_map['struct'] = col_idx
            elif "Наименование" in val_str:
                col_map['name'] = col_idx
            elif "Цена" in val_str and "руб" in val_str:
                col_map['price'] = col_idx
            elif val_str == "Количество":
                col_map['qty'] = col_idx
            elif "Ед. изм" in val_str:
                col_map['unit'] = col_idx
            elif val_str == "Скрыть строку" or val_str == "Скрыть строку, символы /*":
                col_map['hide'] = col_idx

        required = ['struct', 'name', 'price', 'qty', 'unit']
        if all(k in col_map for k in required):
            logger.info(f"Найдены колонки в строке {row_idx}: {col_map}")
            return row_idx + 1, col_map

    raise ValueError("Не удалось определить структуру таблицы")


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


def parse_workbook(wb: Workbook, source_name: str) -> SpecDocument:
    ws = wb[source_name]
    header = _find_project_info(ws)

    doc = SpecDocument(
        source_sheet=source_name,
        target_sheet="Спецификация",
        header=header
    )

    start_row, col_map = _find_header_row(ws)
    logger.info(f"Начало данных: строка {start_row}")

    current_section = None
    section_counter = 0
    item_counter = 0

    for row_idx in range(start_row, min(ws.max_row + 1, 500)):
        row = ws[row_idx]
        if not row or len(row) < max(col_map.values()):
            continue

        # Читаем значения БЕЗ очистки от переносов
        struct_raw = row[col_map['struct']].value if row[col_map['struct']] else None
        hide_raw = row[col_map.get('hide', 1)].value if col_map.get('hide') and row[col_map['hide']] else None
        name_raw = row[col_map['name']].value if row[col_map['name']] else None
        price_val = row[col_map['price']].value if row[col_map['price']] else None
        qty_val = row[col_map['qty']].value if row[col_map['qty']] else None
        unit_raw = row[col_map['unit']].value if row[col_map['unit']] else None

        struct_val = _clean(struct_raw) if struct_raw else ""
        hide_val = _clean(hide_raw) if hide_raw else ""
        name_val = str(name_raw) if name_raw else ""
        unit_val = _clean(unit_raw) if unit_raw else "шт"

        # Пропускаем только пустые строки, но НЕ строки с переносами
        if not name_val and not struct_val:
            continue

        if struct_val == STRUCTURE_HEADER:
            continue

        if struct_val == RESULT_MARKER or name_val == RESULT_MARKER:
            break

        is_section = struct_val and struct_val not in ["", "None"] and not name_val
        if is_section:
            section_counter += 1
            section_title = SECTION_TITLES.get(struct_val, struct_val)
            current_section = Section(number=section_counter, title=section_title)
            doc.sections.append(current_section)
            item_counter = 0
            logger.info(f"Секция {section_counter}: {section_title}")
            continue

        if name_val and current_section:
            has_qty = qty_val is not None and qty_val != 0 and qty_val != ""
            if has_qty:
                qty = _as_number(qty_val)
                price = _as_number(price_val)
                total = price * qty
                item_counter += 1
                number = f"{section_counter}.{item_counter}"

                # Сохраняем имя КАК ЕСТЬ (с переносами)
                # Не удаляем пробелы подряд, не меняем структуру
                name_clean = name_val.strip()

                # Логируем для отладки
                if "Механические" in name_clean:
                    logger.info(f"Найдены Механические блокировки: {repr(name_clean[:100])}")
                    logger.info(f"Содержит \\n: {'\\n' in name_clean}")
                    logger.info(f"Содержит chr(10): {chr(10) in name_clean}")
                    logger.info(f"Количество переносов: {name_clean.count(chr(10))}")

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
    wb = load_workbook(str(file_path), data_only=True, keep_vba=True)
    source_name = wb.sheetnames[0]
    doc = parse_workbook(wb, source_name)
    return wb, doc

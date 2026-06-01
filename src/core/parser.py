from pathlib import Path
from loguru import logger
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from .models import SpecDocument, Section, SpecItem

STRUCTURE_HEADER = "Структура"
RESULT_MARKER = "ИТОГ"

# Колонки (0-based)
SP_COL_HIDE = 1  # B — маркер /*
SP_COL_NUM = 2  # C — № п/п
SP_COL_NAME = 3  # D — Наименование
SP_COL_PRICE = 4  # E — Цена
SP_COL_QTY = 5  # F — Кол-во
SP_COL_UNIT = 8  # I — Ед. изм. (в вашем exc.txt это 8-я колонка)


def _clean(v) -> str:
    return str(v).replace("\xa0", " ").strip() if v is not None else ""


def _as_number(v) -> float:
    s = _clean(v).replace(" ", "").replace(",", ".").replace("₽", "")
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def _find_header_row(ws, marker: str) -> int:
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        vals = [_clean(v) for v in row]
        if marker in vals:
            return i
    raise ValueError(f"Заголовок {marker} не найден")


def parse_workbook(wb: Workbook, source_name: str) -> SpecDocument:
    spec_name = f"Спека {source_name}"
    if spec_name not in wb.sheetnames:
        raise ValueError(f"Лист «{spec_name}» не найден")

    ws_vv = wb[source_name]
    ws_spec = wb[spec_name]

    doc = SpecDocument(source_sheet=source_name, target_sheet=spec_name)

    sp_header = _find_header_row(ws_spec, "п/п")
    current_section: Section | None = None
    section_counter = 0

    for row_idx, row in enumerate(ws_spec.iter_rows(values_only=True), 1):
        if row_idx <= sp_header:
            continue

        vals = [_clean(v) for v in row]
        if not vals or not vals[0]:
            continue

        if vals[0] in (RESULT_MARKER, STRUCTURE_HEADER):
            break

        # Проверка: это заголовок раздела (целое число) или позиция (число с точкой)
        num_raw = vals[SP_COL_NUM]

        if num_raw and "." not in num_raw and num_raw.isdigit():
            section_counter += 1
            current_section = Section(
                number=section_counter,
                title=vals[SP_COL_NAME] if vals[SP_COL_NAME] else f"Раздел {section_counter}"
            )
            doc.sections.append(current_section)
            logger.debug(f"Раздел {section_counter}: {current_section.title}")
            continue

        if current_section is not None:
            # Считаем цену и количество
            price = _as_number(vals[SP_COL_PRICE])
            qty = _as_number(vals[SP_COL_QTY])
            total = price * qty

            # Определяем, скрыта ли строка (наличие /* в колонке B)
            is_hidden = "/*" in vals[SP_COL_HIDE]

            item = SpecItem(
                number=num_raw,
                name=vals[SP_COL_NAME],
                unit=vals[SP_COL_UNIT] if len(vals) > SP_COL_UNIT else "шт",
                quantity=qty,
                price=price,
                total=total,
                is_hidden=is_hidden
            )

            current_section.items.append(item)
            current_section.section_total += total
            doc.grand_total += total

    logger.info(f"Парсинг завершен: {len(doc.sections)} разделов, {doc.total_items} позиций.")
    return doc


def load_and_parse(file_path: Path) -> tuple[Workbook, SpecDocument]:
    wb = load_workbook(str(file_path), data_only=True, keep_vba=True)
    source_name = wb.sheetnames[0]
    doc = parse_workbook(wb, source_name)
    return wb, doc

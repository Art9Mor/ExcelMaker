from pathlib import Path

from loguru import logger
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from .models import SpecDocument, Section, SpecItem

HIDE_MARKER = "/*"
STRUCTURE_HEADER = "Структура"
RESULT_MARKER = "ИТОГ"
RESULT_SHEET_NAME = "Спецификация"


def _clean(v) -> str:
    """
    Приведение значения ячейки к строке без лишних пробелов.
    """

    return str(v).replace("\xa0", " ").strip() if v is not None else ""


def _as_number(v):
    """
    Преобразование значения ячейки в число.
    """

    s = _clean(v).replace(" ", "").replace(",", ".").replace("₽", "")

    if not s:
        return 1

    try:
        n = float(s)
        return int(n) if n.is_integer() else n

    except ValueError:
        return 1


def _is_section_marker(num: str, name: str) -> bool:
    """
    Определение строки начала раздела спецификации.
    """

    return (
        bool(num)
        and "." not in num
        and name
        and name != RESULT_MARKER
        and name != STRUCTURE_HEADER
    )


def _find_table_header(ws) -> int:
    """
    Поиск строки заголовка таблицы спецификации.
    """

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        vals = [_clean(v) for v in row]

        if STRUCTURE_HEADER in vals and any(
            "Наименование" in v for v in vals
        ):
            return row_idx

    raise ValueError(
        "Не найден заголовок таблицы спецификации"
    )


def _pick_num_name_unit_qty(
    vals: list[str],
) -> tuple[str, str, str, str]:
    """
    Извлечение номера, наименования, единицы измерения и количества.
    """

    num = ""
    name = ""
    unit = ""
    qty = ""

    for v in vals:

        if not v:
            continue

        if not num and (
            v.isdigit()
            or "." in v and v[0].isdigit()
        ):
            num = v
            continue

        if (
            not name
            and not v.isdigit()
            and "." not in v
            and v
            not in (
                HIDE_MARKER,
                STRUCTURE_HEADER,
                RESULT_MARKER,
            )
        ):
            name = v
            continue

    if len(vals) >= 3 and not name:
        name = (
            vals[2]
            or vals[3]
            if len(vals) > 3
            else vals[2]
        )

    if len(vals) >= 6:
        unit = vals[5]

    if len(vals) >= 7:
        qty = vals[6]

    return num, name, unit, qty


def parse_workbook(
    wb_read: Workbook,
    source_sheet_name: str,
) -> SpecDocument:
    """
    Парсинг первого листа Excel в объект спецификации.
    """

    ws = wb_read[source_sheet_name]

    doc = SpecDocument(
        source_sheet=source_sheet_name,
        target_sheet=RESULT_SHEET_NAME,
    )

    header_row = _find_table_header(ws)

    logger.info(
        f"Заголовок таблицы найден: строка {header_row}"
    )

    current_section = None
    section_counter = 0
    in_table = False

    for row_idx, row in enumerate(
        ws.iter_rows(values_only=True),
        start=1,
    ):
        vals = [_clean(v) for v in row]

        if row_idx <= header_row:
            continue

        if not in_table:

            if any(
                v == STRUCTURE_HEADER
                for v in vals
            ):
                in_table = True

            continue

        if any(
            v == RESULT_MARKER
            for v in vals
        ):
            break

        if any(
            v == HIDE_MARKER
            for v in vals
        ):
            continue

        num, name, unit, qty = (
            _pick_num_name_unit_qty(vals)
        )

        if _is_section_marker(num, name):

            section_counter += 1

            current_section = Section(
                number=section_counter,
                title=name or num,
            )

            doc.sections.append(
                current_section
            )

            continue

        if (
            not num
            or "." not in num
            or current_section is None
        ):
            continue

        current_section.items.append(
            SpecItem(
                number=num,
                name=name,
                unit=unit or "шт",
                quantity=_as_number(qty),
            )
        )

    logger.info(
        f"Парсинг завершён: "
        f"{len(doc.sections)} разделов, "
        f"{doc.total_items} позиций"
    )

    return doc


def load_and_parse(
    file_path: Path,
) -> tuple[Workbook, SpecDocument]:
    """
    Загрузка Excel-файла и построение объекта спецификации.
    """

    wb_read = load_workbook(
        str(file_path),
        data_only=True,
    )

    source_name = wb_read.sheetnames[0]

    doc = parse_workbook(
        wb_read,
        source_name,
    )

    wb_read.close()

    wb_write = load_workbook(
        str(file_path),
        keep_vba=True,
    )

    return wb_write, doc

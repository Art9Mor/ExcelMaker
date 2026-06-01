#!/usr/bin/env python3
from pathlib import Path
from openpyxl import load_workbook

file_path = "/home/amshegar/Загрузки/Attachments_hhz@gknzo.ru_2026-05-27_15-57-58/2055_ ЯКНО-ВВ(ВК)-6кВ_Соврудник.xlsm"
wb = load_workbook(file_path, data_only=True)
ws = wb.worksheets[0]

print(f"Лист: {ws.title}")
print("=" * 80)

# Найдем строку с заголовком
header_row = 0
for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
    if row and len(row) > 2:
        if row[0] == "Структура" and row[2] == "Наименование":
            header_row = row_idx
            print(f"Заголовок найден в строке {header_row}")
            break

if header_row == 0:
    print("Заголовок не найден!")
    exit(1)

start_row = header_row + 1
print(f"Начало данных: строка {start_row}")
print("=" * 80)

section_names = ["Корпус", "Отсек высоковольтного выключателя", "Отсек РЗА", "Прочее"]
current_section = None
section_counter = 0
item_counter = 0

for row_idx in range(start_row, min(ws.max_row + 1, 100)):
    row = ws[row_idx]
    if not row or len(row) < 5:
        continue

    # Используем values_only=True при итерации
    # Но для прямого доступа нужно использовать .value
    struct_val = str(row[0].value).strip() if row[0] and row[0].value else ""
    name_val = str(row[2].value).strip() if row[2] and row[2].value else ""
    price_val = row[3].value if row[3] and row[3].value else None
    qty_val = row[4].value if row[4] and row[4].value else None
    unit_val = str(row[7].value).strip() if len(row) > 7 and row[7] and row[7].value else "шт"

    print(f"Строка {row_idx}: struct='{struct_val}', name='{name_val[:40] if name_val else ''}', qty={qty_val}")

    # Проверка на секцию
    if struct_val and struct_val in section_names:
        section_counter += 1
        current_section = struct_val
        item_counter = 0
        print(f"  -> НАЙДЕНА СЕКЦИЯ {section_counter}: {struct_val}")
        continue

    # Проверка на позицию
    if name_val and current_section:
        has_qty = qty_val is not None and qty_val != 0 and qty_val != ""
        if has_qty:
            item_counter += 1
            number = f"{section_counter}.{item_counter}"
            print(f"  -> ПОЗИЦИЯ {number}: {name_val[:50]}... x{qty_val} {unit_val}")

    if not struct_val and not name_val:
        continue

print("=" * 80)
print(f"Итого секций: {section_counter}, позиций: {item_counter}")
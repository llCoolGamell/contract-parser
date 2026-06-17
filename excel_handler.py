# -*- coding: utf-8 -*-
"""
Запись данных контрактов в Excel по шаблону «кодирование».
Книга: Лист1 (пусто), Лист2 (пусто), Лист3 (данные), Лист4 (справочник поставщиков).
Лист3: строка заголовков + строка «Год фин-я 2026 / было» + строки данных.
Проверочные столбцы P = L×J (цена×кол-во), Q = M−P (расхождение с «Общей ценой»).
Есть защита от повторной записи одинаковых строк в один файл.
"""
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from parser_engine import ContractData

HEADERS = [
    "Собственник товара", "тип", "код ЛС льготный", "№ договора", "№ аукциона",
    "Международное непатентованное название (МНН)", "Торговое наименование",
    "Лек. Форма   (форма выпуска), дозировка", "Ед. изм.", "Общее кол-во, ед. изм.",
    "Производитель", "Цена ед. продукции, руб. в месте поставки ", "Общая цена, руб.\n",
    "поставщик", "способ размещения",
]
COLUMN_WIDTHS = [15, 18, 15, 30, 25, 30, 25, 50, 10, 15, 45, 20, 20, 25, 25, 20, 20]

SHEET_NAME = "Лист3"
HEADER_ROW = 1
SUBHEADER_ROW = 2          # «Год фин-я 2026 / было»
DATA_START_ROW = 3

COL_DOSAGE = 8
COL_QTY = 10
COL_UNIT_PRICE = 12
COL_TOTAL = 13
COL_P = 16
COL_Q = 17

SUPPLIERS = [
    'АО "Р-Фарм"', 'АО "Р-Фарм"', 'АО «Фармимэкс» ', 'ООО "БСС"', 'ООО "БСС"',
    'ООО "Компания "Эталон"', 'ООО "Компания "Эталон"', 'ООО "Компания "Эталон"',
    'ООО "КОРС"', 'ООО "СВИЧ-СТОР"', 'ООО "СВИЧ-СТОР"',
    'ООО "Универсал Мед Сервис"', 'ООО "Универсал Мед Сервис"',
    'ООО "Универсал Мед Сервис"', 'ООО "Универсал Мед Сервис"',
    'ООО "ФАРМЛОГИСТИКА"', 'ООО "ФАРМЛОГИСТИКА"', 'ООО "ФАРМЛОГИСТИКА"',
    'ООО "ФАРМЛОГИСТИКА"', 'ООО "ФАРМРУБЕЖ"', 'ООО "ФАРМРУБЕЖ"',
    'ООО АВК-Альянс', 'ООО АВК-Альянс', 'ООО АВК-Альянс',
]

RED_FILL = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
ORANGE_FILL = PatternFill(start_color="FFB347", end_color="FFB347", fill_type="solid")


def _owner_abbreviation(n):
    u = (n or "").upper()
    for k, a in {"МИНЗДРАВ": "МЗ", "МИНИСТЕРСТВО ЗДРАВООХРАНЕНИЯ": "МЗ"}.items():
        if k in u:
            return a
    return n


def _qty_value(p):
    return int(p) if p == int(p) else p


def _key_norm(v):
    """Нормализация значения для ключа сравнения (числа -> единый вид)."""
    if v is None:
        return ""
    try:
        return repr(round(float(str(v).replace(",", ".")), 2))
    except (ValueError, TypeError):
        return str(v).strip()


def contract_to_row(data):
    cs = data.contract_number + (f" от {data.contract_date}" if data.contract_date else "")
    qv = (data.quantity_all_values
          if (data.quantity_mismatch and data.quantity_all_values)
          else _qty_value(data.quantity_packages))
    return [
        _owner_abbreviation(data.customer_short_name), "Основная заявка", "", cs, "",
        data.mnn, data.trade_name, data.dosage_form, data.unit, qv,
        data.manufacturer, data.unit_price, data.total_price,
        data.supplier_short_name, data.procurement_method,
    ]


def _hdr_style(ws, r):
    fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    font = Font(bold=True, size=10, color="FFFFFF")
    b = Border(*[Side(style="thin")] * 4)
    a = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for c in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=r, column=c)
        cell.font = font
        cell.fill = fill
        cell.border = b
        cell.alignment = a


def _data_style(ws, r):
    b = Border(*[Side(style="thin")] * 4)
    a = Alignment(vertical="center", wrap_text=True)
    for c in range(1, COL_Q + 1):
        cell = ws.cell(row=r, column=c)
        cell.border = b
        cell.alignment = a
        if c in (COL_UNIT_PRICE, COL_TOTAL, COL_P, COL_Q):
            cell.number_format = "#,##0.00"


def _build_book():
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Лист1"
    wb.create_sheet("Лист2")
    ws3 = wb.create_sheet(SHEET_NAME)
    ws4 = wb.create_sheet("Лист4")

    for c, h in enumerate(HEADERS, 1):
        ws3.cell(row=HEADER_ROW, column=c, value=h)
    _hdr_style(ws3, HEADER_ROW)
    ws3.cell(row=SUBHEADER_ROW, column=1, value=" Год фин-я 2026")
    ws3.cell(row=SUBHEADER_ROW, column=5, value="было")
    for c, w in enumerate(COLUMN_WIDTHS, 1):
        ws3.column_dimensions[get_column_letter(c)].width = w

    for i, name in enumerate(SUPPLIERS, 1):
        ws4.cell(row=i, column=1, value=name)
    wb.active = wb.sheetnames.index(SHEET_NAME)  # открывать сразу на Лист3
    return wb


def create_new_excel(file_path):
    wb = _build_book()
    path = file_path if str(file_path).lower().endswith(".xlsx") else str(file_path) + ".xlsx"
    wb.save(path)
    return path


def write_contracts_to_excel(file_path, contracts, sheet_name=None):
    path = Path(file_path)
    try:
        if path.exists() and path.suffix.lower() == ".xlsx":
            wb = load_workbook(str(path))
            if SHEET_NAME not in wb.sheetnames:
                base = _build_book()
                for sh in base.sheetnames:
                    if sh not in wb.sheetnames:
                        src = base[sh]
                        dst = wb.create_sheet(sh)
                        for row in src.iter_rows():
                            for cell in row:
                                if cell.value is not None:
                                    dst[cell.coordinate] = cell.value
            ws = wb[SHEET_NAME]
        else:
            wb = _build_book()
            ws = wb[SHEET_NAME]

        last_row = max(ws.max_row, SUBHEADER_ROW)

        # защита от дублей: ключ = №договора, МНН, торговое, лек.форма, кол-во, общая цена
        key_cols = (4, 6, 7, 8, 10, 13)
        existing = set()
        for r in range(DATA_START_ROW, ws.max_row + 1):
            existing.add(tuple(_key_norm(ws.cell(row=r, column=c).value) for c in key_cols))

        written = 0
        skipped = 0
        for contract in contracts:
            row_vals = contract_to_row(contract)
            key = tuple(_key_norm(row_vals[c - 1]) for c in key_cols)
            if key in existing:
                skipped += 1
                continue
            existing.add(key)
            row_num = last_row + 1 + written
            for c, value in enumerate(row_vals, 1):
                ws.cell(row=row_num, column=c, value=value)
            L = get_column_letter(COL_UNIT_PRICE)
            J = get_column_letter(COL_QTY)
            M = get_column_letter(COL_TOTAL)
            P = get_column_letter(COL_P)
            ws.cell(row=row_num, column=COL_P, value=f"={L}{row_num}*{J}{row_num}")
            ws.cell(row=row_num, column=COL_Q, value=f"={M}{row_num}-{P}{row_num}")
            _data_style(ws, row_num)

            if contract.quantity_mismatch:
                ws.cell(row=row_num, column=COL_QTY).fill = RED_FILL
            if contract.price_mismatch:
                ws.cell(row=row_num, column=COL_UNIT_PRICE).fill = RED_FILL
                ws.cell(row=row_num, column=COL_TOTAL).fill = RED_FILL
            if contract.dosage_form_empty:
                ws.cell(row=row_num, column=COL_DOSAGE).fill = ORANGE_FILL
            elif contract.dosage_form_mnn_only or contract.dosage_form_uncertain:
                ws.cell(row=row_num, column=COL_DOSAGE).fill = YELLOW_FILL
            written += 1

        if SHEET_NAME in wb.sheetnames:
            wb.active = wb.sheetnames.index(SHEET_NAME)  # открывать сразу на Лист3

        save_path = str(path)
        if not save_path.lower().endswith(".xlsx"):
            save_path = save_path.rsplit(".", 1)[0] + ".xlsx"
        wb.save(save_path)
        msg = f"Записано {written} строк(и)"
        if skipped:
            msg += f", пропущено дублей: {skipped}"
        msg += f" в {save_path}"
        return True, msg

    except PermissionError:
        return False, "Файл занят другой программой. Закройте Excel и повторите."
    except Exception as e:
        return False, f"Ошибка записи: {e}"


def get_sheet_names(file_path):
    try:
        p = Path(file_path)
        if p.suffix.lower() == ".xlsx":
            return load_workbook(str(p), read_only=True).sheetnames
    except Exception:
        pass
    return []

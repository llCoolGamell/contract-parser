"""
Обработчик Excel-файлов.
Запись извлечённых данных контрактов в Excel в формате .xlsx.
Поддерживает создание нового файла и добавление данных в существующий.
"""

from pathlib import Path
from typing import Optional

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from parser_engine import ContractData

HEADERS = [
    "Собственник товара",
    "тип",
    "код ЛС льготный",
    "№ договора",
    "№ аукциона",
    "Международное непатентованное название (МНН)",
    "Торговое наименование",
    "Лек. Форма (форма выпуска), дозировка",
    "Ед. изм.",
    "Общее кол-во, ед. изм.",
    "Производитель",
    "Цена ед. продукции, руб. в месте поставки",
    "Общая цена, руб.",
    "поставщик",
    "способ размещения",
]

COLUMN_WIDTHS = [15, 18, 15, 30, 25, 30, 25, 50, 10, 15, 45, 20, 20, 25, 25]

SHEET_NAME = "Контракты"

# Индексы столбцов (1-based)
COL_DOSAGE = 8       # Лек. Форма
COL_QTY = 10         # Общее кол-во
COL_UNIT_PRICE = 12  # Цена ед. продукции
COL_TOTAL = 13       # Общая цена


def _owner_abbreviation(full_name: str) -> str:
    """Сокращение наименования заказчика до аббревиатуры."""
    mapping = {
        "МИНЗДРАВ": "МЗ",
        "МИНИСТЕРСТВО ЗДРАВООХРАНЕНИЯ": "МЗ",
    }
    upper = full_name.upper()
    for key, abbr in mapping.items():
        if key in upper:
            return abbr
    return full_name


# Красный — расхождение количества из источников
RED_FILL = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
# Жёлтый — Лек.форма требует проверки (только МНН / нестандартная упаковка)
YELLOW_FILL = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
# Оранжевый — Лек.форма пустая (незарегистрированный препарат)
ORANGE_FILL = PatternFill(start_color="FFB347", end_color="FFB347", fill_type="solid")


def _qty_value(packages: float):
    """Целое, если значение целое, иначе дробное."""
    if packages == int(packages):
        return int(packages)
    return packages


def contract_to_row(data: ContractData) -> list:
    contract_str = data.contract_number
    if data.contract_date:
        contract_str += f" от {data.contract_date}"

    owner = _owner_abbreviation(data.customer_short_name)

    return [
        owner,
        "Основная заявка",
        "",                       # код ЛС льготный — заполняется вручную
        contract_str,
        "",                       # № аукциона — не заполняется
        data.mnn,
        data.trade_name,
        data.dosage_form,
        data.unit,
        _qty_value(data.quantity_packages),
        data.manufacturer,
        data.unit_price,
        data.total_price,
        data.supplier_short_name,
        data.procurement_method,
    ]


def _apply_header_style(ws, row_num: int) -> None:
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font_white = Font(bold=True, size=10, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col_idx in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=row_num, column=col_idx)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = center_align


def _apply_data_style(ws, row_num: int) -> None:
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    wrap_align = Alignment(vertical="center", wrap_text=True)

    for col_idx in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=row_num, column=col_idx)
        cell.border = thin_border
        cell.alignment = wrap_align
        if col_idx in (COL_UNIT_PRICE, COL_TOTAL):
            cell.number_format = "#,##0.00"


def create_new_excel(file_path: str) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    for col_idx, header in enumerate(HEADERS, 1):
        ws.cell(row=1, column=col_idx, value=header)

    _apply_header_style(ws, 1)

    for col_idx, width in enumerate(COLUMN_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    path = file_path
    if not path.endswith(".xlsx"):
        path += ".xlsx"

    wb.save(path)
    return path


def write_contracts_to_excel(
    file_path: str,
    contracts: list,
    sheet_name: Optional[str] = None,
) -> tuple:
    path = Path(file_path)
    target_sheet = sheet_name or SHEET_NAME

    try:
        if path.exists() and path.suffix == ".xlsx":
            wb = load_workbook(str(path))
            if target_sheet in wb.sheetnames:
                ws = wb[target_sheet]
            else:
                ws = wb.create_sheet(target_sheet)
                for col_idx, header in enumerate(HEADERS, 1):
                    ws.cell(row=1, column=col_idx, value=header)
                _apply_header_style(ws, 1)
                for col_idx, width in enumerate(COLUMN_WIDTHS, 1):
                    ws.column_dimensions[get_column_letter(col_idx)].width = width
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = target_sheet
            for col_idx, header in enumerate(HEADERS, 1):
                ws.cell(row=1, column=col_idx, value=header)
            _apply_header_style(ws, 1)
            for col_idx, width in enumerate(COLUMN_WIDTHS, 1):
                ws.column_dimensions[get_column_letter(col_idx)].width = width

        # Find last row with data
        last_row = ws.max_row
        if last_row == 1:
            first_cell = ws.cell(row=1, column=1).value
            if first_cell is None:
                last_row = 0

        written = 0
        for contract in contracts:
            row_data = contract_to_row(contract)
            row_num = last_row + 1 + written
            for col_idx, value in enumerate(row_data, 1):
                ws.cell(row=row_num, column=col_idx, value=value)
            _apply_data_style(ws, row_num)

            # Красный — расхождение количества
            if contract.quantity_mismatch:
                ws.cell(row=row_num, column=COL_QTY).fill = RED_FILL

            # Оранжевый — пустая форма; жёлтый — только МНН / нестандартная упаковка
            if contract.dosage_form_empty:
                ws.cell(row=row_num, column=COL_DOSAGE).fill = ORANGE_FILL
            elif contract.dosage_form_mnn_only or contract.dosage_form_uncertain:
                ws.cell(row=row_num, column=COL_DOSAGE).fill = YELLOW_FILL

            written += 1

        save_path = str(path)
        if not save_path.endswith(".xlsx"):
            save_path = save_path.rsplit(".", 1)[0] + ".xlsx"

        wb.save(save_path)
        msg = f"Записано {written} строк(и) в {save_path}"
        return True, msg

    except PermissionError:
        return False, (
            "Файл занят другой программой. Закройте Excel и повторите попытку."
        )
    except Exception as e:
        return False, f"Ошибка записи: {e}"


def get_sheet_names(file_path: str) -> list:
    try:
        path = Path(file_path)
        if path.suffix == ".xlsx":
            wb = load_workbook(str(path), read_only=True)
            return wb.sheetnames
        elif path.suffix == ".xls":
            import xlrd
            wb = xlrd.open_workbook(str(path))
            return wb.sheet_names()
    except Exception:
        pass
    return []

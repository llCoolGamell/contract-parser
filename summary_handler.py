# -*- coding: utf-8 -*-
"""
Сводный файл «Сводка по ГК»: один госконтракт = одна строка.
Столбцы: Поставщик | № и дата ГК | ГК на услугу | № извещения | Оплата по контракту.
Дозапись в существующий файл, без повторов (ключ — «№ и дата ГК»).
"""
import re
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEADERS = ["Поставщик", "№ и дата ГК", "ГК на услугу", "№ извещения", "Оплата по контракту"]
WIDTHS = [30, 28, 28, 22, 80]
SHEET = "Сводка"


def extract_payment_line(pdf_path):
    """Фраза «Оплата по контракту осуществляется … .» из PDF (до точки). Иначе 'не найдено'."""
    try:
        from pypdf import PdfReader
    except ImportError:
        return "не найдено"
    try:
        reader = PdfReader(str(pdf_path))
        full = ""
        for page in reader.pages:
            full += (page.extract_text() or "") + "\n"
        full = " ".join(full.split())
        m = re.search(r"(Оплата по контракту осуществляется[^.]*\.)", full)
        return m.group(1).strip() if m else "не найдено"
    except Exception:
        return "не найдено"


def _build_book():
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    font = Font(bold=True, size=10, color="FFFFFF")
    b = Border(*[Side(style="thin")] * 4)
    a = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = fill
        cell.font = font
        cell.border = b
        cell.alignment = a
    for c, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    return wb


def write_summary(path, rows):
    """rows: список dict {supplier, contract, service, notice, payment}."""
    p = Path(path)
    try:
        if p.exists() and p.suffix.lower() == ".xlsx":
            wb = load_workbook(str(p))
            ws = wb[SHEET] if SHEET in wb.sheetnames else wb.active
        else:
            wb = _build_book()
            ws = wb.active

        existing = set()
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=2).value
            if v:
                existing.add(str(v).strip())

        last = ws.max_row
        b = Border(*[Side(style="thin")] * 4)
        a = Alignment(vertical="top", wrap_text=True)
        written = 0
        skipped = 0
        for row in rows:
            key = str(row.get("contract", "")).strip()
            if key and key in existing:
                skipped += 1
                continue
            existing.add(key)
            rn = last + 1 + written
            vals = [row.get("supplier", ""), row.get("contract", ""),
                    row.get("service", ""), row.get("notice", ""), row.get("payment", "")]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=rn, column=c, value=v)
                cell.border = b
                cell.alignment = a
            written += 1

        sp = str(p)
        if not sp.lower().endswith(".xlsx"):
            sp = sp.rsplit(".", 1)[0] + ".xlsx"
        wb.save(sp)
        msg = f"Сводка по ГК: записано {written}"
        if skipped:
            msg += f", пропущено дублей: {skipped}"
        return True, msg, sp
    except PermissionError:
        return False, "Сводка по ГК: файл занят, закройте Excel.", str(p)
    except Exception as e:
        return False, f"Сводка по ГК: ошибка {e}", str(p)

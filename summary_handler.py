# -*- coding: utf-8 -*-
"""
Сводный файл «Сводка по ГК»: один госконтракт = одна строка.
Столбцы: Поставщик | № и дата ГК | ГК на услугу | Действует с | № извещения |
         Оплата по контракту | Сумма по ГК | Источник финансирования.
Источник определяется по тексту PDF. Дозапись без повторов (ключ — «№ и дата ГК»).
«Действует с» — дата начала исполнения ГК на услугу: раньше неё товар по контракту
принимать нельзя. Красится при каждой записи сводки: красный — дата ещё не
наступила либо в контракте нет ГК на услугу (нужен допник), зелёный — наступила.
"""
import re
from datetime import date
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from eis_downloader import NO_SERVICE, parse_date, start_state

HEADERS = ["Поставщик", "№ и дата ГК", "ГК на услугу", "Действует с", "№ извещения",
           "Оплата по контракту", "Сумма по ГК", "Источник финансирования"]
WIDTHS = [30, 28, 28, 16, 22, 80, 18, 28]
SHEET = "Сводка"
COL_SERVICE = 3
COL_START = 4
_START_FILL = {
    "future": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
    "noservice": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
    "ok": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
    "": PatternFill(fill_type=None),
}


def extract_pdf_text(pdf_path):
    """Весь текст PDF одной строкой (для поиска оплаты и источника). '' если не вышло."""
    try:
        from pypdf import PdfReader
    except ImportError:
        return ""
    try:
        reader = PdfReader(str(pdf_path))
        full = ""
        for page in reader.pages:
            full += (page.extract_text() or "") + "\n"
        return " ".join(full.split())
    except Exception:
        return ""


def payment_line_from_text(text):
    """Фраза «Оплата по контракту осуществляется … .» (до точки). Иначе 'не найдено'."""
    if not text:
        return "не найдено"
    m = re.search(r"(Оплата по контракту осуществляется[^.]*\.)", text)
    return m.group(1).strip() if m else "не найдено"


# Источники финансирования: (короткое имя, список регэкспов)
_FUNDING_RULES = [
    ("федеральный", [
        r"федеральн\w*\s+бюджет", r"бюджет\w*\s+росси",
        r"из\s+федеральн\w*\s+бюджет", r"субвенц\w*\s+из\s+федеральн",
    ]),
    ("региональный", [
        r"бюджет\w*\s+субъект", r"(областн|краев|республиканск)\w*\s+бюджет",
        r"бюджет\w*\s+[а-яё]+ск\w+\s+(област|кра|республик)",
        r"бюджет\w*\s+челябинск",
    ]),
    ("местный", [
        r"местн\w*\s+бюджет", r"бюджет\w*\s+муниципальн", r"муниципальн\w*\s+бюджет",
        r"бюджет\w*\s+городского\s+округа", r"бюджет\w*\s+поселени",
    ]),
    ("ОМС", [
        r"\bомс\b", r"обязательного\s+медицинского\s+страхования", r"фонд\w*\s+омс",
        r"территориальн\w*\s+фонд",
    ]),
    ("собственные средства", [
        r"собственн\w*\s+средств", r"приносящ\w*\s+доход",
    ]),
]

_LABELS = {
    "федеральный": "Федеральный бюджет",
    "региональный": "Региональный бюджет",
    "местный": "Местный бюджет",
    "ОМС": "ОМС",
    "собственные средства": "Собственные средства",
}


def classify_funding(text):
    """Определяет источник финансирования по тексту PDF (короткая категория)."""
    if not text:
        return "не определено"
    low = text.lower()
    found = []
    for key, pats in _FUNDING_RULES:
        if any(re.search(p, low) for p in pats):
            found.append(key)
    if not found:
        return "не определено"
    if len(found) == 1:
        return _LABELS[found[0]]
    return "Смешанное (" + " + ".join(found) + ")"


def _header_cell(ws, c, h):
    cell = ws.cell(row=1, column=c, value=h)
    cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    cell.font = Font(bold=True, size=10, color="FFFFFF")
    cell.border = Border(*[Side(style="thin")] * 4)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _set_widths(ws):
    for c, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[get_column_letter(c)].width = w


def _build_book():
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    for c, h in enumerate(HEADERS, 1):
        _header_cell(ws, c, h)
    _set_widths(ws)
    return wb


def _open_book(p):
    """Открывает сводку (или создаёт новую). В файл старого формата вставляет
    столбец «Действует с» после «ГК на услугу»."""
    if p.exists() and p.suffix.lower() == ".xlsx":
        wb = load_workbook(str(p))
        ws = wb[SHEET] if SHEET in wb.sheetnames else wb.active
        if str(ws.cell(row=1, column=COL_START).value or "").strip() != HEADERS[COL_START - 1]:
            ws.insert_cols(COL_START)
            _header_cell(ws, COL_START, HEADERS[COL_START - 1])
            b = Border(*[Side(style="thin")] * 4)
            a = Alignment(vertical="top", wrap_text=True)
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(row=r, column=COL_START)
                cell.border = b
                cell.alignment = a
            _set_widths(ws)
        return wb, ws
    wb = _build_book()
    return wb, wb.active


def _paint_start(ws, today=None):
    """Красит столбец «Действует с» относительно сегодняшнего дня."""
    today = today or date.today()
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(row=r, column=COL_START)
        cell.fill = _START_FILL[start_state(cell.value, today)]


def _save_path(p):
    sp = str(p)
    if not sp.lower().endswith(".xlsx"):
        sp = sp.rsplit(".", 1)[0] + ".xlsx"
    return sp


def write_summary(path, rows):
    """rows: список dict {supplier, contract, service, service_start, notice,
    payment, amount, funding}."""
    p = Path(path)
    try:
        wb, ws = _open_book(p)

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
                    row.get("service", ""), row.get("service_start", ""),
                    row.get("notice", ""),
                    row.get("payment", ""), row.get("amount", ""), row.get("funding", "")]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=rn, column=c, value=v)
                cell.border = b
                cell.alignment = a
            written += 1

        _paint_start(ws)
        sp = _save_path(p)
        wb.save(sp)
        msg = f"Сводка по ГК: записано {written}"
        if skipped:
            msg += f", пропущено дублей: {skipped}"
        return True, msg, sp
    except PermissionError:
        return False, "Сводка по ГК: файл занят, закройте Excel.", str(p)
    except Exception as e:
        return False, f"Сводка по ГК: ошибка {e}", str(p)


def backfill_service_start(path, resolve):
    """
    Дозаполняет «Действует с» в уже записанных строках сводки, где даты ещё нет.
    resolve(service) -> 'ДД.ММ.ГГГГ' или «не найдено»; исключение = ошибка связи,
    такая строка остаётся как была. Заодно перекрашивает весь столбец.
    """
    p = Path(path)
    if not p.exists():
        return False, "Сводка по ГК: файл не найден.", str(p)
    try:
        wb, ws = _open_book(p)
        filled = missing = errors = 0
        for r in range(2, ws.max_row + 1):
            if not ws.cell(row=r, column=2).value:
                continue
            cell = ws.cell(row=r, column=COL_START)
            if parse_date(cell.value):
                continue
            service = str(ws.cell(row=r, column=COL_SERVICE).value or "").strip()
            if not service:
                cell.value = NO_SERVICE
                missing += 1
                continue
            try:
                value = resolve(service)
            except Exception:
                errors += 1
                continue
            cell.value = value
            if parse_date(value):
                filled += 1
            else:
                missing += 1
        _paint_start(ws)
        sp = _save_path(p)
        wb.save(sp)
        msg = f"«Действует с»: дозаполнено {filled}, не найдено {missing}"
        if errors:
            msg += f", ошибок связи: {errors}"
        return True, msg, sp
    except PermissionError:
        return False, "Сводка по ГК: файл занят, закройте Excel.", str(p)
    except Exception as e:
        return False, f"Сводка по ГК: ошибка {e}", str(p)

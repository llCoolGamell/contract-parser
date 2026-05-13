"""
Парсер электронных контрактов с сайта zakupki.gov.ru (ЕИС).
Извлекает данные из HTML-файлов печатных форм контрактов.
"""

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from bs4 import BeautifulSoup


@dataclass
class ContractData:
    contract_number: str = ""
    contract_date: str = ""
    notice_number: str = ""
    procurement_method: str = ""
    customer_short_name: str = ""
    supplier_short_name: str = ""
    mnn: str = ""
    trade_name: str = ""
    dosage_form: str = ""
    dosage_form_mnn_only: bool = False
    unit: str = "упак."
    quantity_packages: float = 0.0
    quantity_mismatch: bool = False
    quantity_all_values: str = ""
    manufacturer: str = ""
    unit_price: float = 0.0
    total_price: float = 0.0
    source_file: str = ""
    errors: list = field(default_factory=list)


class ContractParser:
    """Парсер HTML-файлов контрактов с zakupki.gov.ru."""

    def parse_file(self, file_path: str) -> Optional[ContractData]:
        path = Path(file_path)
        if not path.exists():
            return None

        suffix = path.suffix.lower()
        if suffix in (".html", ".htm"):
            return self._parse_html(path)
        elif suffix == ".pdf":
            return self._parse_pdf(path)
        return None

    def _parse_html(self, path: Path) -> ContractData:
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()

        soup = BeautifulSoup(content, "lxml")
        texts = self._extract_texts(soup)
        data = ContractData(source_file=str(path))

        self._extract_contract_number(texts, data)
        self._extract_notice_number(texts, data)
        self._extract_procurement_method(texts, data)
        self._extract_customer_info(texts, data)
        self._extract_supplier_info(texts, data)
        self._extract_object_info(texts, data)
        self._extract_drug_details(texts, data)
        self._extract_contract_date(texts, data)

        return data

    def _extract_texts(self, soup: BeautifulSoup) -> list[str]:
        for tag in soup.find_all(["style", "script", "svg"]):
            tag.decompose()
        texts = []
        for text in soup.stripped_strings:
            cleaned = " ".join(text.split())
            if cleaned:
                texts.append(cleaned)
        return texts

    def _find_value_after(self, texts: list[str], label: str) -> str:
        for i, t in enumerate(texts):
            if label in t and i + 1 < len(texts):
                return texts[i + 1]
        return ""

    def _find_value_after_exact(self, texts: list[str], label: str) -> str:
        for i, t in enumerate(texts):
            if t.strip() == label and i + 1 < len(texts):
                return texts[i + 1]
        return ""

    def _extract_contract_number(self, texts: list[str], data: ContractData) -> None:
        data.contract_number = self._find_value_after_exact(
            texts, "Номер контракта"
        )

    def _extract_notice_number(self, texts: list[str], data: ContractData) -> None:
        for i, t in enumerate(texts):
            if "Номер извещения" in t and i + 1 < len(texts):
                data.notice_number = texts[i + 1]
                break

    def _extract_procurement_method(
        self, texts: list[str], data: ContractData
    ) -> None:
        for i, t in enumerate(texts):
            if "Способ определения поставщика" in t and i + 1 < len(texts):
                data.procurement_method = texts[i + 1].lower()
                break

    def _extract_customer_info(self, texts: list[str], data: ContractData) -> None:
        in_customer = False
        for i, t in enumerate(texts):
            if "2.1. Информация о заказчике" in t:
                in_customer = True
            if "2.2. Информация о поставщике" in t:
                break
            if in_customer and t == "Сокращенное наименование" and i + 1 < len(texts):
                data.customer_short_name = texts[i + 1]
                break

    def _extract_supplier_info(self, texts: list[str], data: ContractData) -> None:
        in_supplier = False
        for i, t in enumerate(texts):
            if "2.2. Информация о поставщике" in t:
                in_supplier = True
            if "3. Предмет контракта" in t:
                break
            if in_supplier and t == "Сокращенное наименование" and i + 1 < len(texts):
                data.supplier_short_name = texts[i + 1]
                break

    def _extract_object_info(self, texts: list[str], data: ContractData) -> None:
        for i, t in enumerate(texts):
            if t == "Итого:" and i + 1 < len(texts):
                price_str = texts[i + 1].replace("\u00a0", "").replace(" ", "").replace(",", ".")
                try:
                    data.total_price = float(price_str)
                except ValueError:
                    data.errors.append(f"Не удалось разобрать сумму: {texts[i + 1]}")
                break

    def _parse_qty_str(self, s: str) -> float:
        cleaned = re.sub(r"[^\d.,]", "", s.replace("\u00a0", "").replace(" ", ""))
        cleaned = cleaned.replace(",", ".")
        try:
            return float(cleaned)
        except ValueError:
            return 0.0

    def _extract_drug_details(self, texts: list[str], data: ContractData) -> None:
        # --- Extract MNN from object table ("Наименование объекта закупки" row) ---
        mnn_object = ""
        qty_object_str = ""
        for i, t in enumerate(texts):
            if "Наименование объекта закупки" in t:
                # Find data row: after header row (columns 1-9), then data row
                # Pattern: headers end with "9", then data: row_num, MNN_name, ...
                for j in range(i + 1, min(i + 30, len(texts))):
                    if texts[j] == "9" and j + 2 < len(texts):
                        # j+1 = row number "1", j+2 = MNN from object table
                        mnn_object = texts[j + 2]
                        break
                break

        # Extract "Количество (объем) и единица измерения" from object table
        for i, t in enumerate(texts):
            if "Наименование объекта закупки" in t:
                for j in range(i + 1, min(i + 30, len(texts))):
                    if texts[j] == "9" and j + 2 < len(texts):
                        # After 9, data row: row_num, MNN, type, code, qty_and_unit, price, vat, country, sum
                        idx = j + 1  # row_num
                        idx += 1  # MNN
                        idx += 1  # type (Товар)
                        idx += 1  # code
                        if idx < len(texts):
                            qty_object_str = texts[idx]
                        break
                break

        drug_section_start = None
        for i, t in enumerate(texts):
            if "Дополнительная информация о лекарственном препарате" in t:
                drug_section_start = i
                break

        if drug_section_start is None:
            data.errors.append("Раздел с информацией о лекарственном препарате не найден")
            if mnn_object:
                data.mnn = mnn_object
            return

        section_texts = texts[drug_section_start:]

        trade_name_raw = ""
        grls_form = ""
        holder = ""
        producer = ""
        qty_in_primary = 0
        qty_primary_per_consumer = 0
        qty_per_consumer = 0
        total_qty = 0
        qty_consumer_units = 0
        completeness = ""
        primary_pack_type = ""

        for i, t in enumerate(section_texts):
            if t == "Торговое наименование, номер РУ":
                for j in range(i + 1, min(i + 20, len(section_texts))):
                    if section_texts[j] == "6":
                        idx = j + 1
                        idx += 2  # skip "1" and "1"
                        if idx < len(section_texts):
                            trade_name_raw = section_texts[idx]
                        if idx + 2 < len(section_texts):
                            primary_pack_type = section_texts[idx + 2]
                        if idx + 3 < len(section_texts):
                            try:
                                qty_consumer_units = int(section_texts[idx + 3])
                            except ValueError:
                                pass
                        break

            if "МНН и форма выпуска в соответствии с ГРЛС" in t and i + 1 < len(section_texts):
                grls_form = section_texts[i + 1]

            if t == "Наименование держателя или владельца РУ" and i + 1 < len(section_texts):
                holder = section_texts[i + 1]

            if t == "Производитель" and i + 1 < len(section_texts):
                producer = section_texts[i + 1]

            if t == "Количество лекарственных форм в первичной упаковке" and i + 1 < len(section_texts):
                try:
                    qty_in_primary = int(section_texts[i + 1])
                except ValueError:
                    pass

            if t == "Количество первичных упаковок в потребительской упаковке" and i + 1 < len(section_texts):
                try:
                    qty_primary_per_consumer = int(section_texts[i + 1])
                except ValueError:
                    pass

            if t == "Количество потребительских единиц в потребительской упаковке" and i + 1 < len(section_texts):
                try:
                    qty_per_consumer = int(section_texts[i + 1])
                except ValueError:
                    pass

            if t == "Общее количество" and i + 1 < len(section_texts):
                try:
                    total_qty = int(section_texts[i + 1])
                except ValueError:
                    pass

            if t == "Комплектность потребительской упаковки" and i + 1 < len(section_texts):
                completeness = section_texts[i + 1]

            if "4. Условия контракта" in t:
                break

        # --- 1) МНН: from both ГРЛС and object table ---
        mnn_grls = ""
        if grls_form:
            parts = grls_form.split(":", 1)
            if len(parts) == 2:
                mnn_grls = parts[0].strip()

        if mnn_grls and mnn_object:
            if mnn_grls.upper() == mnn_object.upper():
                data.mnn = mnn_grls
            else:
                data.mnn = f"{mnn_grls}, {mnn_object}"
        elif mnn_grls:
            data.mnn = mnn_grls
        elif mnn_object:
            data.mnn = mnn_object

        # --- 2) Количество: compare from 3 sources ---
        qty_from_object = self._parse_qty_str(qty_object_str) if qty_object_str else 0.0
        qty_from_consumer = float(qty_consumer_units) if qty_consumer_units else 0.0
        qty_from_total = float(total_qty) if total_qty else 0.0

        qty_values = {}
        if qty_from_object > 0:
            qty_values["Объём закупки"] = qty_from_object
        if qty_from_consumer > 0:
            qty_values["Потреб. ед."] = qty_from_consumer
        if qty_from_total > 0:
            qty_values["Общее кол-во"] = qty_from_total

        unique_vals = set(qty_values.values())
        if len(unique_vals) > 1:
            data.quantity_mismatch = True
            parts = [f"{k}: {int(v) if v == int(v) else v}" for k, v in qty_values.items()]
            data.quantity_all_values = ", ".join(parts)

        # Calculate packages (use total_qty / qty_per_consumer as before)
        if qty_per_consumer > 0 and total_qty > 0:
            data.quantity_packages = total_qty / qty_per_consumer
        elif total_qty > 0:
            data.quantity_packages = total_qty

        # --- 3) Лек. форма: from ГРЛС, strip МНН (part before colon) ---
        if grls_form:
            parts = grls_form.split(":", 1)
            if len(parts) == 2:
                dosage_text = parts[1].strip()
            else:
                dosage_text = grls_form.strip()
            if not dosage_text:
                dosage_text = grls_form
                data.dosage_form_mnn_only = True
            if qty_per_consumer > 0:
                dosage_text += f" \u2116{qty_per_consumer}"
            if completeness and completeness != "~":
                dosage_text += f", {completeness}"
            data.dosage_form = dosage_text
        elif mnn_object:
            data.dosage_form = mnn_object
            data.dosage_form_mnn_only = True

        # Extract trade name (part before comma or registration number)
        if trade_name_raw:
            match = re.match(r"^(.+?),\s*ЛП", trade_name_raw)
            if match:
                data.trade_name = match.group(1).strip()
            else:
                data.trade_name = trade_name_raw.split(",")[0].strip()

        # Calculate unit price
        if data.quantity_packages > 0 and data.total_price > 0:
            data.unit_price = round(data.total_price / data.quantity_packages, 2)

        # Build manufacturer string
        producer_clean = re.sub(r"\s*\(\d+\)\s*$", "", producer).strip()
        if holder and producer_clean:
            if producer_clean.startswith(holder):
                data.manufacturer = producer_clean
            else:
                data.manufacturer = f"{holder}/ {producer_clean}"
        elif producer_clean:
            data.manufacturer = producer_clean
        elif holder:
            data.manufacturer = holder

    def _extract_contract_date(self, texts: list[str], data: ContractData) -> None:
        start_date = ""
        for i, t in enumerate(texts):
            if "Дата начала исполнения контракта" in t and i + 1 < len(texts):
                val = texts[i + 1].strip()
                date_match = re.match(r"\d{2}\.\d{2}\.\d{4}", val)
                if date_match:
                    start_date = date_match.group(0)
                break

        if start_date:
            data.contract_date = start_date
            return

        signing_dates = []
        for i, t in enumerate(texts):
            if t == "Дата и время подписания:" and i + 1 < len(texts):
                date_match = re.match(r"\d{2}\.\d{2}\.\d{4}", texts[i + 1])
                if date_match:
                    signing_dates.append(date_match.group(0))

        if signing_dates:
            data.contract_date = signing_dates[-1]

    def _parse_pdf(self, path: Path) -> ContractData:
        data = ContractData(source_file=str(path))
        try:
            import pdfplumber

            with pdfplumber.open(path) as pdf:
                full_text = ""
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        full_text += text + "\n"

            if not full_text.strip():
                data.errors.append(
                    "PDF не содержит текстового слоя. "
                    "Требуется OCR (Tesseract)."
                )
                return data

            texts = [
                " ".join(line.split())
                for line in full_text.split("\n")
                if line.strip()
            ]
            self._extract_contract_number(texts, data)
            self._extract_notice_number(texts, data)
            self._extract_procurement_method(texts, data)
            self._extract_customer_info(texts, data)
            self._extract_supplier_info(texts, data)
            self._extract_object_info(texts, data)
            self._extract_drug_details(texts, data)
            self._extract_contract_date(texts, data)

        except ImportError:
            data.errors.append(
                "Модуль pdfplumber не установлен. "
                "Установите: pip install pdfplumber"
            )
        except Exception as e:
            data.errors.append(f"Ошибка чтения PDF: {e}")

        return data

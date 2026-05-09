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
    unit: str = "упак."
    quantity_packages: float = 0.0
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
        found_short = False
        for i, t in enumerate(texts):
            if "2.1. Информация о заказчике" in t:
                in_customer = True
            if "2.2. Информация о поставщике" in t:
                break
            if in_customer and t == "Сокращенное наименование" and i + 1 < len(texts):
                data.customer_short_name = texts[i + 1]
                found_short = True
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
                price_str = texts[i + 1].replace(" ", "").replace(",", ".")
                try:
                    data.total_price = float(price_str)
                except ValueError:
                    data.errors.append(f"Не удалось разобрать сумму: {texts[i + 1]}")
                break

    def _extract_drug_details(self, texts: list[str], data: ContractData) -> None:
        drug_section_start = None
        for i, t in enumerate(texts):
            if "Дополнительная информация о лекарственном препарате" in t:
                drug_section_start = i
                break

        if drug_section_start is None:
            data.errors.append("Раздел с информацией о лекарственном препарате не найден")
            return

        section_texts = texts[drug_section_start:]

        # Trade name (before comma)
        trade_name_raw = ""
        grls_form = ""
        holder = ""
        producer = ""
        qty_in_primary = 0
        qty_primary_per_consumer = 0
        qty_per_consumer = 0
        total_qty = 0
        completeness = ""
        primary_pack_type = ""

        for i, t in enumerate(section_texts):
            if t == "Торговое наименование, номер РУ":
                # Data rows: header cols (1-6), then data cols (1, 1, name, form, pack, qty)
                # Find the data after the "6" marker
                # Pattern: after headers (1,2,3,4,5,6), then data (1, 1, trade_name, form, pack_type, qty)
                for j in range(i + 1, min(i + 20, len(section_texts))):
                    if section_texts[j] == "6":
                        # Next items: row_num, obj_num, trade_name, dosage, pack_type, qty
                        idx = j + 1
                        # Skip row number and obj number
                        idx += 2  # skip "1" and "1"
                        if idx < len(section_texts):
                            trade_name_raw = section_texts[idx]
                        if idx + 1 < len(section_texts):
                            # dosage form from drug details table
                            pass  # We'll use ГРЛС form instead
                        if idx + 2 < len(section_texts):
                            primary_pack_type = section_texts[idx + 2]
                        if idx + 3 < len(section_texts):
                            try:
                                total_qty = int(section_texts[idx + 3])
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

        # Extract trade name (part before comma or registration number)
        if trade_name_raw:
            match = re.match(r"^(.+?),\s*ЛП", trade_name_raw)
            if match:
                data.trade_name = match.group(1).strip()
            else:
                data.trade_name = trade_name_raw.split(",")[0].strip()

        # Extract MNN and dosage form from ГРЛС
        if grls_form:
            parts = grls_form.split(":", 1)
            if len(parts) == 2:
                data.mnn = parts[0].strip()
                base_form = parts[1].strip()
            else:
                base_form = grls_form
                data.mnn = ""

            # Construct dosage form string
            is_tablet = any(
                kw in base_form.upper()
                for kw in ["ТАБЛЕТ", "КАПСУЛ", "ДРАЖЕ"]
            )
            if is_tablet and qty_per_consumer > 0:
                data.dosage_form = f"{base_form} № {qty_per_consumer}"
            elif qty_in_primary > 0:
                data.dosage_form = f"{base_form} {qty_in_primary} мл"
                if primary_pack_type and primary_pack_type != "~":
                    if completeness and completeness != "~":
                        comp_lower = completeness.lower()
                        data.dosage_form += (
                            f" {primary_pack_type.lower()} с {comp_lower}"
                        )
                data.dosage_form += f" №{qty_primary_per_consumer}"
            else:
                data.dosage_form = base_form

        # Calculate packages
        if qty_per_consumer > 0 and total_qty > 0:
            data.quantity_packages = total_qty / qty_per_consumer
        elif total_qty > 0:
            data.quantity_packages = total_qty

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
        # Try to get start date
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

        # Fallback: last signing date (customer's)
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

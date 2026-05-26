"""
Парсер электронных контрактов с сайта zakupki.gov.ru (ЕИС).
Извлекает данные из HTML-файлов печатных форм контрактов.
Поддерживает контракты с несколькими позициями лекарственных препаратов.
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

    def parse_file(self, file_path: str) -> list[ContractData]:
        path = Path(file_path)
        if not path.exists():
            return []

        suffix = path.suffix.lower()
        if suffix in (".html", ".htm"):
            return self._parse_html(path)
        elif suffix == ".pdf":
            return self._parse_pdf(path)
        return []

    def _parse_html(self, path: Path) -> list[ContractData]:
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()

        soup = BeautifulSoup(content, "lxml")
        texts = self._extract_texts(soup)

        base = ContractData(source_file=str(path))
        self._extract_contract_number(texts, base)
        self._extract_notice_number(texts, base)
        self._extract_procurement_method(texts, base)
        self._extract_customer_info(texts, base)
        self._extract_supplier_info(texts, base)
        self._extract_contract_date(texts, base)

        object_rows = self._parse_object_table(texts)
        drug_entries = self._parse_drug_entries(texts)

        if drug_entries:
            return self._build_contracts(base, object_rows, drug_entries)

        self._extract_object_info(texts, base)
        self._extract_drug_details(texts, base)
        return [base] if base.contract_number else []

    # ------------------------------------------------------------------
    # Multi-row: object table parsing
    # ------------------------------------------------------------------

    def _parse_object_table(self, texts: list[str]) -> dict[int, dict]:
        """Parse the object table (3.1) into {row_num: row_data}."""
        result = {}

        header_pos = None
        for i, t in enumerate(texts):
            if "Наименование объекта закупки" in t:
                header_pos = i
                break

        if header_pos is None:
            return result

        data_start = None
        for j in range(header_pos + 9, min(header_pos + 40, len(texts))):
            if (
                texts[j] == "9"
                and j >= 8
                and texts[j - 1] == "8"
                and texts[j - 2] == "7"
                and texts[j - 3] == "6"
                and texts[j - 4] == "5"
                and texts[j - 5] == "4"
                and texts[j - 6] == "3"
                and texts[j - 7] == "2"
                and texts[j - 8] == "1"
            ):
                data_start = j + 1
                break

        if data_start is None:
            return result

        idx = data_start
        while idx + 8 < len(texts):
            if texts[idx] == "Итого:" or "Дополнительная информация" in texts[idx]:
                break

            try:
                row_num = int(texts[idx])
            except ValueError:
                break

            name = texts[idx + 1]
            qty_and_unit = texts[idx + 4]
            unit_price_str = texts[idx + 5]
            total_sum_str = texts[idx + 8]

            result[row_num] = {
                "name": name,
                "qty": self._parse_qty_str(qty_and_unit),
                "unit_price": self._parse_qty_str(unit_price_str),
                "sum": self._parse_qty_str(total_sum_str),
            }
            idx += 9

        return result

    # ------------------------------------------------------------------
    # Multi-row: drug detail entries parsing
    # ------------------------------------------------------------------

    def _parse_drug_entries(self, texts: list[str]) -> list[dict]:
        """Parse all drug detail entries from 'Дополнительная информация о лекарственном препарате'."""
        entries = []

        section_start = None
        for i, t in enumerate(texts):
            if "Дополнительная информация о лекарственном препарате" in t:
                section_start = i
                break

        if section_start is None:
            return entries

        data_start = None
        for j in range(section_start + 6, min(section_start + 30, len(texts))):
            if (
                texts[j] == "6"
                and j >= 5
                and texts[j - 1] == "5"
                and texts[j - 2] == "4"
                and texts[j - 3] == "3"
                and texts[j - 4] == "2"
                and texts[j - 5] == "1"
            ):
                data_start = j + 1
                break

        if data_start is None:
            return entries

        section_end = len(texts)
        for i in range(data_start, len(texts)):
            t = texts[i]
            if "4. Условия контракта" in t:
                section_end = i
                break
            if "Дополнительная информация об адресе" in t:
                section_end = i
                break

        idx = data_start
        while idx + 5 < section_end:
            try:
                row_num = int(texts[idx])
            except ValueError:
                idx += 1
                continue

            try:
                obj_num = int(texts[idx + 1])
            except ValueError:
                idx += 1
                continue

            entry = {
                "row_num": row_num,
                "obj_num": obj_num,
                "trade_name_raw": texts[idx + 2],
                "dosage_form_raw": texts[idx + 3],
                "primary_pack_type": texts[idx + 4],
                "qty_consumer": 0,
                "grls_form": "",
                "holder": "",
                "producer": "",
                "qty_in_primary": 0,
                "qty_primary_per_consumer": 0,
                "qty_per_consumer": 0,
                "total_qty": 0,
                "completeness": "",
            }

            try:
                entry["qty_consumer"] = int(texts[idx + 5])
            except ValueError:
                pass

            idx += 6

            while idx < section_end:
                t = texts[idx]

                if t == "МНН и форма выпуска в соответствии с ГРЛС" and idx + 1 < section_end:
                    entry["grls_form"] = texts[idx + 1]
                    idx += 2
                elif t == "Наименование держателя или владельца РУ" and idx + 1 < section_end:
                    entry["holder"] = texts[idx + 1]
                    idx += 2
                elif t == "Производитель" and idx + 1 < section_end:
                    entry["producer"] = texts[idx + 1]
                    idx += 2
                elif t == "Количество лекарственных форм в первичной упаковке" and idx + 1 < section_end:
                    try:
                        entry["qty_in_primary"] = int(texts[idx + 1])
                    except ValueError:
                        pass
                    idx += 2
                elif t == "Количество первичных упаковок в потребительской упаковке" and idx + 1 < section_end:
                    try:
                        entry["qty_primary_per_consumer"] = int(texts[idx + 1])
                    except ValueError:
                        pass
                    idx += 2
                elif t == "Количество потребительских единиц в потребительской упаковке" and idx + 1 < section_end:
                    try:
                        entry["qty_per_consumer"] = int(texts[idx + 1])
                    except ValueError:
                        pass
                    idx += 2
                elif t == "Общее количество" and idx + 1 < section_end:
                    try:
                        entry["total_qty"] = int(texts[idx + 1])
                    except ValueError:
                        pass
                    idx += 2
                elif t == "Комплектность потребительской упаковки" and idx + 1 < section_end:
                    val = texts[idx + 1].strip()
                    if val and val != "~":
                        entry["completeness"] = val
                    idx += 2
                elif t == "Признак включения в ЖНВЛП" and idx + 1 < section_end:
                    idx += 2
                    break
                else:
                    idx += 1

            entries.append(entry)

        return entries

    # ------------------------------------------------------------------
    # Multi-row: build ContractData list from parsed entries
    # ------------------------------------------------------------------

    def _build_contracts(
        self,
        base: ContractData,
        object_rows: dict[int, dict],
        drug_entries: list[dict],
    ) -> list[ContractData]:
        results = []

        for entry in drug_entries:
            data = ContractData(
                source_file=base.source_file,
                contract_number=base.contract_number,
                contract_date=base.contract_date,
                notice_number=base.notice_number,
                procurement_method=base.procurement_method,
                customer_short_name=base.customer_short_name,
                supplier_short_name=base.supplier_short_name,
            )

            grls_form = entry["grls_form"]
            obj_num = entry["obj_num"]
            completeness = entry.get("completeness", "")

            # --- MNN ---
            mnn_grls = ""
            if grls_form:
                parts = grls_form.split(":", 1)
                if len(parts) == 2:
                    mnn_grls = parts[0].strip()

            mnn_object = ""
            if obj_num in object_rows:
                mnn_object = object_rows[obj_num]["name"]

            if mnn_grls and mnn_object:
                if (
                    mnn_grls.upper() == mnn_object.upper()
                    or mnn_grls.upper() in mnn_object.upper()
                    or mnn_object.upper() in mnn_grls.upper()
                ):
                    data.mnn = mnn_grls
                else:
                    data.mnn = f"{mnn_grls}, {mnn_object}"
            elif mnn_grls:
                data.mnn = mnn_grls
            elif mnn_object:
                data.mnn = mnn_object

            # --- Dosage form: <форма> [+ qty мл] [+ комплектность] + №X ---
            qty_per_consumer = entry["qty_per_consumer"]
            if grls_form:
                parts = grls_form.split(":", 1)
                if len(parts) == 2:
                    dosage_text = parts[1].strip()
                else:
                    dosage_text = grls_form.strip()
                if not dosage_text:
                    dosage_text = grls_form
                    data.dosage_form_mnn_only = True

                body = dosage_text
                num_part = ""
                if qty_per_consumer > 0:
                    if "мл" in dosage_text.lower():
                        body = f"{dosage_text} {qty_per_consumer} мл"
                        num_part = "№1"
                    else:
                        num_part = f"№{qty_per_consumer}"

                if completeness:
                    body = f"{body}, {completeness}"

                if num_part:
                    data.dosage_form = f"{body} {num_part}"
                else:
                    data.dosage_form = body
            elif mnn_object:
                data.dosage_form = mnn_object
                data.dosage_form_mnn_only = True

            # --- Trade name ---
            trade_name_raw = entry["trade_name_raw"]
            if trade_name_raw:
                match = re.match(r"^(.+?),\s*ЛП", trade_name_raw)
                if match:
                    data.trade_name = match.group(1).strip()
                else:
                    data.trade_name = trade_name_raw.split(",")[0].strip()

            # --- Manufacturer: holder / producer ---
            producer = entry["producer"]
            holder = entry["holder"]
            producer_clean = re.sub(r"\s*\(\d+\)\s*$", "", producer).strip()
            if holder and producer_clean:
                if producer_clean.upper().startswith(holder.upper()):
                    data.manufacturer = producer_clean
                else:
                    data.manufacturer = f"{holder}/ {producer_clean}"
            elif producer_clean:
                data.manufacturer = producer_clean
            elif holder:
                data.manufacturer = holder

            # --- Quantities ---
            total_qty = entry["total_qty"]
            qty_consumer = entry["qty_consumer"]

            if qty_per_consumer > 0 and total_qty > 0:
                data.quantity_packages = total_qty / qty_per_consumer
            elif total_qty > 0:
                data.quantity_packages = float(total_qty)

            # Compare quantities from 3 sources
            qty_from_object = 0.0
            if obj_num in object_rows:
                qty_from_object = object_rows[obj_num]["qty"]
            qty_from_consumer = float(qty_consumer) if qty_consumer else 0.0
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
                parts = [
                    f"{k}: {int(v) if v == int(v) else v}"
                    for k, v in qty_values.items()
                ]
                data.quantity_all_values = ", ".join(parts)

            # --- Price from object table ---
            if obj_num in object_rows:
                data.total_price = object_rows[obj_num]["sum"]

            if data.quantity_packages > 0 and data.total_price > 0:
                data.unit_price = round(data.total_price / data.quantity_packages, 2)

            results.append(data)

        return results

    # ------------------------------------------------------------------
    # Common helpers
    # ------------------------------------------------------------------

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
                price_str = texts[i + 1].replace(" ", "").replace("\u00a0", "").replace(",", ".")
                try:
                    data.total_price = float(price_str)
                except ValueError:
                    data.errors.append(f"Не удалось разобрать сумму: {texts[i + 1]}")
                break

    def _parse_qty_str(self, s: str) -> float:
        """
        Извлекает первое число из начала строки.
        Примеры:
        - "15120 Кубический сантиметр; миллилитр (СМ3; МЛ)" -> 15120.0
        - "15 483 859.20" -> 15483859.20
        - "1 024,06476190476" -> 1024.06476190476
        - "2 799 960.30" -> 2799960.30
        """
        if not s:
            return 0.0
        # Нормализуем неразрывные пробелы
        normalized = s.replace("\u00a0", " ").strip()
        # Берём подстроку с начала: цифры, пробелы, точка/запятая
        match = re.match(r"^\s*([\d][\d\s]*(?:[.,]\d+)?)", normalized)
        if not match:
            return 0.0
        num_str = match.group(1).replace(" ", "").replace(",", ".")
        try:
            return float(num_str)
        except ValueError:
            return 0.0

    def _extract_drug_details(self, texts: list[str], data: ContractData) -> None:
        mnn_object = ""
        qty_object_str = ""
        for i, t in enumerate(texts):
            if "Наименование объекта закупки" in t:
                for j in range(i + 1, min(i + 30, len(texts))):
                    if texts[j] == "9" and j + 2 < len(texts):
                        mnn_object = texts[j + 2]
                        break
                break

        for i, t in enumerate(texts):
            if "Наименование объекта закупки" in t:
                for j in range(i + 1, min(i + 30, len(texts))):
                    if texts[j] == "9" and j + 2 < len(texts):
                        idx = j + 1
                        idx += 1
                        idx += 1
                        idx += 1
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
        qty_per_consumer = 0
        total_qty = 0
        qty_consumer_units = 0
        primary_pack_type = ""
        completeness = ""

        for i, t in enumerate(section_texts):
            if t == "Торговое наименование, номер РУ":
                for j in range(i + 1, min(i + 20, len(section_texts))):
                    if section_texts[j] == "6":
                        idx = j + 1
                        idx += 2
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

        if qty_per_consumer > 0 and total_qty > 0:
            data.quantity_packages = total_qty / qty_per_consumer
        elif total_qty > 0:
            data.quantity_packages = total_qty

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
                if "мл" in dosage_text.lower():
                    dosage_text += f" {qty_per_consumer} мл №1"
                else:
                    dosage_text += f" №{qty_per_consumer}"
            if completeness and completeness != "~":
                dosage_text += f", {completeness}"
            data.dosage_form = dosage_text
        elif mnn_object:
            data.dosage_form = mnn_object
            data.dosage_form_mnn_only = True

        if trade_name_raw:
            match = re.match(r"^(.+?),\s*ЛП", trade_name_raw)
            if match:
                data.trade_name = match.group(1).strip()
            else:
                data.trade_name = trade_name_raw.split(",")[0].strip()

        if data.quantity_packages > 0 and data.total_price > 0:
            data.unit_price = round(data.total_price / data.quantity_packages, 2)

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

    def _parse_pdf(self, path: Path) -> list[ContractData]:
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
                return [data]

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

        return [data]

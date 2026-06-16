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
    dosage_form_empty: bool = False
    dosage_form_uncertain: bool = False  # нестандартная упаковка -> пометить жёлтым
    unit: str = "упак."
    quantity_packages: float = 0.0
    quantity_mismatch: bool = False
    quantity_all_values: str = ""
    manufacturer: str = ""
    unit_price: float = 0.0
    total_price: float = 0.0
    source_file: str = ""
    errors: list = field(default_factory=list)

    @property
    def is_empty(self) -> bool:
        """Контракт без полезных данных (нет препарата) — выводить не нужно."""
        return (
            not self.mnn
            and not self.trade_name
            and self.quantity_packages == 0
            and self.total_price == 0
        )


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

        # Нет раздела о препарате — данных нет, пишем ошибку в лог
        base.errors.append(
            "Нет данных о лекарственном препарате (раздел не найден) — строка пропущена"
        )
        base.dosage_form_empty = True
        return [base]

    # ------------------------------------------------------------------
    # Числовые помощники
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_num(s: str) -> float:
        """Парсит число (целое или дробное) из текста: '0.4' -> 0.4, '1,91' -> 1.91, '8' -> 8."""
        if not s:
            return 0.0
        normalized = s.replace(" ", " ").strip()
        match = re.match(r"^\s*(\d[\d\s]*(?:[.,]\d+)?)", normalized)
        if not match:
            return 0.0
        num_str = match.group(1).replace(" ", "").replace(",", ".")
        try:
            return float(num_str)
        except ValueError:
            return 0.0

    @staticmethod
    def _format_num(v: float) -> str:
        """Число для текста Лек.формы: 120 -> '120', 0.4 -> '0,4', 2.5 -> '2,5'."""
        if v == int(v):
            return str(int(v))
        s = ("%f" % v).rstrip("0").rstrip(".")
        return s.replace(".", ",")

    def _pack_word(self, pack_type: str, count: float) -> str:
        """
        Слово упаковки для Лек.формы.
        Добавляется ТОЛЬКО для шприцев, шприц-ручек и ингаляторов.
        Для флаконов, баллонов, ампул и т.п. — пусто (пометится жёлтым на проверку).
        """
        pt = (pack_type or "").strip().upper()
        n = count if count and count > 0 else 1
        if "ШПРИЦ-РУЧК" in pt:
            return "шприц-ручка" if n <= 1 else "шприц-ручки"
        if "ШПРИЦ" in pt:
            return "шприц" if n <= 1 else "шприцы"
        if "ИНГАЛЯТОР" in pt:
            return "ИНГАЛЯТОР"  # как в эталоне: верхний регистр, единственное число
        return ""

    def _build_dosage_form(
        self,
        grls_form: str,
        mnn_object: str,
        qty_in_primary: float,
        qty_primary_per_consumer: float,
        pack_type: str,
    ) -> tuple[str, bool, bool, bool]:
        """
        Возвращает (текст, mnn_only, empty, uncertain).
        Формат: <форма из ГРЛС> + <число> + <ед-ца мл/доз> + [слово упаковки] + №<кол-во перв. упак.>
        Единица: 'мл' для растворов/капель, 'доз' для аэрозолей/ингаляций,
                 для лиофилизата (МЕ и т.п.) число/единица не добавляются, только №N.
        """
        # База: форма из ГРЛС (часть после двоеточия)
        if grls_form:
            parts = grls_form.split(":", 1)
            if len(parts) == 2 and parts[1].strip():
                form_text = parts[1].strip()
            else:
                # после двоеточия пусто — есть только МНН
                return grls_form.strip(), True, False, True
        elif mnn_object:
            return mnn_object, True, False, True
        else:
            return "ПУСТО", False, True, False

        low = form_text.lower()
        if "мл" in low:
            unit = "мл"
        elif "доз" in low:
            unit = "доз"
        else:
            unit = ""

        word = self._pack_word(pack_type, qty_primary_per_consumer)
        n_str = self._format_num(
            qty_primary_per_consumer if qty_primary_per_consumer > 0 else 1
        )

        if unit and qty_in_primary > 0:
            suffix = f"{self._format_num(qty_in_primary)} {unit}"
            if word:
                suffix += f" {word}"
            suffix += f" №{n_str}"
        else:
            # лиофилизат / нераспознанная единица — только №N (+ слово, если есть)
            suffix = f"{word} №{n_str}".strip() if word else f"№{n_str}"

        text = f"{form_text} {suffix}"
        # жёлтый, если слово упаковки не добавлено (флакон/баллон/прочее) — на ручную проверку
        uncertain = not bool(word)
        return text, False, False, uncertain

    # ------------------------------------------------------------------
    # Multi-row: object table parsing
    # ------------------------------------------------------------------

    def _parse_object_table(self, texts: list[str]) -> dict:
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
                "qty": self._parse_num(qty_and_unit),
                "unit_price": self._parse_num(unit_price_str),
                "sum": self._parse_num(total_sum_str),
            }
            idx += 9

        return result

    # ------------------------------------------------------------------
    # Multi-row: drug detail entries parsing
    # ------------------------------------------------------------------

    def _parse_drug_entries(self, texts: list[str]) -> list:
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
                "qty_consumer": 0.0,
                "grls_form": "",
                "holder": "",
                "producer": "",
                "qty_in_primary": 0.0,
                "qty_primary_per_consumer": 0.0,
                "qty_per_consumer": 0.0,
                "total_qty": 0.0,
                "completeness": "",
            }

            entry["qty_consumer"] = self._parse_num(texts[idx + 5])

            idx += 6

            while idx < section_end:
                t = texts[idx]

                if t == "МНН и форма выпуска в соответствии с ГРЛС" and idx + 1 < section_end:
                    next_val = texts[idx + 1]
                    if next_val in (
                        "Наименование держателя или владельца РУ",
                        "Производитель",
                        "Количество лекарственных форм в первичной упаковке",
                        "Признак включения в ЖНВЛП",
                    ):
                        entry["grls_form"] = ""
                        idx += 1
                    else:
                        entry["grls_form"] = next_val
                        idx += 2
                elif t == "Наименование держателя или владельца РУ" and idx + 1 < section_end:
                    entry["holder"] = texts[idx + 1]
                    idx += 2
                elif t == "Производитель" and idx + 1 < section_end:
                    entry["producer"] = texts[idx + 1]
                    idx += 2
                elif t == "Количество лекарственных форм в первичной упаковке" and idx + 1 < section_end:
                    entry["qty_in_primary"] = self._parse_num(texts[idx + 1])
                    idx += 2
                elif t == "Количество первичных упаковок в потребительской упаковке" and idx + 1 < section_end:
                    entry["qty_primary_per_consumer"] = self._parse_num(texts[idx + 1])
                    idx += 2
                elif t == "Количество потребительских единиц в потребительской упаковке" and idx + 1 < section_end:
                    entry["qty_per_consumer"] = self._parse_num(texts[idx + 1])
                    idx += 2
                elif t == "Общее количество" and idx + 1 < section_end:
                    entry["total_qty"] = self._parse_num(texts[idx + 1])
                    idx += 2
                elif t == "Комплектность потребительской упаковки" and idx + 1 < section_end:
                    entry["completeness"] = texts[idx + 1]
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
        object_rows: dict,
        drug_entries: list,
    ) -> list:
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

            # --- Dosage form ---
            (
                data.dosage_form,
                data.dosage_form_mnn_only,
                data.dosage_form_empty,
                data.dosage_form_uncertain,
            ) = self._build_dosage_form(
                grls_form,
                mnn_object,
                entry["qty_in_primary"],
                entry["qty_primary_per_consumer"],
                entry["primary_pack_type"],
            )

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

            # --- Quantities (упаковки = Общее количество / потреб. ед. в упаковке) ---
            total_qty = entry["total_qty"]
            qty_per_consumer = entry["qty_per_consumer"]
            qty_consumer = entry["qty_consumer"]

            if qty_per_consumer > 0 and total_qty > 0:
                data.quantity_packages = total_qty / qty_per_consumer
            elif total_qty > 0:
                data.quantity_packages = float(total_qty)

            # Округление до целого, если результат практически целый
            if abs(data.quantity_packages - round(data.quantity_packages)) < 1e-6:
                data.quantity_packages = float(round(data.quantity_packages))

            # Сравнение 3 источников количества -> красный при расхождении
            qty_from_object = object_rows[obj_num]["qty"] if obj_num in object_rows else 0.0
            qty_values = [v for v in (qty_from_object, qty_consumer, total_qty) if v > 0]
            if len(set(round(v, 4) for v in qty_values)) > 1:
                data.quantity_mismatch = True

            # --- Price ---
            if obj_num in object_rows:
                data.total_price = object_rows[obj_num]["sum"]

            if data.quantity_packages > 0 and data.total_price > 0:
                data.unit_price = round(data.total_price / data.quantity_packages, 2)

            results.append(data)

        return results

    # ------------------------------------------------------------------
    # Common helpers
    # ------------------------------------------------------------------

    def _extract_texts(self, soup: BeautifulSoup) -> list:
        for tag in soup.find_all(["style", "script", "svg"]):
            tag.decompose()
        texts = []
        for text in soup.stripped_strings:
            cleaned = " ".join(text.split())
            if cleaned:
                texts.append(cleaned)
        return texts

    def _find_value_after_exact(self, texts: list, label: str) -> str:
        for i, t in enumerate(texts):
            if t.strip() == label and i + 1 < len(texts):
                return texts[i + 1]
        return ""

    def _extract_contract_number(self, texts: list, data: ContractData) -> None:
        data.contract_number = self._find_value_after_exact(texts, "Номер контракта")

    def _extract_notice_number(self, texts: list, data: ContractData) -> None:
        for i, t in enumerate(texts):
            if "Номер извещения" in t and i + 1 < len(texts):
                data.notice_number = texts[i + 1]
                break

    def _extract_procurement_method(self, texts: list, data: ContractData) -> None:
        for i, t in enumerate(texts):
            if "Способ определения поставщика" in t and i + 1 < len(texts):
                data.procurement_method = texts[i + 1].lower()
                break

    def _extract_customer_info(self, texts: list, data: ContractData) -> None:
        in_customer = False
        for i, t in enumerate(texts):
            if "2.1. Информация о заказчике" in t:
                in_customer = True
            if "2.2. Информация о поставщике" in t:
                break
            if in_customer and t == "Сокращенное наименование" and i + 1 < len(texts):
                data.customer_short_name = texts[i + 1]
                break

    def _extract_supplier_info(self, texts: list, data: ContractData) -> None:
        in_supplier = False
        for i, t in enumerate(texts):
            if "2.2. Информация о поставщике" in t:
                in_supplier = True
            if "3. Предмет контракта" in t:
                break
            if in_supplier and t == "Сокращенное наименование" and i + 1 < len(texts):
                data.supplier_short_name = texts[i + 1]
                break

    def _extract_contract_date(self, texts: list, data: ContractData) -> None:
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

    def _parse_pdf(self, path: Path) -> list:
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
                    "PDF не содержит текстового слоя. Требуется OCR (Tesseract)."
                )
                return [data]

            texts = [
                " ".join(line.split())
                for line in full_text.split("\n")
                if line.strip()
            ]
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

            base.errors.append(
                "Нет данных о лекарственном препарате (раздел не найден) — строка пропущена"
            )
            base.dosage_form_empty = True
            return [base]

        except ImportError:
            data.errors.append(
                "Модуль pdfplumber не установлен. Установите: pip install pdfplumber"
            )
        except Exception as e:
            data.errors.append(f"Ошибка чтения PDF: {e}")

        return [data]

import re
from dataclasses import dataclass, field
from pathlib import Path

from bs4 import BeautifulSoup


@dataclass
class ContractData:
    contract_number: str = ""
    contract_date: str = ""
    notice_number: str = ""
    procurement_method: str = ""
    customer_short_name: str = ""
    supplier_short_name: str = ""
    service_contract: str = ""
    mnn: str = ""
    trade_name: str = ""
    dosage_form: str = ""
    dosage_form_mnn_only: bool = False
    dosage_form_empty: bool = False
    dosage_form_uncertain: bool = False
    unit: str = "упак."
    quantity_packages: float = 0.0
    quantity_mismatch: bool = False
    quantity_all_values: str = ""
    price_mismatch: bool = False
    manufacturer: str = ""
    unit_price: float = 0.0
    total_price: float = 0.0
    source_file: str = ""
    errors: list = field(default_factory=list)

    @property
    def is_empty(self) -> bool:
        return (not self.mnn and not self.trade_name
                and self.quantity_packages == 0 and self.total_price == 0)


class ContractParser:
    _SOLID_FORMS = ("таблет", "капсул", "драже", "пастил", "пилюл", "гранул")

    def parse_file(self, file_path):
        path = Path(file_path)
        if not path.exists():
            return []
        s = path.suffix.lower()
        if s in (".html", ".htm"):
            return self._parse_html(path)
        return []

    def _parse_html(self, path):
        content = open(path, encoding="utf-8").read()
        soup = BeautifulSoup(content, "lxml")
        texts = self._extract_texts(soup)
        base = ContractData(source_file=str(path))
        self._extract_contract_number(texts, base)
        self._extract_notice_number(texts, base)
        self._extract_procurement_method(texts, base)
        self._extract_customer_info(texts, base)
        self._extract_supplier_info(texts, base)
        self._extract_contract_date(texts, base)
        self._extract_service_contract(texts, base)
        object_rows = self._parse_object_table(texts)
        drug_entries = self._parse_drug_entries(texts)
        if drug_entries:
            return self._build_contracts(base, object_rows, drug_entries)
        base.errors.append("Нет данных о лекарственном препарате")
        base.dosage_form_empty = True
        return [base]

    @staticmethod
    def _parse_num(s):
        if not s:
            return 0.0
        m = re.match(r"^\s*(\d[\d\s]*(?:[.,]\d+)?)", s.strip())
        if not m:
            return 0.0
        try:
            return float(m.group(1).replace(" ", "").replace(",", "."))
        except ValueError:
            return 0.0

    @staticmethod
    def _format_num(v):
        if v == int(v):
            return str(int(v))
        return ("%f" % v).rstrip("0").rstrip(".").replace(".", ",")

    def _pack_word(self, pack_type, count):
        pt = (pack_type or "").strip().upper()
        n = count if count and count > 0 else 1
        if "ШПРИЦ-РУЧК" in pt:
            return "шприц-ручка" if n <= 1 else "шприц-ручки"
        if "ШПРИЦ" in pt:
            return "шприц" if n <= 1 else "шприцы"
        if "ИНГАЛЯТОР" in pt:
            return "ИНГАЛЯТОР"
        return ""

    def _build_dosage_form(self, grls_form, mnn_object, qty_in_primary,
                           qty_primary_per_consumer, qty_per_consumer, pack_type):
        if grls_form:
            parts = grls_form.split(":", 1)
            if len(parts) == 2 and parts[1].strip():
                form_text = parts[1].strip()
            else:
                return grls_form.strip(), True, False, True
        elif mnn_object:
            return mnn_object, True, False, True
        else:
            return "ПУСТО", False, True, False
        low = form_text.lower()
        if any(k in low for k in self._SOLID_FORMS):
            if qty_per_consumer > 0:
                return f"{form_text} №{self._format_num(qty_per_consumer)}", False, False, False
            return form_text, False, False, True
        if "мл" in low:
            unit = "мл"
        elif "доз" in low:
            unit = "доз"
        else:
            unit = ""
        word = self._pack_word(pack_type, qty_primary_per_consumer)
        n_str = self._format_num(qty_primary_per_consumer if qty_primary_per_consumer > 0 else 1)
        if unit and qty_in_primary > 0:
            suffix = f"{self._format_num(qty_in_primary)} {unit}"
            if word:
                suffix += f" {word}"
            suffix += f" №{n_str}"
        else:
            suffix = f"{word} №{n_str}".strip() if word else f"№{n_str}"
        return f"{form_text} {suffix}", False, False, (not bool(word))

    def _parse_object_table(self, texts):
        result = {}
        hp = None
        for i, t in enumerate(texts):
            if "Наименование объекта закупки" in t:
                hp = i; break
        if hp is None:
            return result
        ds = None
        for j in range(hp + 9, min(hp + 40, len(texts))):
            if (texts[j] == "9" and j >= 8 and texts[j-1]=="8" and texts[j-2]=="7"
                and texts[j-3]=="6" and texts[j-4]=="5" and texts[j-5]=="4"
                and texts[j-6]=="3" and texts[j-7]=="2" and texts[j-8]=="1"):
                ds = j + 1; break
        if ds is None:
            return result
        idx = ds
        while idx + 8 < len(texts):
            if texts[idx] == "Итого:" or "Дополнительная информация" in texts[idx]:
                break
            try:
                rn = int(texts[idx])
            except ValueError:
                break
            result[rn] = {"name": texts[idx+1], "qty": self._parse_num(texts[idx+4]),
                          "unit_price": self._parse_num(texts[idx+5]),
                          "sum": self._parse_num(texts[idx+8])}
            idx += 9
        return result

    def _parse_drug_entries(self, texts):
        entries = []
        ss = None
        for i, t in enumerate(texts):
            if "Дополнительная информация о лекарственном препарате" in t:
                ss = i; break
        if ss is None:
            return entries
        ds = None
        for j in range(ss + 6, min(ss + 30, len(texts))):
            if (texts[j]=="6" and j>=5 and texts[j-1]=="5" and texts[j-2]=="4"
                and texts[j-3]=="3" and texts[j-4]=="2" and texts[j-5]=="1"):
                ds = j + 1; break
        if ds is None:
            return entries
        se = len(texts)
        for i in range(ds, len(texts)):
            if "4. Условия контракта" in texts[i] or "Дополнительная информация об адресе" in texts[i]:
                se = i; break
        idx = ds
        while idx + 5 < se:
            try:
                int(texts[idx]); obj_num = int(texts[idx+1])
            except ValueError:
                idx += 1; continue
            e = {"obj_num": obj_num, "trade_name_raw": texts[idx+2],
                 "primary_pack_type": texts[idx+4], "qty_consumer": self._parse_num(texts[idx+5]),
                 "grls_form": "", "holder": "", "producer": "", "qty_in_primary": 0.0,
                 "qty_primary_per_consumer": 0.0, "qty_per_consumer": 0.0, "total_qty": 0.0}
            idx += 6
            while idx < se:
                t = texts[idx]
                if t == "МНН и форма выпуска в соответствии с ГРЛС" and idx+1 < se:
                    nv = texts[idx+1]
                    if nv in ("Наименование держателя или владельца РУ","Производитель",
                              "Количество лекарственных форм в первичной упаковке","Признак включения в ЖНВЛП"):
                        idx += 1
                    else:
                        e["grls_form"] = nv; idx += 2
                elif t == "Наименование держателя или владельца РУ" and idx+1 < se:
                    e["holder"] = texts[idx+1]; idx += 2
                elif t == "Производитель" and idx+1 < se:
                    e["producer"] = texts[idx+1]; idx += 2
                elif t == "Количество лекарственных форм в первичной упаковке" and idx+1 < se:
                    e["qty_in_primary"] = self._parse_num(texts[idx+1]); idx += 2
                elif t == "Количество первичных упаковок в потребительской упаковке" and idx+1 < se:
                    e["qty_primary_per_consumer"] = self._parse_num(texts[idx+1]); idx += 2
                elif t == "Количество потребительских единиц в потребительской упаковке" and idx+1 < se:
                    e["qty_per_consumer"] = self._parse_num(texts[idx+1]); idx += 2
                elif t == "Общее количество" and idx+1 < se:
                    e["total_qty"] = self._parse_num(texts[idx+1]); idx += 2
                elif t == "Признак включения в ЖНВЛП" and idx+1 < se:
                    idx += 2; break
                else:
                    idx += 1
            entries.append(e)
        return entries

    def _build_contracts(self, base, object_rows, drug_entries):
        results = []
        for entry in drug_entries:
            d = ContractData(source_file=base.source_file, contract_number=base.contract_number,
                contract_date=base.contract_date, notice_number=base.notice_number,
                procurement_method=base.procurement_method, customer_short_name=base.customer_short_name,
                supplier_short_name=base.supplier_short_name, service_contract=base.service_contract)
            gf = entry["grls_form"]; on = entry["obj_num"]
            mnn_grls = gf.split(":",1)[0].strip() if gf and ":" in gf else ""
            mnn_object = object_rows[on]["name"] if on in object_rows else ""
            if mnn_grls and mnn_object:
                if (mnn_grls.upper()==mnn_object.upper() or mnn_grls.upper() in mnn_object.upper()
                    or mnn_object.upper() in mnn_grls.upper()):
                    d.mnn = mnn_grls
                else:
                    d.mnn = f"{mnn_grls}, {mnn_object}"
            else:
                d.mnn = mnn_grls or mnn_object
            (d.dosage_form, d.dosage_form_mnn_only, d.dosage_form_empty, d.dosage_form_uncertain) = \
                self._build_dosage_form(gf, mnn_object, entry["qty_in_primary"],
                    entry["qty_primary_per_consumer"], entry["qty_per_consumer"], entry["primary_pack_type"])
            tnr = entry["trade_name_raw"]
            if tnr:
                m = re.match(r"^(.+?),\s*ЛП", tnr)
                d.trade_name = m.group(1).strip() if m else tnr.split(",")[0].strip()
            producer = entry["producer"]; holder = entry["holder"]
            pc = re.sub(r"\s*\(\d+\)\s*$", "", producer).strip()
            if holder and pc:
                d.manufacturer = pc if pc.upper().startswith(holder.upper()) else f"{holder}/ {pc}"
            else:
                d.manufacturer = pc or holder
            tq = entry["total_qty"]; qpc = entry["qty_per_consumer"]; qc = entry["qty_consumer"]
            if qpc > 0 and tq > 0:
                d.quantity_packages = tq / qpc
            elif tq > 0:
                d.quantity_packages = float(tq)
            if abs(d.quantity_packages - round(d.quantity_packages)) < 1e-6:
                d.quantity_packages = float(round(d.quantity_packages))
            qfo = object_rows[on]["qty"] if on in object_rows else 0.0
            named = {}
            if qfo > 0: named["Объём закупки"] = qfo
            if qc > 0: named["Потреб. ед."] = qc
            if tq > 0: named["Общее кол-во"] = tq
            if len(set(round(v,4) for v in named.values())) > 1:
                d.quantity_mismatch = True
                d.quantity_all_values = ", ".join(f"{k}: {int(v) if v==int(v) else v}" for k,v in named.items())
            if on in object_rows:
                d.total_price = object_rows[on]["sum"]
            if d.quantity_packages > 0 and d.total_price > 0:
                d.unit_price = round(d.total_price / d.quantity_packages, 2)
            oup = object_rows[on]["unit_price"] if on in object_rows else 0.0
            if oup > 0 and qpc > 0 and d.quantity_packages > 0 and d.total_price > 0:
                calc = round(oup * qpc * d.quantity_packages, 2)
                if abs(calc - round(d.total_price, 2)) > 0.0001:
                    d.price_mismatch = True
                    d.errors.append(f"Расхождение цены: расчёт {calc} != сумма {round(d.total_price,2)}")
            results.append(d)
        return results

    def _extract_texts(self, soup):
        for tag in soup.find_all(["style","script","svg"]):
            tag.decompose()
        return [" ".join(s.split()) for s in soup.stripped_strings if s.split()]

    def _find_value_after_exact(self, texts, label):
        for i, t in enumerate(texts):
            if t.strip() == label and i+1 < len(texts):
                return texts[i+1]
        return ""

    def _extract_contract_number(self, texts, d):
        d.contract_number = self._find_value_after_exact(texts, "Номер контракта")

    def _extract_notice_number(self, texts, d):
        for i, t in enumerate(texts):
            if "Номер извещения" in t and i+1 < len(texts):
                d.notice_number = texts[i+1]; break

    def _extract_procurement_method(self, texts, d):
        for i, t in enumerate(texts):
            if "Способ определения поставщика" in t and i+1 < len(texts):
                d.procurement_method = texts[i+1].lower(); break

    def _extract_customer_info(self, texts, d):
        inc = False
        for i, t in enumerate(texts):
            if "2.1. Информация о заказчике" in t: inc = True
            if "2.2. Информация о поставщике" in t: break
            if inc and t == "Сокращенное наименование" and i+1 < len(texts):
                d.customer_short_name = texts[i+1]; break

    def _extract_supplier_info(self, texts, d):
        ins = False
        for i, t in enumerate(texts):
            if "2.2. Информация о поставщике" in t: ins = True
            if "3. Предмет контракта" in t: break
            if ins and t == "Сокращенное наименование" and i+1 < len(texts):
                d.supplier_short_name = texts[i+1]; break

    def _extract_service_contract(self, texts, d):
        """ГК на услугу из п.4.3 («Дополнительная информация об адресе»)."""
        for i, t in enumerate(texts):
            if "Дополнительная информация об адресе" in t:
                for j in range(i + 1, min(i + 8, len(texts))):
                    m = re.search(
                        r"контракта\s+от\s+(\d{2}\.\d{2}\.\d{4})\s*г\.?\s*№\s*([^\s,;]+)",
                        texts[j])
                    if m:
                        d.service_contract = f"{m.group(2)} от {m.group(1)}"
                        return
                break

    def _extract_contract_date(self, texts, d):
        for i, t in enumerate(texts):
            if "Дата начала исполнения контракта" in t and i+1 < len(texts):
                m = re.match(r"\d{2}\.\d{2}\.\d{4}", texts[i+1].strip())
                if m:
                    d.contract_date = m.group(0); return
                break
        sd = []
        for i, t in enumerate(texts):
            if t == "Дата и время подписания:" and i+1 < len(texts):
                m = re.match(r"\d{2}\.\d{2}\.\d{4}", texts[i+1])
                if m: sd.append(m.group(0))
        if sd:
            d.contract_date = sd[-1]

# -*- coding: utf-8 -*-
"""
Скачивание контрактов с ЕИС (zakupki.gov.ru) по номеру.
Ищет в реестре контрактов по любому номеру (реестровый/извещение/ИКЗ/внутренний),
скачивает HTML печатной формы и PDF контракта (по внутреннему номеру в имени файла).

Тест доступа:      python eis_downloader.py test
Скачать контракты: python eis_downloader.py "C:\\папка" 01085-ФЛ/2026 2745313582726000543
"""
import re
import sys
from pathlib import Path
from datetime import date

try:
    import requests
except ImportError:
    requests = None
from bs4 import BeautifulSoup

BASE = "https://zakupki.gov.ru"
SEARCH_URL = BASE + "/epz/contract/search/results.html"
DOCS_URL = BASE + "/epz/contract/contractCard/document-info.html"
PRINTFORM_URL = BASE + "/epz/contract/printForm/view.html"

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ru-RU,ru;q=0.9",
}

SEARCH_PARAMS = {
    "morphology": "on", "fz44": "on",
    "contractStageList_0": "on", "contractStageList_1": "on",
    "contractStageList_2": "on", "contractStageList_3": "on",
    "contractStageList": "0,1,2,3",
    "sortBy": "UPDATE_DATE", "pageNumber": "1",
    "sortDirection": "false", "recordsPerPage": "_10",
}


def parse_numbers(text):
    """Разбор поля ввода: через запятую/точку с запятой/столбцом, пробелы обрезаются."""
    parts = re.split(r"[,;\n\r\t]+", text or "")
    return [p.strip() for p in parts if p.strip()]


def safe_name(name):
    """Имя для файла/папки без запрещённых символов Windows."""
    name = re.sub(r'[\\/:*?"<>|]+', "-", name or "")
    return name.strip().strip(".") or "file"


def _norm(s):
    """Нормализация номера для сравнения: только буквы/цифры в нижнем регистре."""
    return re.sub(r"[^0-9a-zа-я]", "", (s or "").lower())


def extract_internal_number(html):
    """Внутренний номер контракта из HTML печатной формы (после 'Номер контракта')."""
    soup = BeautifulSoup(html, "lxml")
    texts = [" ".join(s.split()) for s in soup.stripped_strings if s.split()]
    for i, t in enumerate(texts):
        if t.strip() == "Номер контракта" and i + 1 < len(texts):
            return texts[i + 1].strip()
    return ""


def documents_from_html(html):
    """Список (url, имя) всех файлов-вложений со страницы документов карточки."""
    soup = BeautifulSoup(html, "lxml")
    out, seen = [], set()
    for a in soup.find_all("a", href=lambda h: h and "filestore/public" in h):
        href = a.get("href")
        if href.startswith("/"):
            href = BASE + href
        uid = href.split("uid=")[-1]
        if uid in seen:
            continue
        seen.add(uid)
        name = ""
        node = a
        for _ in range(5):
            node = node.parent
            if node is None:
                break
            txt = " ".join(node.get_text().split())
            if txt:
                name = re.split(r"\s*№\s*\d{6,}", txt)[0].strip()
                if name:
                    break
        out.append((href, name))
    return out


def pdf_matches_internal(name, internal):
    """Файл подходит, если в имени есть внутренний номер контракта."""
    inum = _norm(internal)
    return bool(inum) and inum in _norm(name)


def find_existing_folder(base_dir, folder_name):
    """Поиск папки контракта где-либо в базовой папке (защита от повторов)."""
    base = Path(base_dir)
    if not base.exists():
        return None
    target = folder_name.lower()
    for p in base.rglob("*"):
        if p.is_dir() and p.name.lower() == target:
            return p
    return None


class EISClient:
    def __init__(self, timeout=30):
        if requests is None:
            raise RuntimeError("Не установлен модуль requests (pip install requests)")
        self.s = requests.Session()
        self.s.headers.update(HEADERS)
        self.timeout = timeout

    def _blocked(self, text):
        low = text.lower()
        marks = ("incapsula", "captcha", "доступ ограничен", "запрос отклонён",
                 "request unsuccessful", "ддос", "ddos-guard")
        return any(m in low for m in marks)

    def test_connection(self):
        """Проверка доступа к ЕИС. Возвращает (ok, сообщение)."""
        try:
            r = self.s.get(SEARCH_URL,
                           params=dict(SEARCH_PARAMS, searchString="лекарственный"),
                           timeout=self.timeout)
        except Exception as e:
            return False, f"Нет связи с ЕИС: {e}"
        if r.status_code != 200:
            return False, f"ЕИС вернул код {r.status_code}"
        if self._blocked(r.text) or len(r.text) < 1500:
            return False, "ЕИС не пускает (антибот/капча)"
        return True, "ЕИС: всё хорошо"

    def search_reestr(self, number):
        r = self.s.get(SEARCH_URL, params=dict(SEARCH_PARAMS, searchString=number),
                       timeout=self.timeout)
        r.raise_for_status()
        if self._blocked(r.text):
            raise RuntimeError("ЕИС не пускает (антибот)")
        nums = re.findall(r"reestrNumber=(\d{15,25})", r.text)
        return nums[0] if nums else None  # первый = самый свежий (сортировка по дате)

    def get_printform_html(self, reestr):
        r = self.s.get(PRINTFORM_URL, params={"contractReestrNumber": reestr},
                       timeout=self.timeout)
        r.raise_for_status()
        return r.text

    def get_documents(self, reestr):
        r = self.s.get(DOCS_URL, params={"reestrNumber": reestr}, timeout=self.timeout)
        r.raise_for_status()
        return documents_from_html(r.text)

    def download(self, url, dest_dir):
        r = self.s.get(url, timeout=self.timeout, stream=True)
        r.raise_for_status()
        fn = self._filename(r, url)
        path = Path(dest_dir) / fn
        with open(path, "wb") as f:
            for chunk in r.iter_content(8192):
                if chunk:
                    f.write(chunk)
        return fn

    def _filename(self, r, url):
        cd = r.headers.get("Content-Disposition", "")
        m = re.search(r"filename\*=UTF-8''([^;]+)", cd, re.I)
        if m:
            return safe_name(requests.utils.unquote(m.group(1)))
        m = re.search(r'filename="?([^";]+)"?', cd)
        if m:
            fn = m.group(1)
            # HTTP-заголовки приходят как latin-1; восстанавливаем кириллицу из UTF-8
            try:
                fn = fn.encode("latin-1").decode("utf-8")
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass
            return safe_name(fn)
        return "file_" + url.split("uid=")[-1][:8] + ".bin"


def download_contract(client, number, base_dir, log=print):
    """
    Качает один контракт: печатная форма HTML + PDF с внутренним номером.
    Возвращает dict: {status: ok|skip|error, message, internal, reestr, folder, files}.
    """
    try:
        reestr = client.search_reestr(number)
    except Exception as e:
        return {"status": "error", "message": f"не скачался((( ({e})", "internal": number}
    if not reestr:
        return {"status": "error", "message": "не найден в ЕИС, не скачался(((",
                "internal": number}

    try:
        html = client.get_printform_html(reestr)
    except Exception as e:
        return {"status": "error", "message": f"не скачался((( (печатная форма: {e})",
                "internal": number, "reestr": reestr}

    internal = extract_internal_number(html) or number
    folder_name = safe_name(internal)

    existing = find_existing_folder(base_dir, folder_name)
    if existing:
        return {"status": "skip", "message": f"уже скачан ({existing})",
                "internal": internal, "reestr": reestr}

    folder = Path(base_dir) / date.today().strftime("%Y-%m-%d") / folder_name
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "Печатная форма.html").write_text(html, encoding="utf-8")
    saved = ["Печатная форма.html"]

    date_re = re.compile(r"\d{2}\.\d{2}\.\d{4}")
    try:
        for url, name in client.get_documents(reestr):
            if not pdf_matches_internal(name, internal):
                continue
            try:
                fn = client.download(url, folder)
            except Exception as e:
                log(f"  файл не скачался: {e}")
                continue
            # нужен основной PDF контракта: внутр.номер в имени, PDF, без даты
            # (файлы с датой — доп.соглашения / приёмка, не берём)
            keep = (fn.lower().endswith(".pdf")
                    and pdf_matches_internal(fn, internal)
                    and not date_re.search(fn))
            if keep:
                saved.append(fn)
            else:
                try:
                    (folder / fn).unlink()
                except OSError:
                    pass
    except Exception as e:
        log(f"  список документов не получен: {e}")

    return {"status": "ok", "message": f"скачано файлов: {len(saved)}",
            "internal": internal, "reestr": reestr, "folder": str(folder), "files": saved}


def _cli():
    args = sys.argv[1:]
    if not args or args[0] == "test":
        ok, msg = EISClient().test_connection()
        print(msg)
        sys.exit(0 if ok else 1)
    base_dir = args[0]
    client = EISClient()
    for num in parse_numbers(",".join(args[1:])):
        res = download_contract(client, num, base_dir)
        print(f"{num}: {res['status']} — {res['message']}")


if __name__ == "__main__":
    _cli()

# -*- coding: utf-8 -*-
"""
Хранилище мониторинга контрактов (переживает перезапуск).
Файл JSON в домашней папке пользователя. Записи старше 30 дней удаляются.
"""
import json
import re
from pathlib import Path
from datetime import datetime, timedelta

STORE_PATH = Path.home() / ".contract_parser_monitor.json"
KEEP_DAYS = 30


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def _norm(s):
    """Номер для сравнения: только буквы/цифры, регистр/пробелы/слэши не важны."""
    return re.sub(r"[^0-9a-zа-яё]", "", str(s or "").lower())


def find_key(data, number, reestr=""):
    """Ключ уже имеющейся записи о том же контракте или None. Тот же контракт =
    совпал реестровый номер либо номер (в любом написании); номер может быть и
    реестровым, если его так ввели вручную."""
    n, r = _norm(number), _norm(reestr)
    for key, c in data["contracts"].items():
        ids = {_norm(key), _norm(c.get("number")), _norm(c.get("reestr"))} - {""}
        if (r and r in ids) or (n and n in ids):
            return key
    return None


def dedupe(data):
    """Склеивает записи об одном контракте (остаётся более ранняя, недостающие
    поля берутся из дубля). Возвращает число удалённых дублей."""
    removed = 0
    kept = {}
    for key, c in list(data["contracts"].items()):
        probe = {"contracts": kept}
        first = find_key(probe, c.get("number") or key, c.get("reestr", ""))
        if first is None:
            kept[key] = c
            continue
        base = kept[first]
        if not base.get("last_check") and c.get("last_check"):
            # проверенная запись ценнее непроверенной: берём её данные, дату добавления храним старую
            c = dict(c, added=min(base.get("added") or c.get("added", ""),
                                  c.get("added") or base.get("added", "")))
            base, c = c, base
            kept[first] = base
        elif c.get("status") == "changed":
            base["status"] = "changed"
            base["last_event"] = c.get("last_event", base.get("last_event", ""))
        for field in ("reestr", "service", "service_start", "service_checked"):
            if not base.get(field) and c.get(field):
                base[field] = c[field]
        removed += 1
    data["contracts"] = kept
    return removed


def load():
    try:
        if STORE_PATH.exists():
            with open(STORE_PATH, encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and "contracts" in data:
                data = _prune(data)
                dedupe(data)
                return data
    except Exception:
        pass
    return {"contracts": {}}


def save(data):
    try:
        with open(STORE_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _prune(data):
    cutoff = datetime.now() - timedelta(days=KEEP_DAYS)
    keep = {}
    for k, v in data.get("contracts", {}).items():
        added = v.get("added", "")
        try:
            dt = datetime.strptime(added, "%Y-%m-%d %H:%M")
        except Exception:
            dt = datetime.now()
        if dt >= cutoff:
            keep[k] = v
    data["contracts"] = keep
    return data


def add_contract(data, number, reestr=""):
    """Добавить контракт под наблюдение (если ещё нет). key = reestr или number.
    Контракт, который уже есть под другим ключом (добавлен вручную по номеру, а
    теперь пришёл из ЕИС с реестровым — или наоборот), повторно не добавляется."""
    key = reestr or number
    if not key:
        return False
    found = find_key(data, number, reestr)
    if found is not None:
        # обновим реестр, если был пуст
        c = data["contracts"][found]
        if reestr and not c.get("reestr"):
            c["reestr"] = reestr
        return False
    data["contracts"][key] = {
        "number": number, "reestr": reestr,
        "dop_count": None, "dop_max": None,
        "last_check": "", "added": _now(),
        "status": "new", "last_event": "не проверялся",
        "service": "", "service_start": "", "service_checked": "",
    }
    return True


def set_service(data, key, service, service_start):
    """ГК на услугу и дата, с которой он действует (см. «Действует с» в сводке)."""
    c = data["contracts"].get(key)
    if c is None:
        return
    if service:
        c["service"] = service
    c["service_start"] = service_start or ""
    c["service_checked"] = _now()


def remove_contract(data, key):
    data["contracts"].pop(key, None)


def update_after_check(data, key, reestr, dop_count, dop_max):
    """Обновляет запись после проверки. Возвращает True, если появились изменения.
    Статус «changed» НЕ сбрасывается автоматически — только вручную (acknowledge)."""
    c = data["contracts"].get(key)
    if c is None:
        return False
    prev_count = c.get("dop_count")
    prev_max = c.get("dop_max")
    changed = (prev_count is not None and
               ((dop_count or 0) > (prev_count or 0) or (dop_max or 0) > (prev_max or 0)))
    c["reestr"] = reestr or c.get("reestr", "")
    c["dop_count"] = dop_count
    c["dop_max"] = dop_max
    c["last_check"] = _now()
    if changed:
        c["status"] = "changed"
        c["last_event"] = f"{_now()}: НОВОЕ доп. соглашение (стало {dop_count}, макс №{dop_max})"
    elif c.get("status") != "changed":
        c["status"] = "nochange"
        c["last_event"] = f"{_now()}: без изменений"
    return changed


def set_error(data, key, msg):
    c = data["contracts"].get(key)
    if c is not None:
        c["last_check"] = _now()
        if c.get("status") != "changed":
            c["status"] = "error"
        c["last_event"] = f"{_now()}: ошибка — {msg}"


def acknowledge(data, key):
    """Сбросить статус «изменения» после того, как пользователь увидел."""
    c = data["contracts"].get(key)
    if c is not None and c.get("status") == "changed":
        c["status"] = "nochange"


def list_contracts(data):
    return list(data["contracts"].items())

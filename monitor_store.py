# -*- coding: utf-8 -*-
"""
Хранилище мониторинга контрактов (переживает перезапуск).
Файл JSON в домашней папке пользователя. Записи старше 30 дней удаляются.
"""
import json
from pathlib import Path
from datetime import datetime, timedelta

STORE_PATH = Path.home() / ".contract_parser_monitor.json"
KEEP_DAYS = 30


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def load():
    try:
        if STORE_PATH.exists():
            with open(STORE_PATH, encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and "contracts" in data:
                return _prune(data)
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
    """Добавить контракт под наблюдение (если ещё нет). key = reestr или number."""
    key = reestr or number
    if not key:
        return False
    if key in data["contracts"]:
        # обновим отображаемый номер/реестр, если был пуст
        c = data["contracts"][key]
        if reestr and not c.get("reestr"):
            c["reestr"] = reestr
        return False
    data["contracts"][key] = {
        "number": number, "reestr": reestr,
        "dop_count": None, "dop_max": None,
        "last_check": "", "added": _now(),
        "status": "new", "last_event": "не проверялся",
    }
    return True


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

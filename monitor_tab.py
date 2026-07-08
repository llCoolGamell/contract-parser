# -*- coding: utf-8 -*-
"""Вкладка «Мониторинг» — отслеживание появления доп. соглашений в контрактах."""
import re
import time
import webbrowser
from datetime import datetime, timedelta

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit,
    QComboBox, QTextEdit, QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QAbstractItemView, QFileDialog, QApplication,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QColor

import monitor_store as store
from eis_downloader import EISClient, parse_numbers, get_dop_info

BASE = "https://zakupki.gov.ru"
CARD_URL = BASE + "/epz/contract/contractCard/common-info.html?reestrNumber="

STATUS_LABEL = {
    "new": "⚪ Не проверялся",
    "nochange": "🟢 Нет изменений",
    "changed": "🔵 Внесены изменения",
    "error": "🔴 Ошибка проверки",
}
STATUS_COLOR = {
    "new": "#ECEFF1", "nochange": "#E8F5E9",
    "changed": "#E3F2FD", "error": "#FFEBEE",
}
INTERVALS = [("1 час", 1), ("3 часа", 3), ("12 часов", 12)]
FILTERS = ["Все", "🔵 Внесены изменения", "🟢 Нет изменений",
           "⚪ Не проверялся", "🔴 Ошибка"]
_FILTER_STATUS = {FILTERS[1]: "changed", FILTERS[2]: "nochange",
                  FILTERS[3]: "new", FILTERS[4]: "error"}

STATUS_COL = 3


def _norm(s):
    return re.sub(r"[^0-9a-zа-яё]", "", (s or "").lower())


class CheckThread(QThread):
    one_done = pyqtSignal(str, dict)
    all_done = pyqtSignal()

    def __init__(self, items):
        super().__init__()
        self.items = items

    def run(self):
        try:
            client = EISClient()
        except Exception as e:
            for key, _ in self.items:
                self.one_done.emit(key, {"error": str(e)})
            self.all_done.emit()
            return
        for i, (key, number) in enumerate(self.items):
            try:
                info = get_dop_info(client, number)
            except Exception as e:
                info = {"error": str(e)}
            self.one_done.emit(key, info)
            if i < len(self.items) - 1:
                time.sleep(2.0)
        self.all_done.emit()


class MonitorTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.data = store.load()
        store.save(self.data)
        self.check_thread = None
        self.interval_hours = 1
        self._last_tick = datetime.now()
        self._next_check = datetime.now() + timedelta(hours=self.interval_hours)
        self._changed_now = []
        self._build_ui()
        self.refresh_table()

        self.tick = QTimer(self)
        self.tick.timeout.connect(self._on_tick)
        self.tick.start(60 * 1000)

    # ---------------- UI ----------------
    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(8)

        add_row = QHBoxLayout()
        self.input = QLineEdit()
        self.input.setPlaceholderText("Вставьте номер(а) контракта и нажмите Enter или «Добавить»...")
        self.input.returnPressed.connect(self.add_from_input)
        add_row.addWidget(self.input)
        btn_add = QPushButton("Добавить")
        btn_add.setStyleSheet("QPushButton{background:#4CAF50;color:white;border:none;"
                              "border-radius:6px;padding:8px 16px;font-weight:bold;}")
        btn_add.clicked.connect(self.add_from_input)
        add_row.addWidget(btn_add)
        outer.addLayout(add_row)

        # поиск + фильтр
        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Поиск:"))
        self.search = QLineEdit()
        self.search.setPlaceholderText("Поиск по № ГК (по мере ввода, регистр/пробелы/слэши не важны)")
        self.search.textChanged.connect(self.refresh_table)
        search_row.addWidget(self.search, stretch=1)
        search_row.addWidget(QLabel("Статус:"))
        self.filter_box = QComboBox()
        for f in FILTERS:
            self.filter_box.addItem(f)
        self.filter_box.currentIndexChanged.connect(self.refresh_table)
        search_row.addWidget(self.filter_box)
        outer.addLayout(search_row)

        # управление проверкой
        ctl = QHBoxLayout()
        self.btn_check = QPushButton("🔄 Проверить ВСЕ")
        self.btn_check.setStyleSheet("QPushButton{background:#2196F3;color:white;border:none;"
                                     "border-radius:6px;padding:8px 16px;font-weight:bold;}"
                                     "QPushButton:disabled{background:#bbb;}")
        self.btn_check.clicked.connect(self.check_all)
        ctl.addWidget(self.btn_check)
        self.btn_check_sel = QPushButton("✔ Проверить выбранные")
        self.btn_check_sel.setStyleSheet("QPushButton{background:#3F51B5;color:white;border:none;"
                                         "border-radius:6px;padding:8px 16px;font-weight:bold;}"
                                         "QPushButton:disabled{background:#bbb;}")
        self.btn_check_sel.clicked.connect(self.check_selected)
        ctl.addWidget(self.btn_check_sel)
        ctl.addWidget(QLabel("Авто:"))
        self.interval_box = QComboBox()
        for name, _ in INTERVALS:
            self.interval_box.addItem(name)
        self.interval_box.currentIndexChanged.connect(self._on_interval_change)
        ctl.addWidget(self.interval_box)
        ctl.addStretch()
        outer.addLayout(ctl)

        # таблица
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(
            ["Номер ГК", "Доп. согл.", "Последняя проверка", "Статус"])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.cellClicked.connect(self._on_cell_clicked)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.Stretch)
        hh.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        outer.addWidget(self.table)

        # действия со строками
        act = QHBoxLayout()
        self.btn_ack = QPushButton("Обновить статус (просмотрено)")
        self.btn_ack.clicked.connect(self.acknowledge_selected)
        act.addWidget(self.btn_ack)
        self.btn_del = QPushButton("Удалить выбранные")
        self.btn_del.setStyleSheet("QPushButton{background:#f44336;color:white;border:none;"
                                   "border-radius:6px;padding:6px 12px;}")
        self.btn_del.clicked.connect(self.remove_selected)
        act.addWidget(self.btn_del)
        self.btn_clear = QPushButton("Очистить всё")
        self.btn_clear.setStyleSheet("QPushButton{background:#888;color:white;border:none;"
                                     "border-radius:6px;padding:6px 12px;}")
        self.btn_clear.clicked.connect(self.clear_all)
        act.addWidget(self.btn_clear)
        self.btn_export = QPushButton("Экспорт списка №")
        self.btn_export.clicked.connect(self.export_numbers)
        act.addWidget(self.btn_export)
        act.addStretch()
        outer.addLayout(act)

        outer.addWidget(QLabel("Журнал:"))
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(120)
        self.log_area.setStyleSheet("font-family:Consolas,monospace; font-size:11px;")
        outer.addWidget(self.log_area)

    def log(self, m):
        self.log_area.append(f"{datetime.now():%Y-%m-%d %H:%M}  {m}")

    # ---------------- таблица ----------------
    def refresh_table(self):
        flt = _FILTER_STATUS.get(self.filter_box.currentText())
        q = _norm(self.search.text())
        self.table.setRowCount(0)
        for key, c in store.list_contracts(self.data):
            if flt and c.get("status") != flt:
                continue
            num = c.get("number", key)
            if q and q not in _norm(num) and q not in _norm(key):
                continue
            r = self.table.rowCount()
            self.table.insertRow(r)
            dop = ""
            if c.get("dop_count") is not None:
                dop = f"{c.get('dop_count')} (макс №{c.get('dop_max')})"
            status = c.get("status", "new")
            cells = [num, dop, c.get("last_check", ""), STATUS_LABEL.get(status, status)]
            color = QColor(STATUS_COLOR.get(status, "#FFFFFF"))
            for col, val in enumerate(cells):
                it = QTableWidgetItem(str(val))
                it.setBackground(color)
                if col == 0:
                    it.setData(Qt.UserRole, key)
                if col == STATUS_COL and status == "changed":
                    it.setToolTip("Клик — открыть карточку контракта в ЕИС")
                self.table.setItem(r, col, it)

    def _selected_keys(self):
        keys, seen = [], set()
        for it in self.table.selectedItems():
            k = self.table.item(it.row(), 0).data(Qt.UserRole)
            if k and k not in seen:
                seen.add(k)
                keys.append(k)
        return keys

    def _on_cell_clicked(self, row, col):
        if col != STATUS_COL:
            return
        key = self.table.item(row, 0).data(Qt.UserRole)
        c = self.data["contracts"].get(key, {})
        if c.get("status") != "changed":
            return
        reestr = c.get("reestr")
        if not reestr:
            QMessageBox.information(self, "ЕИС", "Сначала проверьте контракт (нет реестрового номера).")
            return
        try:
            webbrowser.open(CARD_URL + str(reestr))
            self.log(f"Открыта карточка ЕИС: {c.get('number', key)}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть браузер: {e}")

    # ---------------- добавление/экспорт ----------------
    def add_from_input(self):
        nums = parse_numbers(self.input.text())
        added = sum(1 for n in nums if store.add_contract(self.data, n))
        if added:
            store.save(self.data)
            self.refresh_table()
            self.log(f"Добавлено в мониторинг: {added}")
        self.input.clear()

    def add_contracts(self, pairs):
        """Авто-добавление из ЕИС: pairs = [(number, reestr), ...]."""
        added = 0
        for number, reestr in pairs:
            if store.add_contract(self.data, number, reestr):
                added += 1
        if added:
            store.save(self.data)
            self.refresh_table()
            self.log(f"Из ЕИС добавлено в мониторинг: {added}")

    def export_numbers(self):
        rows = store.list_contracts(self.data)
        if not rows:
            QMessageBox.information(self, "Экспорт", "Список пуст.")
            return
        text = "\n".join(c.get("number", key) for key, c in rows)
        try:
            QApplication.clipboard().setText(text)
        except Exception:
            pass
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить список номеров",
                                              "контракты_мониторинг.txt", "Текст (*.txt)")
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(text)
                self.log(f"Экспортировано номеров: {len(rows)} → {path}")
                QMessageBox.information(self, "Экспорт",
                                       f"Сохранено {len(rows)} номеров.\nТакже скопировано в буфер обмена.")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", str(e))
        else:
            QMessageBox.information(self, "Экспорт",
                                   f"{len(rows)} номеров скопировано в буфер обмена.")

    # ---------------- действия ----------------
    def acknowledge_selected(self):
        keys = self._selected_keys()
        if not keys:
            QMessageBox.information(self, "Мониторинг", "Выберите строки.")
            return
        for k in keys:
            store.acknowledge(self.data, k)
        store.save(self.data)
        self.refresh_table()

    def remove_selected(self):
        keys = self._selected_keys()
        if not keys:
            QMessageBox.information(self, "Мониторинг", "Выберите строки.")
            return
        if QMessageBox.question(self, "Удалить", f"Удалить выбранные ({len(keys)}) из мониторинга?",
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            for k in keys:
                store.remove_contract(self.data, k)
            store.save(self.data)
            self.refresh_table()

    def clear_all(self):
        if not self.data["contracts"]:
            return
        if QMessageBox.question(self, "Очистить всё", "Удалить ВСЕ контракты из мониторинга?",
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.data["contracts"] = {}
            store.save(self.data)
            self.refresh_table()

    def _on_interval_change(self, idx):
        self.interval_hours = INTERVALS[idx][1]
        self._next_check = datetime.now() + timedelta(hours=self.interval_hours)
        self.log(f"Интервал авто-проверки: {INTERVALS[idx][0]}")

    # ---------------- проверка ----------------
    def check_all(self):
        items = [(key, c.get("number", key)) for key, c in store.list_contracts(self.data)]
        if not items:
            QMessageBox.information(self, "Мониторинг", "Список пуст — добавьте контракты.")
            return
        self._start_check(items)

    def check_selected(self):
        keys = self._selected_keys()
        if not keys:
            QMessageBox.information(self, "Мониторинг", "Выберите строки для проверки.")
            return
        items = [(k, self.data["contracts"][k].get("number", k)) for k in keys
                 if k in self.data["contracts"]]
        self._start_check(items)

    def _start_check(self, items):
        if self.check_thread and self.check_thread.isRunning():
            QMessageBox.information(self, "Мониторинг", "Проверка уже идёт, подождите.")
            return
        self.btn_check.setEnabled(False)
        self.btn_check_sel.setEnabled(False)
        self.log(f"Проверка {len(items)} контракт(ов)...")
        self._changed_now = []
        self.check_thread = CheckThread(items)
        self.check_thread.one_done.connect(self._on_one)
        self.check_thread.all_done.connect(self._on_all)
        self.check_thread.start()

    def _on_one(self, key, info):
        if "error" in info:
            store.set_error(self.data, key, info["error"])
        else:
            changed = store.update_after_check(self.data, key, info.get("reestr", ""),
                                               info.get("dop_count", 0), info.get("dop_max", 0))
            if changed:
                self._changed_now.append(self.data["contracts"][key].get("number", key))
        self.refresh_table()

    def _on_all(self):
        store.save(self.data)
        self.btn_check.setEnabled(True)
        self.btn_check_sel.setEnabled(True)
        self._next_check = datetime.now() + timedelta(hours=self.interval_hours)
        if self._changed_now:
            for n in self._changed_now:
                self.log(f"⚠ ВНЕСЕНЫ ИЗМЕНЕНИЯ: {n}")
            QMessageBox.warning(self, "Внесены изменения!",
                                "Появились доп. соглашения по контрактам:\n\n• " +
                                "\n• ".join(self._changed_now) +
                                "\n\nКлик по синему статусу откроет карточку в ЕИС.")
        else:
            self.log("Проверка завершена: изменений нет.")

    # ---------------- умный таймер (учёт сна) ----------------
    def _on_tick(self):
        now = datetime.now()
        gap = (now - self._last_tick).total_seconds()
        self._last_tick = now
        if gap > 300:
            self._next_check = now + timedelta(minutes=15)
            self.log("Обнаружено пробуждение ПК — авто-проверка через 15 минут.")
            return
        if now >= self._next_check and self.data["contracts"]:
            if not (self.check_thread and self.check_thread.isRunning()):
                items = [(key, c.get("number", key))
                         for key, c in store.list_contracts(self.data)]
                self._start_check(items)

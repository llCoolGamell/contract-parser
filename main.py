"""
Приложение для извлечения данных из контрактов (HTML/PDF) и переноса в Excel.
Поддерживает drag-and-drop, выбор файлов через проводник,
выбор существующего Excel или создание нового.
"""

import sys
import os
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QHBoxLayout,
    QVBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QGroupBox,
    QRadioButton,
    QLineEdit,
    QProgressBar,
    QTextEdit,
    QSplitter,
    QAbstractItemView,
    QSizePolicy,
)
from PyQt5.QtCore import Qt, QMimeData, QSize, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon, QDragEnterEvent, QDropEvent, QPalette, QColor

from parser_engine import ContractParser, ContractData
from excel_handler import write_contracts_to_excel, create_new_excel, get_sheet_names


SUPPORTED_EXTENSIONS = {".html", ".htm", ".pdf"}


class FileListWidget(QListWidget):
    """Список файлов с поддержкой drag-and-drop."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setMinimumWidth(280)
        self.setStyleSheet(
            """
            QListWidget {
                border: 2px dashed #aaa;
                border-radius: 8px;
                padding: 8px;
                background-color: #fafafa;
                font-size: 12px;
            }
            QListWidget::item {
                padding: 6px 4px;
                border-bottom: 1px solid #eee;
            }
            QListWidget::item:selected {
                background-color: #d0e8ff;
            }
        """
        )

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path:
                    self._add_file_if_valid(file_path)
        else:
            super().dropEvent(event)

    def _add_file_if_valid(self, file_path: str) -> bool:
        path = Path(file_path)
        if path.is_dir():
            added = False
            for f in path.iterdir():
                if f.suffix.lower() in SUPPORTED_EXTENSIONS:
                    if self._add_single_file(str(f)):
                        added = True
            return added
        return self._add_single_file(file_path)

    def _add_single_file(self, file_path: str) -> bool:
        path = Path(file_path)
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            return False
        for i in range(self.count()):
            if self.item(i).data(Qt.UserRole) == file_path:
                return False
        item = QListWidgetItem(path.name)
        item.setData(Qt.UserRole, file_path)
        item.setToolTip(file_path)
        self.addItem(item)
        return True

    def add_files(self, file_paths: list[str]) -> int:
        added = 0
        for fp in file_paths:
            if self._add_file_if_valid(fp):
                added += 1
        return added

    def get_all_paths(self) -> list[str]:
        paths = []
        for i in range(self.count()):
            path = self.item(i).data(Qt.UserRole)
            if path:
                paths.append(path)
        return paths


class ProcessThread(QThread):
    """Поток для обработки файлов."""

    progress = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str, list)

    def __init__(
        self,
        file_paths: list[str],
        excel_path: str,
        sheet_name: str | None = None,
    ):
        super().__init__()
        self.file_paths = file_paths
        self.excel_path = excel_path
        self.sheet_name = sheet_name

    def run(self) -> None:
        parser = ContractParser()
        contracts: list[ContractData] = []
        errors: list[str] = []
        total = len(self.file_paths)

        for idx, fp in enumerate(self.file_paths):
            self.progress.emit(
                int((idx / total) * 100),
                f"Обработка: {Path(fp).name}",
            )
            try:
                result = parser.parse_file(fp)
                if result:
                    if result.errors:
                        for err in result.errors:
                            errors.append(f"{Path(fp).name}: {err}")
                    if result.contract_number:
                        contracts.append(result)
                    else:
                        errors.append(
                            f"{Path(fp).name}: Не удалось извлечь номер контракта"
                        )
                else:
                    errors.append(f"{Path(fp).name}: Неподдерживаемый формат файла")
            except Exception as e:
                errors.append(f"{Path(fp).name}: {e}")

        if not contracts:
            self.finished_signal.emit(
                False,
                "Не удалось извлечь данные ни из одного файла.",
                errors,
            )
            return

        self.progress.emit(90, "Запись в Excel...")

        success, msg = write_contracts_to_excel(
            self.excel_path, contracts, self.sheet_name
        )
        self.finished_signal.emit(success, msg, errors)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Контракты -> Excel")
        self.setMinimumSize(900, 550)
        self.resize(1050, 600)

        palette = self.palette()
        palette.setColor(QPalette.Window, QColor("#f0f2f5"))
        self.setPalette(palette)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        # --- Left panel ---
        left_group = QGroupBox("Файлы для обработки")
        left_group.setStyleSheet(
            """
            QGroupBox {
                font-size: 13px;
                font-weight: bold;
                border: 1px solid #ccc;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 18px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }
        """
        )
        left_layout = QVBoxLayout(left_group)

        hint_label = QLabel("Перетащите файлы сюда или нажмите кнопку ниже")
        hint_label.setAlignment(Qt.AlignCenter)
        hint_label.setStyleSheet("color: #888; font-size: 11px; padding: 4px;")
        left_layout.addWidget(hint_label)

        self.file_list = FileListWidget()
        left_layout.addWidget(self.file_list)

        btn_row = QHBoxLayout()
        self.btn_add = QPushButton("+ Добавить файлы")
        self.btn_add.setStyleSheet(
            """
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """
        )
        self.btn_add.clicked.connect(self.add_files)
        btn_row.addWidget(self.btn_add)

        self.btn_remove = QPushButton("Удалить выбранные")
        self.btn_remove.setStyleSheet(
            """
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
        """
        )
        self.btn_remove.clicked.connect(self.remove_selected)
        btn_row.addWidget(self.btn_remove)

        self.btn_clear = QPushButton("Очистить все")
        self.btn_clear.setStyleSheet(
            """
            QPushButton {
                background-color: #888;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #666;
            }
        """
        )
        self.btn_clear.clicked.connect(self.clear_files)
        btn_row.addWidget(self.btn_clear)

        left_layout.addLayout(btn_row)
        main_layout.addWidget(left_group, stretch=2)

        # --- Center panel ---
        center_widget = QWidget()
        center_layout = QVBoxLayout(center_widget)
        center_layout.setAlignment(Qt.AlignCenter)

        self.btn_transfer = QPushButton("ПЕРЕНЕСТИ\n>>>")
        self.btn_transfer.setMinimumSize(140, 100)
        self.btn_transfer.setStyleSheet(
            """
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 12px;
                padding: 20px;
                font-size: 18px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #bbb;
            }
        """
        )
        self.btn_transfer.clicked.connect(self.start_transfer)
        center_layout.addWidget(self.btn_transfer)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimumWidth(140)
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet(
            """
            QProgressBar {
                border: 1px solid #ccc;
                border-radius: 6px;
                text-align: center;
                height: 24px;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 5px;
            }
        """
        )
        center_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setWordWrap(True)
        self.status_label.setMaximumWidth(160)
        self.status_label.setStyleSheet("font-size: 11px; color: #555;")
        center_layout.addWidget(self.status_label)

        main_layout.addWidget(center_widget, stretch=1)

        # --- Right panel ---
        right_group = QGroupBox("Файл Excel для записи")
        right_group.setStyleSheet(
            """
            QGroupBox {
                font-size: 13px;
                font-weight: bold;
                border: 1px solid #ccc;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 18px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }
        """
        )
        right_layout = QVBoxLayout(right_group)

        self.radio_existing = QRadioButton("Существующий файл:")
        self.radio_existing.setChecked(True)
        self.radio_existing.toggled.connect(self._toggle_excel_mode)
        right_layout.addWidget(self.radio_existing)

        existing_row = QHBoxLayout()
        self.excel_path_edit = QLineEdit()
        self.excel_path_edit.setPlaceholderText("Путь к .xlsx файлу...")
        self.excel_path_edit.setStyleSheet(
            """
            QLineEdit {
                border: 1px solid #ccc;
                border-radius: 6px;
                padding: 6px;
                font-size: 12px;
            }
        """
        )
        existing_row.addWidget(self.excel_path_edit)

        self.btn_browse_excel = QPushButton("Обзор...")
        self.btn_browse_excel.setStyleSheet(
            """
            QPushButton {
                background-color: #607D8B;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #455A64;
            }
        """
        )
        self.btn_browse_excel.clicked.connect(self.browse_excel)
        existing_row.addWidget(self.btn_browse_excel)
        right_layout.addLayout(existing_row)

        right_layout.addSpacing(12)

        self.radio_new = QRadioButton("Создать новый файл:")
        self.radio_new.toggled.connect(self._toggle_excel_mode)
        right_layout.addWidget(self.radio_new)

        new_row = QHBoxLayout()
        self.new_path_edit = QLineEdit()
        self.new_path_edit.setPlaceholderText("Путь для нового файла...")
        self.new_path_edit.setEnabled(False)
        self.new_path_edit.setStyleSheet(
            """
            QLineEdit {
                border: 1px solid #ccc;
                border-radius: 6px;
                padding: 6px;
                font-size: 12px;
            }
            QLineEdit:disabled {
                background-color: #eee;
            }
        """
        )
        new_row.addWidget(self.new_path_edit)

        self.btn_browse_new = QPushButton("Обзор...")
        self.btn_browse_new.setEnabled(False)
        self.btn_browse_new.setStyleSheet(
            """
            QPushButton {
                background-color: #607D8B;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #455A64;
            }
            QPushButton:disabled {
                background-color: #bbb;
            }
        """
        )
        self.btn_browse_new.clicked.connect(self.browse_new_excel)
        new_row.addWidget(self.btn_browse_new)
        right_layout.addLayout(new_row)

        right_layout.addSpacing(16)

        # Log area
        log_label = QLabel("Лог обработки:")
        log_label.setStyleSheet("font-weight: bold; font-size: 12px;")
        right_layout.addWidget(log_label)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(200)
        self.log_area.setStyleSheet(
            """
            QTextEdit {
                border: 1px solid #ccc;
                border-radius: 6px;
                padding: 6px;
                font-size: 11px;
                background-color: #fff;
                font-family: Consolas, monospace;
            }
        """
        )
        right_layout.addWidget(self.log_area)

        right_layout.addStretch()
        main_layout.addWidget(right_group, stretch=2)

        self.process_thread = None

    def _toggle_excel_mode(self) -> None:
        is_existing = self.radio_existing.isChecked()
        self.excel_path_edit.setEnabled(is_existing)
        self.btn_browse_excel.setEnabled(is_existing)
        self.new_path_edit.setEnabled(not is_existing)
        self.btn_browse_new.setEnabled(not is_existing)

    def add_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите файлы контрактов",
            "",
            "Документы (*.html *.htm *.pdf);;Все файлы (*)",
        )
        if files:
            added = self.file_list.add_files(files)
            self.log(f"Добавлено файлов: {added}")

    def remove_selected(self) -> None:
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))

    def clear_files(self) -> None:
        self.file_list.clear()

    def browse_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл Excel",
            "",
            "Excel файлы (*.xlsx *.xls);;Все файлы (*)",
        )
        if path:
            self.excel_path_edit.setText(path)

    def browse_new_excel(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Создать новый файл Excel",
            "contracts.xlsx",
            "Excel файлы (*.xlsx)",
        )
        if path:
            if not path.endswith(".xlsx"):
                path += ".xlsx"
            self.new_path_edit.setText(path)

    def log(self, message: str) -> None:
        self.log_area.append(message)

    def start_transfer(self) -> None:
        file_paths = self.file_list.get_all_paths()
        if not file_paths:
            QMessageBox.warning(
                self,
                "Нет файлов",
                "Добавьте файлы контрактов для обработки.",
            )
            return

        if self.radio_existing.isChecked():
            excel_path = self.excel_path_edit.text().strip()
            if not excel_path:
                QMessageBox.warning(
                    self,
                    "Не выбран файл",
                    "Укажите путь к существующему Excel файлу.",
                )
                return
            if not Path(excel_path).exists():
                QMessageBox.warning(
                    self,
                    "Файл не найден",
                    f"Файл не найден: {excel_path}",
                )
                return
        else:
            excel_path = self.new_path_edit.text().strip()
            if not excel_path:
                QMessageBox.warning(
                    self,
                    "Не указан путь",
                    "Укажите путь для нового Excel файла.",
                )
                return
            if not excel_path.endswith(".xlsx"):
                excel_path += ".xlsx"
            create_new_excel(excel_path)
            self.log(f"Создан новый файл: {excel_path}")

        self.btn_transfer.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.log_area.clear()
        self.log(f"Начало обработки {len(file_paths)} файл(ов)...")

        self.process_thread = ProcessThread(file_paths, excel_path)
        self.process_thread.progress.connect(self._on_progress)
        self.process_thread.finished_signal.connect(self._on_finished)
        self.process_thread.start()

    def _on_progress(self, percent: int, message: str) -> None:
        self.progress_bar.setValue(percent)
        self.status_label.setText(message)

    def _on_finished(
        self, success: bool, message: str, errors: list[str]
    ) -> None:
        self.progress_bar.setValue(100)
        self.btn_transfer.setEnabled(True)

        if errors:
            self.log("\nПредупреждения:")
            for err in errors:
                self.log(f"  - {err}")

        if success:
            self.log(f"\n{message}")
            self.status_label.setText("Готово!")
            QMessageBox.information(self, "Готово", message)
        else:
            self.log(f"\nОшибка: {message}")
            self.status_label.setText("Ошибка!")
            QMessageBox.critical(self, "Ошибка", message)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    font = QFont("Segoe UI", 10)
    app.setFont(font)

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

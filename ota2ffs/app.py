from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .converter import convert_excel


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("OTA2FFS Converter")
        self.resize(860, 620)
        self.excel_path: Path | None = None
        self.output_dir: Path | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        central = QWidget()
        root = QVBoxLayout(central)

        file_row = QHBoxLayout()
        self.excel_edit = QLineEdit()
        self.excel_edit.setReadOnly(True)
        choose_excel_button = QPushButton("选择 Excel")
        choose_excel_button.clicked.connect(self.choose_excel)
        file_row.addWidget(QLabel("Excel 文件"))
        file_row.addWidget(self.excel_edit, 1)
        file_row.addWidget(choose_excel_button)

        output_row = QHBoxLayout()
        self.output_edit = QLineEdit()
        self.output_edit.setReadOnly(True)
        choose_output_button = QPushButton("选择输出目录")
        choose_output_button.clicked.connect(self.choose_output_dir)
        output_row.addWidget(QLabel("输出目录"))
        output_row.addWidget(self.output_edit, 1)
        output_row.addWidget(choose_output_button)

        sheet_buttons = QHBoxLayout()
        select_all_button = QPushButton("全选")
        clear_all_button = QPushButton("取消全选")
        select_all_button.clicked.connect(lambda: self.set_all_sheets_checked(True))
        clear_all_button.clicked.connect(lambda: self.set_all_sheets_checked(False))
        sheet_buttons.addWidget(QLabel("Sheet 列表"))
        sheet_buttons.addStretch(1)
        sheet_buttons.addWidget(select_all_button)
        sheet_buttons.addWidget(clear_all_button)

        self.sheet_list = QListWidget()
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)

        self.convert_button = QPushButton("开始转换")
        self.convert_button.clicked.connect(self.start_conversion)

        root.addLayout(file_row)
        root.addLayout(output_row)
        root.addLayout(sheet_buttons)
        root.addWidget(self.sheet_list, 2)
        root.addWidget(self.convert_button)
        root.addWidget(QLabel("转换日志"))
        root.addWidget(self.log_output, 3)

        self.setCentralWidget(central)

    def choose_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Excel 文件",
            "",
            "Excel Files (*.xlsx *.xlsm);;All Files (*)",
        )
        if not path:
            return
        self.excel_path = Path(path)
        self.excel_edit.setText(str(self.excel_path))
        self.load_sheet_names()

    def choose_output_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if not path:
            return
        self.output_dir = Path(path)
        self.output_edit.setText(str(self.output_dir))

    def load_sheet_names(self) -> None:
        self.sheet_list.clear()
        if self.excel_path is None:
            return
        try:
            workbook = load_workbook(self.excel_path, read_only=True, data_only=True)
            for sheet_name in workbook.sheetnames:
                item = QListWidgetItem(sheet_name)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked)
                self.sheet_list.addItem(item)
            workbook.close()
            self.append_log(f"已读取 {self.sheet_list.count()} 个 sheet")
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", str(exc))
            self.append_log(f"[失败] 读取 sheet 列表: {exc}")

    def set_all_sheets_checked(self, checked: bool) -> None:
        state = Qt.Checked if checked else Qt.Unchecked
        for index in range(self.sheet_list.count()):
            self.sheet_list.item(index).setCheckState(state)

    def selected_sheets(self) -> list[str]:
        sheets: list[str] = []
        for index in range(self.sheet_list.count()):
            item = self.sheet_list.item(index)
            if item.checkState() == Qt.Checked:
                sheets.append(item.text())
        return sheets

    def start_conversion(self) -> None:
        if self.excel_path is None:
            QMessageBox.warning(self, "缺少 Excel 文件", "请先选择 Excel 文件。")
            return
        if self.output_dir is None:
            QMessageBox.warning(self, "缺少输出目录", "请先选择输出目录。")
            return

        sheets = self.selected_sheets()
        if not sheets:
            QMessageBox.warning(self, "缺少 Sheet", "请至少选择一个 sheet。")
            return

        self.convert_button.setEnabled(False)
        self.append_log("开始转换...")
        try:
            result = convert_excel(self.excel_path, self.output_dir, sheets)
            for line in result.log_lines:
                self.append_log(line)
            if result.log_path is not None:
                self.append_log(f"日志文件: {result.log_path}")
            QMessageBox.information(self, "转换完成", f"转换完成，生成 {result.generated_count} 个文件。")
        finally:
            self.convert_button.setEnabled(True)

    def append_log(self, message: str) -> None:
        self.log_output.append(message)


def main() -> None:
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()

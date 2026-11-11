from __future__ import annotations
import json, sys, traceback
from pathlib import Path
from datetime import datetime
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QPushButton, QLineEdit, QFileDialog, QComboBox, QTableWidget,
    QTableWidgetItem, QMessageBox, QGroupBox, QProgressBar, QAbstractItemView, QHeaderView)
from engine import load_sheet, detect_placeholders, replace_placeholders, safe_filename, unique_output_path, normalize_key

APP_NAME = "Excel to Word Autofill Automation"

class GenerationWorker(QThread):
    progress = Signal(int, str)
    finished_result = Signal(dict)
    def __init__(self, template, output_dir, headers, rows, selected_rows, mappings, filename_column):
        super().__init__()
        self.template, self.output_dir = template, Path(output_dir)
        self.headers, self.rows = headers, rows
        self.selected_rows, self.mappings = selected_rows, mappings
        self.filename_column = filename_column
    def run(self):
        successes, failures, warnings = [], [], []
        total = len(self.selected_rows)
        for i, row in enumerate(self.selected_rows, 1):
            name_value = row.get(self.filename_column) if self.filename_column else None
            base = safe_filename(name_value, f"Record_{i}")
            out = unique_output_path(self.output_dir, base)
            self.progress.emit(round((i-1)/max(total, 1)*100), f"Generating {i} of {total}: {base}")
            try:
                result = replace_placeholders(self.template, out, row, self.mappings)
                successes.append(str(out))
                if result["unresolved_placeholders"]:
                    warnings.append(f"{out.name}: unresolved placeholders: {', '.join(result['unresolved_placeholders'])}")
            except Exception as e:
                failures.append({"record": base, "error": str(e)})
                # Remove a partial file if a failed save left one behind.
                try:
                    if out.exists(): out.unlink()
                except Exception: pass
        self.progress.emit(100, "Generation finished")
        self.finished_result.emit({"successes": successes, "failures": failures, "warnings": warnings})

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1120, 820)
        self.template_path = ""
        self.excel_path = ""
        self.output_dir = str(Path.home() / "Documents" / "Generated Word Documents")
        self.headers, self.rows, self.placeholders = [], [], []
        self.mapping_widgets = {}
        self.worker = None
        self._build_ui()
        self._load_last_config()

    def _build_ui(self):
        central = QWidget(); self.setCentralWidget(central)
        root = QVBoxLayout(central); root.setSpacing(12)
        heading = QLabel(APP_NAME); heading.setStyleSheet("font-size: 23px; font-weight: 700;")
        subtitle = QLabel("Create Word documents from selected Excel rows using a reusable Word template.")
        subtitle.setWordWrap(True); subtitle.setStyleSheet("color: #555; font-size: 12px;")
        root.addWidget(heading); root.addWidget(subtitle)

        files = QGroupBox("1. Choose your files"); grid = QGridLayout(files)
        self.template_edit = QLineEdit(); self.template_edit.setReadOnly(True)
        self.excel_edit = QLineEdit(); self.excel_edit.setReadOnly(True)
        self.sheet_combo = QComboBox(); self.sheet_combo.currentTextChanged.connect(self._load_selected_sheet)
        self.output_edit = QLineEdit(self.output_dir); self.output_edit.setReadOnly(True)
        btn_template = QPushButton("Choose Word Template…"); btn_template.clicked.connect(self.choose_template)
        btn_excel = QPushButton("Choose Excel Workbook…"); btn_excel.clicked.connect(self.choose_excel)
        btn_output = QPushButton("Choose Output Folder…"); btn_output.clicked.connect(self.choose_output)
        grid.addWidget(QLabel("Word template (.docx)"), 0, 0); grid.addWidget(self.template_edit, 0, 1); grid.addWidget(btn_template, 0, 2)
        grid.addWidget(QLabel("Excel workbook (.xlsx)"), 1, 0); grid.addWidget(self.excel_edit, 1, 1); grid.addWidget(btn_excel, 1, 2)
        grid.addWidget(QLabel("Worksheet"), 2, 0); grid.addWidget(self.sheet_combo, 2, 1, 1, 2)
        grid.addWidget(QLabel("Output folder"), 3, 0); grid.addWidget(self.output_edit, 3, 1); grid.addWidget(btn_output, 3, 2)
        grid.setColumnStretch(1, 1); root.addWidget(files)

        mapping_box = QGroupBox("2. Review placeholder mappings")
        mapping_layout = QVBoxLayout(mapping_box)
        note = QLabel("In Word, add markers such as {{Full Name}} where Excel data should appear. Matching Excel columns are selected automatically.")
        note.setWordWrap(True); mapping_layout.addWidget(note)
        self.mapping_table = QTableWidget(0, 3); self.mapping_table.setHorizontalHeaderLabels(["Word Placeholder", "Excel Column", "Status"])
        self.mapping_table.setAlternatingRowColors(True); self.mapping_table.verticalHeader().setVisible(False)
        self.mapping_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.mapping_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.mapping_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        mapping_layout.addWidget(self.mapping_table)
        rowbtn = QHBoxLayout(); refresh_btn = QPushButton("Rescan Word Template"); refresh_btn.clicked.connect(self.scan_template)
        save_btn = QPushButton("Save Project Configuration…"); save_btn.clicked.connect(self.save_config)
        load_btn = QPushButton("Load Project Configuration…"); load_btn.clicked.connect(self.load_config)
        rowbtn.addWidget(refresh_btn); rowbtn.addStretch(1); rowbtn.addWidget(load_btn); rowbtn.addWidget(save_btn)
        mapping_layout.addLayout(rowbtn); root.addWidget(mapping_box, 2)

        select_box = QGroupBox("3. Select Excel rows to process")
        sl = QVBoxLayout(select_box)
        selection_note = QLabel(
            "Check the rows you want to turn into Word documents. You can select any "
            "combination of rows, including non-consecutive rows."
        )
        selection_note.setWordWrap(True)
        sl.addWidget(selection_note)

        selection_buttons = QHBoxLayout()
        self.select_all_btn = QPushButton("Select All Rows")
        self.select_all_btn.clicked.connect(lambda: self._set_all_rows_checked(True))
        self.deselect_all_btn = QPushButton("Deselect All Rows")
        self.deselect_all_btn.clicked.connect(lambda: self._set_all_rows_checked(False))
        selection_buttons.addWidget(self.select_all_btn)
        selection_buttons.addWidget(self.deselect_all_btn)
        selection_buttons.addStretch(1)
        sl.addLayout(selection_buttons)

        self.rows_table = QTableWidget(0, 2)
        self.rows_table.setHorizontalHeaderLabels(["Process", "Excel Row"])
        self.rows_table.setAlternatingRowColors(True)
        self.rows_table.verticalHeader().setVisible(False)
        self.rows_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.rows_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.rows_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.rows_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.rows_table.itemChanged.connect(self._update_row_selection_count)
        sl.addWidget(self.rows_table, 1)

        self.row_count_label = QLabel("Load an Excel workbook to view its rows.")
        sl.addWidget(self.row_count_label)
        root.addWidget(select_box, 2)

        bottom = QHBoxLayout(); self.status_label = QLabel("Ready"); self.status_label.setWordWrap(True)
        self.progress = QProgressBar(); self.progress.setValue(0); self.progress.setMaximumWidth(220)
        self.generate_btn = QPushButton("Generate Word Documents"); self.generate_btn.setMinimumHeight(42); self.generate_btn.setStyleSheet("font-weight: 700; padding: 8px 18px;")
        self.generate_btn.clicked.connect(self.generate)
        bottom.addWidget(self.status_label, 1); bottom.addWidget(self.progress); bottom.addWidget(self.generate_btn); root.addLayout(bottom)

        self.log_btn = QPushButton("View Latest Generation Log"); self.log_btn.clicked.connect(self.view_log); self.log_btn.setEnabled(False)
        open_btn = QPushButton("Open Output Folder"); open_btn.clicked.connect(self.open_output)
        bottom2 = QHBoxLayout(); bottom2.addWidget(self.log_btn); bottom2.addStretch(1); bottom2.addWidget(open_btn); root.addLayout(bottom2)
        self.setStyleSheet("QGroupBox { font-weight: 600; margin-top: 8px; } QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 4px; } QPushButton { padding: 5px 10px; }")

    def choose_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose Word Template", "", "Word documents (*.docx)")
        if path:
            self.template_path = path; self.template_edit.setText(path); self.scan_template(); self._save_last_config()

    def choose_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose Excel Master List", "", "Excel workbooks (*.xlsx)")
        if not path: return
        self.excel_path = path; self.excel_edit.setText(path)
        try:
            self.sheet_combo.blockSignals(True); self.sheet_combo.clear()
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
            self.sheet_combo.addItems(wb.sheetnames); wb.close()
            self.sheet_combo.blockSignals(False)
            if self.sheet_combo.count(): self._load_selected_sheet(self.sheet_combo.currentText())
            self._save_last_config()
        except Exception as e:
            self.sheet_combo.blockSignals(False); self.show_error(f"Could not read the Excel workbook:\n{e}")

    def choose_output(self):
        path = QFileDialog.getExistingDirectory(self, "Choose Output Folder", self.output_dir)
        if path:
            self.output_dir = path; self.output_edit.setText(path); self._save_last_config()

    def _load_selected_sheet(self, sheet):
        if not self.excel_path or not sheet:
            return
        try:
            self.headers, self.rows, warnings = load_sheet(self.excel_path, sheet)
            self._populate_mapping_table()
            self._populate_rows_table()
            msg = f"Loaded {len(self.rows)} Excel row(s)."
            if warnings:
                msg += " " + " ".join(warnings)
            self.status_label.setText(msg)
        except Exception as e:
            self.show_error(f"Could not load worksheet:\\n{e}")

    def _populate_rows_table(self):
        self.rows_table.blockSignals(True)
        self.rows_table.clear()
        self.rows_table.setColumnCount(2 + len(self.headers))
        self.rows_table.setHorizontalHeaderLabels(["Process", "Excel Row"] + self.headers)
        self.rows_table.setRowCount(len(self.rows))

        for table_row, record in enumerate(self.rows):
            check_item = QTableWidgetItem()
            check_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsSelectable)
            check_item.setCheckState(Qt.Unchecked)
            check_item.setData(Qt.UserRole, table_row)
            self.rows_table.setItem(table_row, 0, check_item)

            excel_row = QTableWidgetItem(str(record.get("__excel_row__", "")))
            excel_row.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.rows_table.setItem(table_row, 1, excel_row)

            for col_index, header in enumerate(self.headers, start=2):
                value = record.get(header)
                display = "" if value is None else str(value)
                cell = QTableWidgetItem(display)
                cell.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                self.rows_table.setItem(table_row, col_index, cell)

        self.rows_table.resizeColumnsToContents()
        self.rows_table.horizontalHeader().setStretchLastSection(False)
        self.rows_table.blockSignals(False)
        self._update_row_selection_count()

    def _set_all_rows_checked(self, checked):
        state = Qt.Checked if checked else Qt.Unchecked
        self.rows_table.blockSignals(True)
        for row_index in range(self.rows_table.rowCount()):
            item = self.rows_table.item(row_index, 0)
            if item:
                item.setCheckState(state)
        self.rows_table.blockSignals(False)
        self._update_row_selection_count()

    def _update_row_selection_count(self, *_):
        if not hasattr(self, "rows_table"):
            return
        selected = sum(
            1 for row_index in range(self.rows_table.rowCount())
            if self.rows_table.item(row_index, 0)
            and self.rows_table.item(row_index, 0).checkState() == Qt.Checked
        )
        total = self.rows_table.rowCount()
        self.row_count_label.setText(f"{selected} of {total} row(s) selected.")

    def _selected_rows(self):
        selected = []
        for table_row in range(self.rows_table.rowCount()):
            item = self.rows_table.item(table_row, 0)
            if item and item.checkState() == Qt.Checked:
                row_index = item.data(Qt.UserRole)
                if row_index is not None and 0 <= row_index < len(self.rows):
                    selected.append(self.rows[row_index])
        return selected

    def scan_template(self):
        if not self.template_path: return
        try:
            self.placeholders = detect_placeholders(self.template_path)
            self._populate_mapping_table()
            self.status_label.setText(f"Found {len(self.placeholders)} placeholder(s) in the Word template.")
        except Exception as e: self.show_error(f"Could not scan Word template:\n{e}")

    def _populate_mapping_table(self):
        self.mapping_table.setRowCount(0); self.mapping_widgets = {}
        self.mapping_table.setRowCount(len(self.placeholders))
        options = ["— Not mapped —"] + self.headers
        header_by_norm = {}
        for h in self.headers: header_by_norm.setdefault(normalize_key(h), h)
        for i, placeholder in enumerate(self.placeholders):
            self.mapping_table.setItem(i, 0, QTableWidgetItem("{{" + placeholder + "}}"))
            combo = QComboBox(); combo.addItems(options)
            auto = header_by_norm.get(normalize_key(placeholder))
            if auto: combo.setCurrentText(auto)
            combo.currentTextChanged.connect(lambda _v, row=i: self._update_mapping_status(row))
            self.mapping_table.setCellWidget(i, 1, combo); self.mapping_widgets[placeholder] = combo
            self.mapping_table.setItem(i, 2, QTableWidgetItem("Matched" if auto else "Not mapped"))
            if not auto: self.mapping_table.item(i, 2).setForeground(Qt.darkRed)
        self.mapping_table.resizeRowsToContents()

    def _update_mapping_status(self, row):
        combo = self.mapping_table.cellWidget(row, 1)
        item = self.mapping_table.item(row, 2)
        if combo and item:
            mapped = combo.currentText() != "— Not mapped —"
            item.setText("Matched" if mapped else "Not mapped")
            item.setForeground(Qt.darkGreen if mapped else Qt.darkRed)

    def _get_mappings(self):
        result = {}
        for placeholder, combo in self.mapping_widgets.items():
            col = combo.currentText()
            if col != "— Not mapped —": result[placeholder] = col
        return result

    def generate(self):
        if not self.template_path or not Path(self.template_path).is_file(): return self.show_error("Please choose a valid Word template.")
        if not self.excel_path or not self.rows: return self.show_error("Please choose an Excel workbook containing data rows.")
        if not self.output_dir: return self.show_error("Please choose an output folder.")
        mappings = self._get_mappings()
        unmapped = [p for p in self.placeholders if p not in mappings]
        if unmapped:
            return self.show_error("Map these placeholders before generating documents:\n\n" + "\n".join("• {{"+p+"}}" for p in unmapped))
        selected = self._selected_rows()
        if not selected: return self.show_error("No Excel rows are selected. Check at least one row in the table.")
        try: Path(self.output_dir).mkdir(parents=True, exist_ok=True)
        except Exception as e: return self.show_error(f"Cannot create or access output folder:\n{e}")
        filename_col = next((h for h in self.headers if normalize_key(h) in ("full name", "student name", "name")), self.headers[0] if self.headers else None)
        self.generate_btn.setEnabled(False); self.progress.setValue(0); self.log_btn.setEnabled(False)
        self.worker = GenerationWorker(self.template_path, self.output_dir, self.headers, self.rows, selected, mappings, filename_col)
        self.worker.progress.connect(self._generation_progress); self.worker.finished_result.connect(self._generation_finished); self.worker.start()

    def _generation_progress(self, value, message): self.progress.setValue(value); self.status_label.setText(message)

    def _generation_finished(self, result):
        self.generate_btn.setEnabled(True); self.log_btn.setEnabled(True)
        log = {"timestamp": datetime.now().isoformat(timespec="seconds"), "template": self.template_path,
               "excel": self.excel_path, "worksheet": self.sheet_combo.currentText(),
               "success_count": len(result["successes"]), "failures": result["failures"],
               "warnings": result["warnings"], "files": result["successes"]}
        self.last_log = log
        try:
            Path(self.output_dir, "generation_log.json").write_text(json.dumps(log, indent=2, ensure_ascii=False), encoding="utf-8")
        except Exception: pass
        self.status_label.setText(f"Finished: {len(result['successes'])} created, {len(result['failures'])} failed, {len(result['warnings'])} warning(s).")
        self._save_last_config()
        QMessageBox.information(self, "Generation complete", f"Successfully created: {len(result['successes'])}\nFailed: {len(result['failures'])}\nWarnings: {len(result['warnings'])}\n\nOutput folder:\n{self.output_dir}")

    def view_log(self):
        path = Path(self.output_dir) / "generation_log.json"
        if path.exists():
            from PySide6.QtGui import QDesktopServices
            from PySide6.QtCore import QUrl
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
        elif hasattr(self, "last_log"):
            QMessageBox.information(self, "Generation Log", json.dumps(self.last_log, indent=2, ensure_ascii=False))
        else: QMessageBox.information(self, "Generation Log", "No generation log is available yet.")

    def open_output(self):
        Path(self.output_dir).mkdir(parents=True, exist_ok=True)
        from PySide6.QtGui import QDesktopServices
        from PySide6.QtCore import QUrl
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(self.output_dir).resolve())))

    def save_config(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Project Configuration", "word_excel_project.json", "JSON configuration (*.json)")
        if not path: return
        config = self._config_data()
        try: Path(path).write_text(json.dumps(config, indent=2, ensure_ascii=False), encoding="utf-8")
        except Exception as e: self.show_error(f"Could not save configuration:\n{e}"); return
        self.status_label.setText("Project configuration saved.")

    def load_config(self):
        path, _ = QFileDialog.getOpenFileName(self, "Load Project Configuration", "", "JSON configuration (*.json)")
        if not path: return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            self._apply_config(data)
            self.status_label.setText("Configuration loaded. Review mappings before generating.")
        except Exception as e: self.show_error(f"Could not load configuration:\n{e}")

    def _config_data(self):
        return {"template_path": self.template_path, "excel_path": self.excel_path, "worksheet": self.sheet_combo.currentText(),
                "output_dir": self.output_dir, "mappings": self._get_mappings()}

    def _apply_config(self, data):
        tp = data.get("template_path", ""); ep = data.get("excel_path", "")
        if tp and Path(tp).is_file(): self.template_path = tp; self.template_edit.setText(tp); self.scan_template()
        if ep and Path(ep).is_file():
            self.excel_path = ep; self.excel_edit.setText(ep)
            from openpyxl import load_workbook
            wb = load_workbook(ep, read_only=True, data_only=True); sheets = wb.sheetnames; wb.close()
            self.sheet_combo.blockSignals(True); self.sheet_combo.clear(); self.sheet_combo.addItems(sheets)
            wanted = data.get("worksheet", "")
            if wanted in sheets: self.sheet_combo.setCurrentText(wanted)
            self.sheet_combo.blockSignals(False); self._load_selected_sheet(self.sheet_combo.currentText())
        if data.get("output_dir"):
            self.output_dir = data["output_dir"]; self.output_edit.setText(self.output_dir)
        mappings = data.get("mappings", {})
        for p, col in mappings.items():
            if p in self.mapping_widgets and col in self.headers: self.mapping_widgets[p].setCurrentText(col)
        self._save_last_config()

    def _last_config_path(self): return Path.home() / ".excel_to_word_autofill_automation.json"
    def _save_last_config(self):
        try: self._last_config_path().write_text(json.dumps(self._config_data(), indent=2), encoding="utf-8")
        except Exception: pass
    def _load_last_config(self):
        path = self._last_config_path()
        if path.exists():
            try: self._apply_config(json.loads(path.read_text(encoding="utf-8")))
            except Exception: pass
    def show_error(self, message): QMessageBox.warning(self, "Please check your selections", message)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    window = MainWindow(); window.show()
    sys.exit(app.exec())

if __name__ == "__main__": main()

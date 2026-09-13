import sys
import time
from pathlib import Path

import pythoncom
import win32com.client

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QCompleter,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from cad_analysis import Detector, DEFAULT_CLEARANCE_M


SUPPORTED_SOURCE = "Drainage Pipe E"
SUPPORTED_TARGETS = {"AG_t", "AG_tttt"}


def com_call(fn, retries=8, delay=0.15):
    last = None
    for _ in range(retries):
        try:
            return fn()
        except Exception as exc:
            last = exc
            if "Call was rejected by callee" not in str(exc):
                raise
            pythoncom.PumpWaitingMessages()
            time.sleep(delay)
    raise last


def active_autocad():
    return win32com.client.GetActiveObject("AutoCAD.Application")


def block_names(doc):
    names = set()
    blocks = com_call(lambda: doc.Blocks)
    count = int(com_call(lambda: blocks.Count))
    for i in range(count):
        try:
            block = com_call(lambda i=i: blocks.Item(i))
            name = str(block.Name).strip()
            if name and not name.startswith("*"):
                names.add(name)
        except Exception:
            continue
    return sorted(names, key=str.casefold)


def layout_names(doc):
    names = ["ModelSpace"]
    layouts = com_call(lambda: doc.Layouts)
    count = int(com_call(lambda: layouts.Count))
    for i in range(count):
        try:
            layout = com_call(lambda i=i: layouts.Item(i))
            name = str(layout.Name)
            if name not in names:
                names.append(name)
        except Exception:
            continue
    return names


class SearchCombo(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)
        self.setMinimumHeight(34)
        self.completer_obj = QCompleter(self.model(), self)
        self.completer_obj.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer_obj.setFilterMode(Qt.MatchContains)
        self.completer_obj.setCompletionMode(QCompleter.PopupCompletion)
        self.setCompleter(self.completer_obj)

    def set_items(self, items):
        current = self.currentText()
        self.blockSignals(True)
        self.clear()
        self.addItems(items)
        self.blockSignals(False)
        self.setCurrentText(current if current in items else (items[0] if items else ""))
        self.completer_obj.setModel(self.model())


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.doc = None
        self.last_doc_name = ""
        self.blocks = []
        self.layouts = []

        self.setWindowTitle("CAD Conflict Corrector")
        self.setMinimumSize(680, 620)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)

        self.build_ui()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.watch_autocad)
        self.timer.start(1500)

    def build_ui(self):
        root = QWidget()
        main = QVBoxLayout(root)
        main.setContentsMargins(16, 16, 16, 16)
        main.setSpacing(12)

        title = QLabel("CAD Conflict Corrector")
        title.setStyleSheet("font-size: 22px; font-weight: 700;")
        main.addWidget(title)

        self.status = QLabel("AutoCAD: not connected")
        self.status.setFrameStyle(QFrame.StyledPanel | QFrame.Sunken)
        main.addWidget(self.status)

        top = QHBoxLayout()
        self.load_btn = QPushButton("Load Active Drawing")
        self.load_btn.clicked.connect(self.load_document)
        top.addWidget(self.load_btn)

        self.refresh_btn = QPushButton("Refresh Blocks / Layouts")
        self.refresh_btn.clicked.connect(self.refresh_lists)
        top.addWidget(self.refresh_btn)
        main.addLayout(top)

        selection = QGroupBox("Conflict rule")
        form = QFormLayout(selection)

        self.block_a = SearchCombo()
        self.block_b = SearchCombo()
        self.layout_combo = SearchCombo()

        form.addRow("Block A:", self.block_a)
        form.addRow("Block B:", self.block_b)
        form.addRow("Layout:", self.layout_combo)

        main.addWidget(selection)

        info = QLabel(
            "Choose which two block types may conflict. "
            "Live search works by typing in either block selector."
        )
        info.setWordWrap(True)
        main.addWidget(info)

        params = QGroupBox("Correction")
        pform = QFormLayout(params)

        self.clearance = QSpinBox()
        self.clearance.setRange(0, 1000)
        self.clearance.setValue(50)
        self.clearance.setSuffix(" mm")
        pform.addRow("Clearance:", self.clearance)

        self.save_check = QCheckBox("Save DWG after correction")
        self.save_check.setChecked(False)
        pform.addRow("", self.save_check)

        main.addWidget(params)

        rules = QGroupBox("Configured rule")
        rules_layout = QVBoxLayout(rules)
        self.rule_list = QListWidget()
        self.rule_list.setMaximumHeight(100)
        rules_layout.addWidget(self.rule_list)

        add_rule = QPushButton("Add selected pair")
        add_rule.clicked.connect(self.add_rule)
        rules_layout.addWidget(add_rule)
        main.addWidget(rules)

        actions = QHBoxLayout()
        self.analyze_btn = QPushButton("Analyze")
        self.analyze_btn.clicked.connect(self.analyze)
        actions.addWidget(self.analyze_btn)

        self.apply_btn = QPushButton("APPLY CORRECTION")
        self.apply_btn.clicked.connect(self.apply_correction)
        self.apply_btn.setMinimumHeight(44)
        actions.addWidget(self.apply_btn)
        main.addLayout(actions)

        self.log = QListWidget()
        main.addWidget(self.log, 1)

        self.setCentralWidget(root)

        self.block_a.currentTextChanged.connect(self.update_supported_state)
        self.block_b.currentTextChanged.connect(self.update_supported_state)
        self.update_supported_state()

    def log_msg(self, text):
        self.log.addItem(text)
        self.log.scrollToBottom()

    def load_document(self):
        try:
            self.doc = active_autocad().ActiveDocument
            self.last_doc_name = str(self.doc.Name)
            self.status.setText(f"AutoCAD: {self.last_doc_name}")
            self.refresh_lists()
            self.log_msg(f"Loaded: {self.last_doc_name}")
        except Exception as exc:
            self.doc = None
            self.status.setText("AutoCAD: connection failed")
            QMessageBox.critical(self, "AutoCAD", str(exc))

    def refresh_lists(self):
        if self.doc is None:
            self.load_document()
            if self.doc is None:
                return
        try:
            self.blocks = block_names(self.doc)
            self.layouts = layout_names(self.doc)

            self.block_a.set_items(self.blocks)
            self.block_b.set_items(self.blocks)
            self.layout_combo.set_items(self.layouts)

            if SUPPORTED_SOURCE in self.blocks:
                self.block_a.setCurrentText(SUPPORTED_SOURCE)
            for target in ("AG_tttt", "AG_t"):
                if target in self.blocks:
                    self.block_b.setCurrentText(target)
                    break

            self.layout_combo.setCurrentText("ModelSpace")
            self.log_msg(
                f"Loaded {len(self.blocks)} block definitions and "
                f"{len(self.layouts)} layouts."
            )
            self.update_supported_state()
        except Exception as exc:
            QMessageBox.critical(self, "Refresh failed", str(exc))

    def watch_autocad(self):
        try:
            doc = active_autocad().ActiveDocument
            name = str(doc.Name)
            if name != self.last_doc_name:
                self.doc = doc
                self.last_doc_name = name
                self.status.setText(f"AutoCAD: {name}")
                self.refresh_lists()
        except Exception:
            pass

    def current_pair_supported(self):
        return (
            self.block_a.currentText() == SUPPORTED_SOURCE
            and self.block_b.currentText() in SUPPORTED_TARGETS
            and self.layout_combo.currentText() == "ModelSpace"
        )

    def update_supported_state(self):
        supported = self.current_pair_supported()
        self.analyze_btn.setEnabled(supported)
        self.apply_btn.setEnabled(supported)
        if supported:
            self.status.setText(
                f"AutoCAD: {self.last_doc_name} | Supported rule selected"
            )
        elif self.doc:
            self.status.setText(
                f"AutoCAD: {self.last_doc_name} | Select Drainage Pipe E + AG_t/AG_tttt + ModelSpace"
            )

    def add_rule(self):
        a = self.block_a.currentText().strip()
        b = self.block_b.currentText().strip()
        layout = self.layout_combo.currentText().strip()
        if not a or not b:
            return
        text = f"{layout}: {a}  <->  {b}"
        for i in range(self.rule_list.count()):
            if self.rule_list.item(i).text() == text:
                return
        self.rule_list.addItem(QListWidgetItem(text))
        self.log_msg(f"Rule added: {text}")

    def run_analysis(self):
        if self.doc is None:
            self.load_document()
        if self.doc is None:
            return None
        if not self.current_pair_supported():
            QMessageBox.warning(
                self,
                "Unsupported rule",
                "The current correction engine supports only "
                "Drainage Pipe E against AG_t / AG_tttt in ModelSpace.",
            )
            return None

        detector = Detector(
            self.doc,
            clearance_m=self.clearance.value() / 1000.0,
        )
        results = detector.analyze_all()
        return detector, results

    def analyze(self):
        try:
            result = self.run_analysis()
            if result is None:
                return
            _, results = result
            blocked = sum(r.status == "BLOCKED" for r in results)
            moves = sum(r.required_move_mm > 0 for r in results)
            self.log_msg(
                f"Analysis complete: {len(results)} drainages, "
                f"{blocked} blocked, {moves} require movement."
            )
            for r in results:
                if r.required_move_mm > 0:
                    self.log_msg(
                        f"{r.drainage}: move {r.required_move_mm:.1f} mm "
                        f"dir=({r.move_dir_x:+.4f},{r.move_dir_y:+.4f})"
                    )
        except Exception as exc:
            QMessageBox.critical(self, "Analysis failed", str(exc))

    def apply_correction(self):
        answer = QMessageBox.question(
            self,
            "Confirm correction",
            "The original drainage blocks will be MOVED in AutoCAD.\n\n"
            "Sleeper blocks will not be moved.\n"
            "Movement marker lines will remain.\n\n"
            "Continue?",
        )
        if answer != QMessageBox.Yes:
            return

        try:
            result = self.run_analysis()
            if result is None:
                return
            detector, results = result
            moved = detector.apply_moves(
                results,
                save=self.save_check.isChecked(),
            )
            self.log_msg(f"CORRECTION COMPLETE: {moved} drainage objects moved.")
            QMessageBox.information(
                self,
                "Complete",
                f"Correction complete.\n\nMoved: {moved}",
            )
        except Exception as exc:
            QMessageBox.critical(self, "Correction failed", str(exc))


def main():
    pythoncom.CoInitialize()
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    window.load_document()
    exit_code = app.exec()
    pythoncom.CoUninitialize()
    sys.exit(exit_code)


if __name__ == "__main__":
    main()

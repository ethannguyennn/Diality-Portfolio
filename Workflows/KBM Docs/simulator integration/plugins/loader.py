import os
import sys

import pandas as pd

from PySide2.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QRadioButton, QButtonGroup,
    QTextEdit, QProgressBar, QStackedWidget, QTreeWidget, QTreeWidgetItem,
    QMessageBox, QCheckBox,
)
from PySide2.QtCore import Qt, QThread, Signal
from PySide2.QtGui import QFont

import core.constants as C
from core.scanning import normalize, extract_dia_code, get_id_from_json, update_excel
from core.processing import process_documents


# ── Stdout capture ─────────────────────────────────────────────────────────────


class LogStream:
    def __init__(self, signal):
        self._signal = signal

    def write(self, text):
        if text and text.strip():
            self._signal.emit(text.rstrip('\n'))

    def flush(self):
        pass


# ── Background workers ─────────────────────────────────────────────────────────


class ScanWorker(QThread):
    log_signal      = Signal(str)
    finished_signal = Signal(object, object)
    error_signal    = Signal(str)

    def __init__(self, excel_path):
        super().__init__()
        self._excel_path = excel_path

    def run(self):
        old_stdout = sys.stdout
        sys.stdout = LogStream(self.log_signal)
        try:
            missing_doc_ids = update_excel(self._excel_path)
            docDF = pd.read_excel(self._excel_path)
            self.finished_signal.emit(docDF, missing_doc_ids)
        except Exception as e:
            self.error_signal.emit(str(e))
        finally:
            sys.stdout = old_stdout


class ProcessWorker(QThread):
    log_signal      = Signal(str)
    progress_signal = Signal(int, int)
    finished_signal = Signal(bool, str)

    def __init__(self, docDF, mode, doc_indices):
        super().__init__()
        self._docDF       = docDF
        self._mode        = mode
        self._doc_indices = doc_indices

    def run(self):
        old_stdout = sys.stdout
        sys.stdout = LogStream(self.log_signal)
        try:
            process_documents(
                self._docDF,
                self._mode,
                self._doc_indices,
                progress_callback=lambda cur, total: self.progress_signal.emit(cur, total)
            )
            self.finished_signal.emit(True, "All operations completed successfully.")
        except Exception as e:
            self.finished_signal.emit(False, f"Error: {e}")
        finally:
            sys.stdout = old_stdout


# ── Page 0: Folder Selection ───────────────────────────────────────────────────


class FolderPage(QWidget):
    folder_changed = Signal(str)

    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(16, 16, 16, 16)

        title = QLabel("Select KBM Docs Folder")
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        layout.addWidget(title)

        desc = QLabel("Choose the folder containing your PDF and DOCX files.")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        self.folder_label = QLabel("No folder selected")
        self.folder_label.setStyleSheet("color: #666; font-style: italic;")
        self.folder_label.setWordWrap(True)
        layout.addWidget(self.folder_label)

        browse_btn = QPushButton("Browse...")
        browse_btn.setFixedWidth(120)
        browse_btn.clicked.connect(self._browse)
        layout.addWidget(browse_btn)

        outer.addLayout(layout)
        outer.addStretch()

        self.selected_folder = ""

    def _browse(self):
        folder = QFileDialog.getExistingDirectory(self, "Select KBM Docs Folder")
        if folder:
            self.selected_folder = folder
            self.folder_label.setText(folder)
            self.folder_label.setStyleSheet("")
            self.folder_changed.emit(folder)

    def reset(self):
        self.selected_folder = ""
        self.folder_label.setText("No folder selected")
        self.folder_label.setStyleSheet("color: #666; font-style: italic;")


# ── Page 1: Scan Results / Summary ────────────────────────────────────────────


class SummaryPage(QWidget):
    scan_finished = Signal(object, object)  # (docDF, missing_doc_ids)

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        title = QLabel("File Summary")
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        layout.addWidget(title)

        self.status_label = QLabel("Scanning folder...")
        layout.addWidget(self.status_label)

        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        self.summary_text.setFont(QFont("Consolas", 9))
        layout.addWidget(self.summary_text)

        self._last_scanned = None
        self._worker       = None

    def on_enter(self, records_folder, excel_path):
        if records_folder == self._last_scanned:
            return
        if self._worker and self._worker.isRunning():
            return

        self.status_label.setText("Scanning folder...")
        self.summary_text.clear()

        C.RECORDS_FOLDER  = records_folder
        C.OUTDATED_FOLDER = os.path.join(C.RECORDS_FOLDER, 'Outdated')

        self._worker = ScanWorker(excel_path)
        self._worker.log_signal.connect(self._on_log)
        self._worker.finished_signal.connect(self._on_scan_done)
        self._worker.error_signal.connect(self._on_scan_error)
        self._worker.start()

    def _on_log(self, text):
        self.summary_text.append(text)

    def _on_scan_done(self, docDF, missing_doc_ids):
        self._last_scanned = C.RECORDS_FOLDER

        self.summary_text.append("\n" + "=" * 60)
        self.summary_text.append("FILE SUMMARY")
        self.summary_text.append("=" * 60)
        self._build_summary(docDF)

        matched_count = sum(
            1 for x in docDF["DOCX Path"].values
            if pd.notna(x) and str(x).strip()
        )
        missing_count = len(docDF) - matched_count
        self.status_label.setText(
            f"Scan complete: {matched_count} matched, {missing_count} missing"
        )

        self.scan_finished.emit(docDF, missing_doc_ids)

    def _on_scan_error(self, error_msg):
        self.status_label.setText("Scan failed!")
        self.summary_text.append(f"\nERROR: {error_msg}")

    def _build_summary(self, docDF):
        pdf_paths  = docDF["PDF Path"].values
        docx_paths = docDF["DOCX Path"].values
        matched    = []
        missing    = []

        for i in range(len(docDF)):
            pdf_name = os.path.basename(str(pdf_paths[i])) if pd.notna(pdf_paths[i]) else "N/A"
            docx_val = docx_paths[i]
            has_docx = pd.notna(docx_val) and str(docx_val).strip() != ""

            if has_docx:
                matched.append((pdf_name, os.path.basename(str(docx_val))))
            else:
                missing.append(pdf_name)

        if matched:
            self.summary_text.append(f"\nMatched files ({len(matched)}):")
            for pdf_name, docx_name in matched:
                self.summary_text.append(f"\n  {pdf_name}\n  ->  {docx_name}")

        if missing:
            self.summary_text.append(f"\nMissing DOCX ({len(missing)}):")
            for pdf_name in missing:
                self.summary_text.append(f"  [!] {pdf_name}")

        self.summary_text.append(f"\nTotal: {len(matched)} matched, {len(missing)} missing")

    def reset(self):
        self.status_label.setText("Scanning folder...")
        self.summary_text.clear()
        self._last_scanned = None


# ── Page 2: Missing Documents Warning ─────────────────────────────────────────


class MissingDocsPage(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        title = QLabel("Missing Documents")
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        layout.addWidget(title)

        desc = QLabel(
            "The following DOCX files could not be found.\n"
            "You can abort to fix your folder, or continue\n"
            "(these documents will be skipped during processing)."
        )
        desc.setWordWrap(True)
        layout.addWidget(desc)

        self.missing_list = QTextEdit()
        self.missing_list.setReadOnly(True)
        self.missing_list.setFont(QFont("Consolas", 9))
        layout.addWidget(self.missing_list)

    def on_enter(self, missing_doc_ids):
        lines = [f"  {doc_id}  --  {pdf_name}" for doc_id, pdf_name in missing_doc_ids]
        self.missing_list.setText("\n".join(lines))


# ── Page 3: Mode & Document Selection ─────────────────────────────────────────


class ModePage(QWidget):
    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(16, 16, 16, 16)

        title = QLabel("Select Operation")
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        layout.addWidget(title)

        mode_label = QLabel("Operation mode:")
        mode_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        layout.addWidget(mode_label)

        self.mode_group        = QButtonGroup(self)
        self.mode_rev_and_redline = QRadioButton("Create new rev and redline")
        self.mode_rev_only        = QRadioButton("Create only new rev")
        self.mode_redline_only    = QRadioButton("Create only new redline")
        self.mode_rev_and_redline.setChecked(True)
        self.mode_group.addButton(self.mode_rev_and_redline, 1)
        self.mode_group.addButton(self.mode_rev_only,        2)
        self.mode_group.addButton(self.mode_redline_only,    3)

        layout.addWidget(self.mode_rev_and_redline)
        layout.addWidget(self.mode_rev_only)
        layout.addWidget(self.mode_redline_only)

        doc_list_label = QLabel("Documents to process:")
        doc_list_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        layout.addWidget(doc_list_label)

        sel_row = QHBoxLayout()
        self.select_all_cb = QCheckBox("Select All")
        self.select_new_cb = QCheckBox("Select New")
        sel_row.addWidget(self.select_all_cb)
        sel_row.addWidget(self.select_new_cb)
        sel_row.addStretch()
        layout.addLayout(sel_row)

        self.doc_list = QTreeWidget()
        self.doc_list.setHeaderHidden(True)
        self.doc_list.setRootIsDecorated(False)
        self.doc_list.setItemsExpandable(False)
        layout.addWidget(self.doc_list)

        outer.addLayout(layout)
        outer.addStretch()

        self.select_all_cb.stateChanged.connect(self._on_select_all_changed)
        self.select_new_cb.stateChanged.connect(self._on_select_new_changed)

    def _on_select_all_changed(self, state):
        check = Qt.Checked if state == Qt.Checked else Qt.Unchecked
        for i in range(self.doc_list.topLevelItemCount()):
            top = self.doc_list.topLevelItem(i)
            if top.flags() & Qt.ItemIsUserCheckable:
                top.setCheckState(0, check)
            for j in range(top.childCount()):
                child = top.child(j)
                if child.flags() & Qt.ItemIsUserCheckable:
                    child.setCheckState(0, check)

    def _on_select_new_changed(self, state):
        check = Qt.Checked if state == Qt.Checked else Qt.Unchecked
        for i in range(self.doc_list.topLevelItemCount()):
            top = self.doc_list.topLevelItem(i)
            if top.flags() & Qt.ItemIsUserCheckable:
                top.setCheckState(0, check)

    def populate_docs(self, docDF):
        self.select_all_cb.setChecked(False)
        self.select_new_cb.setChecked(False)
        self.doc_list.clear()

        pdf_paths  = docDF["PDF Path"].values
        root_docs  = []
        grouped    = {}

        for i, pdf_val in enumerate(pdf_paths):
            if not (pd.notna(pdf_val) and str(pdf_val).strip()):
                continue
            pdf_str = str(pdf_val)
            rel     = os.path.relpath(os.path.dirname(pdf_str), C.RECORDS_FOLDER)
            entry   = (i, pdf_str)
            if rel == '.':
                root_docs.append(entry)
            else:
                grouped.setdefault(rel, []).append(entry)

        def add_file_item(parent, row_idx, pdf_path):
            name     = os.path.basename(pdf_path)
            dia_code = extract_dia_code(name)
            doc_id   = get_id_from_json(C.KBM_MAP, dia_code, normalize(name)) if dia_code else None
            label    = name + (f" ({doc_id})" if doc_id else "")
            item     = QTreeWidgetItem(parent)
            item.setText(0, label)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsSelectable)
            item.setCheckState(0, Qt.Unchecked)
            item.setData(0, Qt.UserRole, row_idx)

        for row_idx, pdf_path in root_docs:
            add_file_item(self.doc_list, row_idx, pdf_path)

        for folder_name, entries in sorted(grouped.items()):
            folder_item = QTreeWidgetItem(self.doc_list)
            folder_item.setText(0, folder_name)
            folder_item.setFlags(Qt.ItemIsEnabled)
            folder_item.setFont(0, QFont("Segoe UI", 9, QFont.Bold))
            folder_item.setExpanded(True)
            for row_idx, pdf_path in entries:
                add_file_item(folder_item, row_idx, pdf_path)

    def get_mode(self):
        return self.mode_group.checkedId()

    def get_selected_indices(self):
        indices = []
        for i in range(self.doc_list.topLevelItemCount()):
            top = self.doc_list.topLevelItem(i)
            if (top.flags() & Qt.ItemIsUserCheckable) and top.checkState(0) == Qt.Checked:
                indices.append(top.data(0, Qt.UserRole))
            for j in range(top.childCount()):
                child = top.child(j)
                if (child.flags() & Qt.ItemIsUserCheckable) and child.checkState(0) == Qt.Checked:
                    indices.append(child.data(0, Qt.UserRole))
        return indices

    def reset(self):
        self.mode_rev_and_redline.setChecked(True)
        self.select_all_cb.setChecked(False)
        self.select_new_cb.setChecked(False)
        self.doc_list.clear()


# ── Page 4: Processing / Results ──────────────────────────────────────────────


class ProcessingPage(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        self.title_label = QLabel("Processing Documents")
        self.title_label.setFont(QFont("Segoe UI", 16, QFont.Bold))
        layout.addWidget(self.title_label)

        self.progress_label = QLabel("")
        layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setFont(QFont("Consolas", 9))
        layout.addWidget(self.log_output)

        self.debug_cb = QCheckBox("Show Debug Info")
        layout.addWidget(self.debug_cb)

        btn_layout = QHBoxLayout()
        self.save_btn       = QPushButton("Save Log")
        self.start_over_btn = QPushButton("Start Over")
        self.close_btn      = QPushButton("Close")
        self.save_btn.hide()
        self.start_over_btn.hide()
        self.close_btn.hide()
        btn_layout.addWidget(self.save_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.start_over_btn)
        btn_layout.addWidget(self.close_btn)
        layout.addLayout(btn_layout)

        self._log_entries = []  # list of (text, is_debug)

        self.save_btn.clicked.connect(self._save_log)
        self.close_btn.clicked.connect(QApplication.quit)
        self.debug_cb.stateChanged.connect(lambda _: self._render_log())

    def append_log(self, text):
        is_debug = text.strip().startswith("[DEBUG]")
        self._log_entries.append((text, is_debug))
        self._render_log()

    def _render_log(self):
        debug_on = self.debug_cb.isChecked()
        parts    = []
        for text, is_debug in self._log_entries:
            if not is_debug or debug_on:
                parts.append(text if text.endswith("\n") else text + "\n")
        self.log_output.setPlainText("".join(parts))
        scrollbar = self.log_output.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def on_progress(self, current, total):
        if total > 0:
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(current)
            self.progress_label.setText(f"Document {current} of {total}")

    def on_finished(self, success, message):
        self.title_label.setText("Complete" if success else "Error")
        self.append_log(f"\n{'='*40}")
        self.append_log(message)
        self.progress_bar.hide()
        self.progress_label.hide()
        self._show_result_buttons()

    def show_aborted(self):
        self.title_label.setText("Aborted")
        self.progress_bar.hide()
        self.progress_label.hide()
        self._log_entries = []
        self.append_log("Operation aborted by user.")
        self.append_log("Missing DOCX files were detected and the user chose to abort.")
        self._show_result_buttons()

    def _show_result_buttons(self):
        self.save_btn.show()
        self.start_over_btn.show()
        self.close_btn.show()

    def _save_log(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Log", "whitney_log.txt", "Text Files (*.txt)"
        )
        if path:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(self.log_output.toPlainText())

    def reset(self):
        self.title_label.setText("Processing Documents")
        self.progress_label.setText("")
        self.progress_label.show()
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.show()
        self._log_entries = []
        self.log_output.clear()
        self.save_btn.hide()
        self.start_over_btn.hide()
        self.close_btn.hide()


# ── Wizard widget (replaces WhitneyApp) ───────────────────────────────────────


class WhitneyWizard(QWidget):
    """Five-page wizard — lives inside the MDI subwindow."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Whitney Document Tool")
        self.setMinimumSize(700, 500)
        self.resize(850, 620)

        self.records_folder  = ""
        self.excel_path      = ""
        self.docDF           = None
        self.missing_doc_ids = []
        self._page_history   = []
        self._process_worker = None

        main_layout = QVBoxLayout(self)

        self.stack = QStackedWidget()
        main_layout.addWidget(self.stack)

        nav_layout = QHBoxLayout()
        self.back_btn  = QPushButton("Back")
        self.abort_btn = QPushButton("Abort")
        self.next_btn  = QPushButton("Next")
        self.abort_btn.setStyleSheet("color: #c0392b; font-weight: bold;")
        self.back_btn.clicked.connect(self._go_back)
        self.abort_btn.clicked.connect(self._on_abort)
        self.next_btn.clicked.connect(self._go_next)
        nav_layout.addWidget(self.back_btn)
        nav_layout.addStretch()
        nav_layout.addWidget(self.abort_btn)
        nav_layout.addWidget(self.next_btn)
        main_layout.addLayout(nav_layout)

        self.folder_page     = FolderPage()
        self.summary_page    = SummaryPage()
        self.missing_page    = MissingDocsPage()
        self.mode_page       = ModePage()
        self.processing_page = ProcessingPage()

        self.stack.addWidget(self.folder_page)      # 0
        self.stack.addWidget(self.summary_page)     # 1
        self.stack.addWidget(self.missing_page)     # 2
        self.stack.addWidget(self.mode_page)        # 3
        self.stack.addWidget(self.processing_page)  # 4

        self.folder_page.folder_changed.connect(self._on_folder_changed)
        self.summary_page.scan_finished.connect(self._on_scan_finished)
        self.processing_page.start_over_btn.clicked.connect(self._start_over)

        self._update_nav()

    # ── Navigation ────────────────────────────────────────────────────────────

    def _on_folder_changed(self, folder):
        self.records_folder = folder
        self.excel_path     = os.path.join(folder, "whitney_output.xlsx")
        self._update_nav()

    def _on_scan_finished(self, docDF, missing_doc_ids):
        self.docDF           = docDF
        self.missing_doc_ids = missing_doc_ids
        if missing_doc_ids:
            self._page_history.append(self.stack.currentIndex())
            self.missing_page.on_enter(missing_doc_ids)
            self._navigate_to(2)
        else:
            self._update_nav()

    def _go_next(self):
        current = self.stack.currentIndex()

        if current == 0:
            self._page_history.append(current)
            self._navigate_to(1)
            self.summary_page.on_enter(self.records_folder, self.excel_path)

        elif current == 1:
            self._page_history.append(current)
            self.mode_page.populate_docs(self.docDF)
            self._navigate_to(3)

        elif current == 2:
            self._page_history.append(current)
            self._navigate_to(1)

        elif current == 3:
            indices = self.mode_page.get_selected_indices()
            if indices is not None and len(indices) == 0:
                QMessageBox.warning(
                    self, "No Documents Selected",
                    "Please select at least one document to process."
                )
                return
            self._page_history.append(current)
            self._navigate_to(4)
            self._start_processing()

    def _go_back(self):
        if self._page_history:
            prev = self._page_history.pop()
            self._navigate_to(prev)

    def _on_abort(self):
        self._page_history.append(self.stack.currentIndex())
        self._navigate_to(4)
        self.processing_page.show_aborted()

    def _navigate_to(self, page_idx):
        self.stack.setCurrentIndex(page_idx)
        self._update_nav()

    def _update_nav(self):
        idx = self.stack.currentIndex()

        self.back_btn.setVisible(idx in (1, 2, 3))
        self.abort_btn.setVisible(idx == 2)

        if idx == 0:
            self.next_btn.setText("Next")
            self.next_btn.setEnabled(bool(self.records_folder))
            self.next_btn.show()
        elif idx == 1:
            self.next_btn.setText("Next")
            self.next_btn.setEnabled(self.docDF is not None)
            self.next_btn.show()
        elif idx == 2:
            self.next_btn.setText("Continue Anyway")
            self.next_btn.setEnabled(True)
            self.next_btn.show()
        elif idx == 3:
            self.next_btn.setText("Run")
            self.next_btn.setEnabled(True)
            self.next_btn.show()
        elif idx == 4:
            self.next_btn.hide()
            self.back_btn.hide()
            self.abort_btn.hide()

    # ── Processing ────────────────────────────────────────────────────────────

    def _start_processing(self):
        C.RECORDS_FOLDER  = self.records_folder
        C.OUTDATED_FOLDER = os.path.join(C.RECORDS_FOLDER, 'Outdated')
        C.REDLINE_FOLDER  = os.path.join(C.RECORDS_FOLDER, 'Redline')
        C.PREV_REV_FOLDER = os.path.join(C.RECORDS_FOLDER, 'Previous Revisions')

        self.processing_page.reset()

        mode        = self.mode_page.get_mode()
        doc_indices = self.mode_page.get_selected_indices()

        self._process_worker = ProcessWorker(self.docDF, mode, doc_indices)
        self._process_worker.log_signal.connect(self.processing_page.append_log)
        self._process_worker.progress_signal.connect(self.processing_page.on_progress)
        self._process_worker.finished_signal.connect(self.processing_page.on_finished)
        self._process_worker.start()

    # ── Start Over ────────────────────────────────────────────────────────────

    def _start_over(self):
        self._page_history.clear()
        self.records_folder  = ""
        self.excel_path      = ""
        self.docDF           = None
        self.missing_doc_ids = []
        self.folder_page.reset()
        self.summary_page.reset()
        self.mode_page.reset()
        self.processing_page.reset()
        self._navigate_to(0)

    def is_processing(self):
        return self._process_worker is not None and self._process_worker.isRunning()

    def closeEvent(self, event):
        if self.is_processing():
            reply = QMessageBox.question(
                self, "Processing in Progress",
                "Documents are still being processed. Are you sure you want to quit?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.No:
                event.ignore()
                return
        event.accept()


# ── Plugin registration ────────────────────────────────────────────────────────


class Loader:
    """
    Entry point instantiated by the engine's dynamic plugin loader (see
    plugins/__init__.py and Engine.__init_plugins in engine/engine.py).
    The wizard builds its widget tree entirely in code rather than from an
    interface.ui file, so it wraps WhitneyWizard directly instead of going
    through DynamicLoader's ui-file loading.
    """

    isVisible   = True
    canMaximize = True

    def __init__(self):
        self.window = WhitneyWizard()

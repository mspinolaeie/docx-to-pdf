from __future__ import annotations

import logging
import os
from typing import Optional

from PySide6.QtCore import QObject, Qt, QThread, Signal, Slot
from PySide6.QtGui import QColor, QCloseEvent, QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QHeaderView,
)

import docx_to_pdf as core


STATUS_LABELS = {
    "queued": "In coda",
    "running": "In corso",
    "converted": "Convertito",
    "skipped": "Saltato",
    "error": "Errore",
}

STATUS_COLORS = {
    "queued": QColor("#6b7280"),
    "running": QColor("#1d4ed8"),
    "converted": QColor("#047857"),
    "skipped": QColor("#92400e"),
    "error": QColor("#b91c1c"),
}


class DropFrame(QFrame):
    paths_dropped = Signal(list)

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setObjectName("dropFrame")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)

        title = QLabel("Trascina qui DOCX o cartelle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 18px; font-weight: 600;")
        subtitle = QLabel("La cartella viene espansa solo al primo livello. Nessuna conversione parte automaticamente.")
        subtitle.setWordWrap(True)
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("color: #475569;")

        layout.addStretch()
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addStretch()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        urls = event.mimeData().urls()
        paths = [url.toLocalFile() for url in urls if url.isLocalFile()]
        if paths:
            self.paths_dropped.emit(paths)
            event.acceptProposedAction()
            return
        event.ignore()


class GuiLogHandler(logging.Handler):
    def __init__(self, emit_message) -> None:
        super().__init__()
        self._emit_message = emit_message

    def emit(self, record: logging.LogRecord) -> None:
        self._emit_message(self.format(record))


class ConversionWorker(QObject):
    log_message = Signal(str)
    file_status = Signal(str, str, str)
    progress = Signal(int, int)
    finished = Signal(object, str)
    failed = Signal(str)

    def __init__(self, files: list[str], config: core.ConversionConfig) -> None:
        super().__init__()
        self._files = files
        self._config = config

    @Slot()
    def run(self) -> None:
        handler = GuiLogHandler(self.log_message.emit)
        handler.setLevel(logging.DEBUG)
        handler.setFormatter(logging.Formatter("%(levelname)s | %(message)s"))

        try:
            core.setup_logging(
                self._config.log_level,
                self._config.log_file,
                console=False,
                extra_handlers=[handler],
            )
            self.progress.emit(0, len(self._files))

            def on_event(event: core.ConversionEvent) -> None:
                if event.status == "started":
                    self.file_status.emit(event.docx_path, "running", event.message or "")
                elif event.status == "skipped":
                    self.file_status.emit(event.docx_path, "skipped", event.message or "")
                    self.progress.emit(event.processed, event.total)
                elif event.status == "completed":
                    row_status = "converted" if event.result and event.result.success else "error"
                    message = ""
                    if event.result and not event.result.success:
                        message = event.result.error_message or "Conversione fallita"
                    self.file_status.emit(event.docx_path, row_status, message)
                    self.progress.emit(event.processed, event.total)

            results, backend = core.run_conversion_for_files(self._files, self._config, progress_callback=on_event)
            self.finished.emit(results, backend)
        except Exception as exc:
            self.failed.emit(str(exc))
        finally:
            logger = core.get_logger()
            if handler in logger.handlers:
                logger.removeHandler(handler)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self._thread: Optional[QThread] = None
        self._worker: Optional[ConversionWorker] = None
        self._row_by_path: dict[str, int] = {}

        self.setWindowTitle("DOCX to PDF")
        self.resize(1180, 760)
        self._build_ui()
        self._update_backend_hint()

    def _build_ui(self) -> None:
        root = QWidget(self)
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(16, 16, 16, 16)
        root_layout.setSpacing(12)

        headline = QLabel("Conversione DOCX -> PDF")
        headline.setStyleSheet("font-size: 24px; font-weight: 700;")
        subhead = QLabel("Seleziona i documenti con drag and drop, imposta le opzioni globali e avvia la coda.")
        subhead.setStyleSheet("color: #475569;")
        root_layout.addWidget(headline)
        root_layout.addWidget(subhead)

        splitter = QSplitter(Qt.Orientation.Horizontal, self)
        splitter.addWidget(self._build_queue_panel())
        splitter.addWidget(self._build_options_panel())
        splitter.setSizes([840, 340])
        root_layout.addWidget(splitter, 1)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%v/%m file")
        root_layout.addWidget(self.progress_bar)

        self.log_box = QPlainTextEdit(self)
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Log conversione...")
        self.log_box.setMaximumBlockCount(1000)
        root_layout.addWidget(self.log_box, 1)

        self.setCentralWidget(root)
        self.statusBar().showMessage("Pronto")
        self.setStyleSheet(
            """
            QMainWindow { background: #f8fafc; }
            QGroupBox {
                background: white;
                border: 1px solid #dbe2ea;
                border-radius: 12px;
                margin-top: 12px;
                font-weight: 600;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 14px;
                padding: 0 4px;
            }
            QTableWidget, QPlainTextEdit, QComboBox, QSpinBox {
                background: white;
                border: 1px solid #dbe2ea;
                border-radius: 8px;
            }
            QPushButton {
                background: #0f172a;
                color: white;
                border: 0;
                border-radius: 8px;
                padding: 8px 14px;
            }
            QPushButton:disabled {
                background: #94a3b8;
            }
            QToolButton {
                color: #0f172a;
                font-weight: 600;
            }
            #dropFrame {
                background: #eff6ff;
                border: 2px dashed #60a5fa;
                border-radius: 14px;
            }
            """
        )

    def _build_queue_panel(self) -> QWidget:
        group = QGroupBox("Coda", self)
        layout = QVBoxLayout(group)
        layout.setSpacing(10)

        self.drop_frame = DropFrame(group)
        self.drop_frame.paths_dropped.connect(self.add_input_paths)
        layout.addWidget(self.drop_frame)

        self.table = QTableWidget(0, 3, group)
        self.table.setHorizontalHeaderLabels(["DOCX", "PDF", "Stato"])
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.table, 1)

        buttons = QHBoxLayout()
        self.add_files_button = QPushButton("Aggiungi file", group)
        self.add_folder_button = QPushButton("Aggiungi cartella", group)
        self.remove_button = QPushButton("Rimuovi selezionati", group)
        self.clear_button = QPushButton("Svuota", group)
        buttons.addWidget(self.add_files_button)
        buttons.addWidget(self.add_folder_button)
        buttons.addWidget(self.remove_button)
        buttons.addWidget(self.clear_button)
        layout.addLayout(buttons)

        self.add_files_button.clicked.connect(self._pick_files)
        self.add_folder_button.clicked.connect(self._pick_folder)
        self.remove_button.clicked.connect(self._remove_selected_rows)
        self.clear_button.clicked.connect(self.clear_queue)

        return group

    def _build_options_panel(self) -> QWidget:
        group = QGroupBox("Opzioni", self)
        layout = QVBoxLayout(group)
        layout.setSpacing(10)

        base_widget = QWidget(group)
        base_form = QFormLayout(base_widget)
        base_form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)

        self.overwrite_checkbox = QCheckBox("Sovrascrivi PDF esistenti", base_widget)
        self.backend_combo = QComboBox(base_widget)
        self.backend_combo.addItem("Auto", "auto")
        self.backend_combo.addItem("Microsoft Word", "word")
        self.backend_combo.addItem("LibreOffice", "libreoffice")
        self.backend_combo.currentIndexChanged.connect(self._update_backend_hint)

        self.workers_spin = QSpinBox(base_widget)
        self.workers_spin.setRange(1, max(1, os.cpu_count() or 1))
        self.workers_spin.setValue(core.default_gui_workers())

        base_form.addRow("Backend", self.backend_combo)
        base_form.addRow("Worker", self.workers_spin)
        base_form.addRow("", self.overwrite_checkbox)
        layout.addWidget(base_widget)

        self.backend_hint = QLabel(group)
        self.backend_hint.setWordWrap(True)
        self.backend_hint.setStyleSheet("color: #475569;")
        layout.addWidget(self.backend_hint)

        self.advanced_toggle = QToolButton(group)
        self.advanced_toggle.setText("Opzioni avanzate")
        self.advanced_toggle.setCheckable(True)
        self.advanced_toggle.setChecked(False)
        self.advanced_toggle.setArrowType(Qt.ArrowType.RightArrow)
        self.advanced_toggle.toggled.connect(self._toggle_advanced)
        layout.addWidget(self.advanced_toggle)

        self.advanced_panel = QWidget(group)
        self.advanced_panel.setVisible(False)
        adv_form = QFormLayout(self.advanced_panel)

        self.bookmarks_combo = QComboBox(self.advanced_panel)
        self.bookmarks_combo.addItem("Titoli Word", "headings")
        self.bookmarks_combo.addItem("Bookmark Word", "word")
        self.bookmarks_combo.addItem("Nessuno", "none")

        self.pdfa_checkbox = QCheckBox("Esporta PDF/A-1", self.advanced_panel)
        self.validate_checkbox = QCheckBox("Valida il PDF finale", self.advanced_panel)
        self.validate_checkbox.setChecked(True)

        self.log_level_combo = QComboBox(self.advanced_panel)
        for level in ["INFO", "DEBUG", "WARNING", "ERROR"]:
            self.log_level_combo.addItem(level, level)

        adv_form.addRow("Segnalibri", self.bookmarks_combo)
        adv_form.addRow("", self.pdfa_checkbox)
        adv_form.addRow("", self.validate_checkbox)
        adv_form.addRow("Log level", self.log_level_combo)
        layout.addWidget(self.advanced_panel)

        layout.addStretch()
        self.convert_button = QPushButton("Converti", group)
        self.convert_button.setDefault(True)
        self.convert_button.clicked.connect(self.start_conversion)
        layout.addWidget(self.convert_button)

        return group

    def _toggle_advanced(self, checked: bool) -> None:
        self.advanced_panel.setVisible(checked)
        self.advanced_toggle.setArrowType(Qt.ArrowType.DownArrow if checked else Qt.ArrowType.RightArrow)

    def _update_backend_hint(self) -> None:
        backend = self.backend_combo.currentData()
        if backend == "libreoffice":
            self.pdfa_checkbox.setChecked(False)
            self.pdfa_checkbox.setEnabled(False)
            self.backend_hint.setText("PDF/A non è disponibile con LibreOffice. Se selezioni Word, i worker effettivi verranno forzati a 1.")
            return
        self.pdfa_checkbox.setEnabled(True)
        if backend == "word":
            self.backend_hint.setText("Con Microsoft Word la conversione usa un'unica istanza dell'applicazione; i worker > 1 vengono ignorati.")
            return
        self.backend_hint.setText("In modalità Auto, PDF/A funziona solo se il backend risolto è Microsoft Word.")

    def _pick_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Seleziona DOCX", "", "Documenti Word (*.docx)")
        if files:
            self.add_input_paths(files)

    def _pick_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Seleziona cartella")
        if folder:
            self.add_input_paths([folder])

    @Slot(list)
    def add_input_paths(self, paths: list[str]) -> None:
        files = core.collect_docx_inputs(paths, recursive=False)
        added = 0
        for path in files:
            if path in self._row_by_path:
                continue
            self._append_row(path)
            added += 1

        if added == 0:
            self.statusBar().showMessage("Nessun nuovo DOCX valido trovato", 4000)
        else:
            self.statusBar().showMessage(f"Aggiunti {added} file", 4000)
            self.progress_bar.setRange(0, max(1, self.table.rowCount()))
            self.progress_bar.setValue(0)

    def _append_row(self, path: str) -> None:
        row = self.table.rowCount()
        self.table.insertRow(row)

        docx_item = QTableWidgetItem(path)
        docx_item.setData(Qt.ItemDataRole.UserRole, path)
        pdf_item = QTableWidgetItem(os.path.splitext(path)[0] + ".pdf")
        status_item = QTableWidgetItem()

        self.table.setItem(row, 0, docx_item)
        self.table.setItem(row, 1, pdf_item)
        self.table.setItem(row, 2, status_item)
        self._row_by_path[path] = row
        self._set_row_status(path, "queued")

    def queue_paths(self) -> list[str]:
        return [self.table.item(row, 0).text() for row in range(self.table.rowCount())]

    def clear_queue(self) -> None:
        if self._thread is not None and self._thread.isRunning():
            return
        self.table.setRowCount(0)
        self._row_by_path.clear()
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.statusBar().showMessage("Coda svuotata", 3000)

    def _remove_selected_rows(self) -> None:
        if self._thread is not None and self._thread.isRunning():
            return
        rows = sorted({index.row() for index in self.table.selectionModel().selectedRows()}, reverse=True)
        for row in rows:
            path = self.table.item(row, 0).text()
            self._row_by_path.pop(path, None)
            self.table.removeRow(row)
        self._rebuild_row_index()

    def _rebuild_row_index(self) -> None:
        self._row_by_path.clear()
        for row in range(self.table.rowCount()):
            path = self.table.item(row, 0).text()
            self._row_by_path[path] = row

    def _set_row_status(self, path: str, status: str, message: str = "") -> None:
        row = self._row_by_path.get(path)
        if row is None:
            return
        item = self.table.item(row, 2)
        item.setText(STATUS_LABELS.get(status, status))
        item.setForeground(STATUS_COLORS.get(status, QColor("#0f172a")))
        item.setToolTip(message)

    def _build_config(self) -> core.ConversionConfig:
        root_dir = core.determine_display_root(self.queue_paths()) or os.getcwd()
        return core.build_gui_config(
            root_dir=root_dir,
            overwrite=self.overwrite_checkbox.isChecked(),
            backend=self.backend_combo.currentData(),
            workers=self.workers_spin.value(),
            bookmarks=self.bookmarks_combo.currentData(),
            pdfa=self.pdfa_checkbox.isChecked() and self.pdfa_checkbox.isEnabled(),
            validate_pdf=self.validate_checkbox.isChecked(),
            log_level=self.log_level_combo.currentData(),
            log_file=None,
        )

    def _set_running_state(self, running: bool) -> None:
        for widget in [
            self.drop_frame,
            self.table,
            self.add_files_button,
            self.add_folder_button,
            self.remove_button,
            self.clear_button,
            self.backend_combo,
            self.workers_spin,
            self.overwrite_checkbox,
            self.bookmarks_combo,
            self.pdfa_checkbox,
            self.validate_checkbox,
            self.log_level_combo,
            self.advanced_toggle,
            self.convert_button,
        ]:
            widget.setEnabled(not running)
        self.progress_bar.setValue(0 if not running and self.table.rowCount() == 0 else self.progress_bar.value())

    def start_conversion(self) -> None:
        files = self.queue_paths()
        if not files:
            QMessageBox.information(self, "DOCX to PDF", "Aggiungi almeno un file DOCX prima di convertire.")
            return

        for path in files:
            self._set_row_status(path, "queued")

        self.log_box.clear()
        self.progress_bar.setRange(0, len(files))
        self.progress_bar.setValue(0)
        self._set_running_state(True)
        self.statusBar().showMessage("Conversione in corso...")

        self._thread = QThread(self)
        self._worker = ConversionWorker(files, self._build_config())
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.log_message.connect(self._append_log)
        self._worker.file_status.connect(self._on_file_status)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._on_failed)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.finished.connect(self._cleanup_worker)

        self._thread.start()

    @Slot(str)
    def _append_log(self, message: str) -> None:
        self.log_box.appendPlainText(message)

    @Slot(str, str, str)
    def _on_file_status(self, path: str, status: str, message: str) -> None:
        self._set_row_status(path, status, message)

    @Slot(int, int)
    def _on_progress(self, value: int, total: int) -> None:
        self.progress_bar.setRange(0, max(1, total))
        self.progress_bar.setValue(value)

    @Slot(object, str)
    def _on_finished(self, _results: object, backend: str) -> None:
        self._set_running_state(False)
        counts = {"converted": 0, "skipped": 0, "error": 0}
        for row in range(self.table.rowCount()):
            status_text = self.table.item(row, 2).text()
            for key, label in STATUS_LABELS.items():
                if label == status_text and key in counts:
                    counts[key] += 1

        self.statusBar().showMessage("Conversione completata", 5000)
        QMessageBox.information(
            self,
            "DOCX to PDF",
            (
                f"Backend: {backend}\n"
                f"Convertiti: {counts['converted']}\n"
                f"Saltati: {counts['skipped']}\n"
                f"Errori: {counts['error']}"
            ),
        )

    @Slot(str)
    def _on_failed(self, message: str) -> None:
        self._set_running_state(False)
        self.statusBar().showMessage("Conversione fallita", 5000)
        QMessageBox.critical(self, "DOCX to PDF", message)

    @Slot()
    def _cleanup_worker(self) -> None:
        self._worker = None
        self._thread = None

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._thread is not None and self._thread.isRunning():
            QMessageBox.warning(self, "DOCX to PDF", "Attendi la fine della conversione prima di chiudere la finestra.")
            event.ignore()
            return
        super().closeEvent(event)


def launch_gui() -> int:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.show()
    return app.exec()

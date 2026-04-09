"""Minimal GUI tests for the PySide6 queue window."""
import os
import sys

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

pytest.importorskip("PySide6")
from PySide6.QtWidgets import QApplication

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_to_pdf_gui import MainWindow


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication([])


def test_main_window_collects_first_level_inputs(app, tmp_path):
    root_file = tmp_path / "root.docx"
    root_file.write_bytes(b"x")
    folder = tmp_path / "folder"
    folder.mkdir()
    nested = folder / "nested.docx"
    nested.write_bytes(b"x")
    deep = folder / "deep"
    deep.mkdir()
    (deep / "deep.docx").write_bytes(b"x")

    window = MainWindow()
    window.add_input_paths([str(folder), str(root_file)])

    assert window.queue_paths() == [str(nested), str(root_file)]
    assert window.table.item(0, 2).text() == "In coda"
    window.close()


def test_main_window_builds_config_and_locks_controls(app):
    window = MainWindow()
    window.overwrite_checkbox.setChecked(True)
    window.backend_combo.setCurrentIndex(window.backend_combo.findData("word"))
    window.workers_spin.setValue(3)
    window.bookmarks_combo.setCurrentIndex(window.bookmarks_combo.findData("word"))
    window.pdfa_checkbox.setChecked(True)
    window.validate_checkbox.setChecked(False)
    window.log_level_combo.setCurrentIndex(window.log_level_combo.findData("DEBUG"))

    cfg = window._build_config()
    assert cfg.overwrite is True
    assert cfg.backend == "word"
    assert cfg.workers == 3
    assert cfg.bookmarks == "word"
    assert cfg.pdfa is True
    assert cfg.validate_pdf is False
    assert cfg.log_level == "DEBUG"

    window._set_running_state(True)
    assert window.add_files_button.isEnabled() is False
    assert window.convert_button.isEnabled() is False
    window._set_running_state(False)
    assert window.add_files_button.isEnabled() is True
    assert window.convert_button.isEnabled() is True
    window.close()

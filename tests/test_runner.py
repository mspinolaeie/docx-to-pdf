"""Tests for explicit file collection and file-list conversion flow."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import docx_to_pdf as mod


def test_collect_docx_inputs_deduplicates_and_ignores_nested_when_non_recursive(tmp_path):
    root_file = tmp_path / "root.docx"
    root_file.write_bytes(b"x")
    folder = tmp_path / "folder"
    folder.mkdir()
    nested_file = folder / "nested.docx"
    nested_file.write_bytes(b"x")
    deep_folder = folder / "deep"
    deep_folder.mkdir()
    deep_file = deep_folder / "deep.docx"
    deep_file.write_bytes(b"x")
    hidden = folder / ".hidden.docx"
    hidden.write_bytes(b"x")
    temp = folder / "~$temp.docx"
    temp.write_bytes(b"x")

    files = mod.collect_docx_inputs([str(folder), str(root_file), str(root_file)], recursive=False)

    assert files == sorted([str(nested_file), str(root_file)])


def test_build_gui_config_maps_values():
    cfg = mod.build_gui_config(
        root_dir="C:/work",
        overwrite=True,
        backend="word",
        workers=2,
        bookmarks="word",
        pdfa=True,
        validate_pdf=False,
        log_level="DEBUG",
        log_file=None,
    )

    assert cfg.root_dir == "C:/work"
    assert cfg.overwrite is True
    assert cfg.backend == "word"
    assert cfg.workers == 2
    assert cfg.bookmarks == "word"
    assert cfg.pdfa is True
    assert cfg.validate_pdf is False
    assert cfg.log_level == "DEBUG"


def test_run_conversion_for_files_uses_adjacent_output_and_emits_skip(tmp_path, monkeypatch):
    first = tmp_path / "a.docx"
    first.write_bytes(b"x")
    second = tmp_path / "b.docx"
    second.write_bytes(b"x")
    existing_pdf = tmp_path / "b.pdf"
    existing_pdf.write_bytes(b"already-here")

    calls = []
    events = []

    def fake_convert_single_file(docx_path, pdf_path, backend, config, word_app=None):
        calls.append((docx_path, pdf_path, backend, config.overwrite))
        return mod.ConversionResult(
            docx_path=docx_path,
            pdf_path=pdf_path,
            success=True,
            backend_used=backend,
            metadata_injected=False,
            validation_passed=False,
        )

    monkeypatch.setattr(mod, "resolve_backend", lambda config: ("libreoffice", False))
    monkeypatch.setattr(mod, "convert_single_file", fake_convert_single_file)

    cfg = mod.ConversionConfig(root_dir=str(tmp_path), overwrite=False, validate_pdf=False, workers=1)
    results, backend = mod.run_conversion_for_files([str(second), str(first)], cfg, progress_callback=events.append)

    assert backend == "libreoffice"
    assert len(results) == 1
    assert results[0].docx_path == str(first)
    assert results[0].pdf_path == str(tmp_path / "a.pdf")
    assert calls == [(str(first), str(tmp_path / "a.pdf"), "libreoffice", False)]

    skip_events = [event for event in events if event.status == "skipped"]
    assert len(skip_events) == 1
    assert skip_events[0].docx_path == str(second)
    assert skip_events[0].message == "PDF already exists"

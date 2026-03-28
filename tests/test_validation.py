"""Tests for PDF validation."""
import builtins
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_to_pdf import validate_pdf


def test_missing_file(tmp_path):
    ok, err = validate_pdf(str(tmp_path / "nonexistent.pdf"))
    assert ok is False
    assert "does not exist" in err


def test_empty_file(tmp_path):
    f = tmp_path / "empty.pdf"
    f.write_bytes(b"")
    ok, err = validate_pdf(str(f))
    assert ok is False
    assert "empty" in err.lower()


def test_too_small_file(tmp_path):
    f = tmp_path / "tiny.pdf"
    f.write_bytes(b"x" * 50)
    ok, err = validate_pdf(str(f))
    assert ok is False
    assert "small" in err.lower()


def test_valid_size_without_pypdf(tmp_path, monkeypatch):
    """A file >100 bytes should pass when pypdf is not available."""
    f = tmp_path / "ok.pdf"
    f.write_bytes(b"%PDF-1.4\n" + b"x" * 200)

    real_import = builtins.__import__

    def mock_import(name, *args, **kwargs):
        if name == "pypdf":
            raise ImportError("pypdf not available")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", mock_import)
    ok, err = validate_pdf(str(f))
    assert ok is True
    assert err is None


def test_valid_pdf_with_pypdf(tmp_path):
    """A well-formed PDF should pass full validation when pypdf is available."""
    pytest.importorskip("pypdf")
    from pypdf import PdfWriter

    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)
    pdf_path = str(tmp_path / "valid.pdf")
    with open(pdf_path, "wb") as f:
        writer.write(f)

    ok, err = validate_pdf(pdf_path)
    assert ok is True
    assert err is None

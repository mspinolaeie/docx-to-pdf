"""Regression tests for LibreOffice success without an output PDF."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pytest

from docx_to_pdf import convert_with_libreoffice


def test_zero_exit_without_pdf_is_failure(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"fixture")
    target = tmp_path / "source.pdf"
    with patch("docx_to_pdf.shutil.which", return_value="soffice"), patch(
        "docx_to_pdf.subprocess.run",
        return_value=SimpleNamespace(returncode=0, stdout="", stderr=""),
    ):
        with pytest.raises(RuntimeError, match="LibreOffice output not found"):
            convert_with_libreoffice(str(source), str(target))
    assert not target.exists()


def test_zero_exit_with_expected_pdf_succeeds(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"fixture")
    target = tmp_path / "source.pdf"
    target.write_bytes(b"%PDF-1.4 fixture")
    with patch("docx_to_pdf.shutil.which", return_value="soffice"), patch(
        "docx_to_pdf.subprocess.run",
        return_value=SimpleNamespace(returncode=0, stdout="", stderr=""),
    ):
        convert_with_libreoffice(str(source), str(target))
    assert target.read_bytes() == b"%PDF-1.4 fixture"

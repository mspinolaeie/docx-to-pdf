"""Tests for DOCX metadata extraction and PDF metadata injection."""
import os
import sys
import zipfile

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_to_pdf import inject_pdf_metadata, read_docx_core_properties


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_minimal_docx(tmp_path, title="Test Title", author="Test Author", subject="Test Subject"):
    core_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>{title}</dc:title>
  <dc:creator>{author}</dc:creator>
  <dc:subject>{subject}</dc:subject>
</cp:coreProperties>"""
    path = tmp_path / "test.docx"
    with zipfile.ZipFile(str(path), "w") as z:
        z.writestr("docProps/core.xml", core_xml)
        z.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
        )
    return str(path)


def _make_minimal_pdf(tmp_path, name="out.pdf"):
    content = (
        b"%PDF-1.4\n"
        b"1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n"
        b"2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n"
        b"3 0 obj<</Type/Page/MediaBox[0 0 612 792]/Parent 2 0 R>>endobj\n"
        b"xref\n0 4\n"
        b"0000000000 65535 f \n"
        b"0000000009 00000 n \n"
        b"0000000058 00000 n \n"
        b"0000000115 00000 n \n"
        b"trailer<</Size 4/Root 1 0 R>>\n"
        b"startxref\n190\n%%EOF"
    )
    path = tmp_path / name
    path.write_bytes(content)
    return str(path)


# ---------------------------------------------------------------------------
# read_docx_core_properties
# ---------------------------------------------------------------------------


def test_read_core_properties_full(tmp_path):
    docx = _make_minimal_docx(tmp_path, title="My Doc", author="John Doe", subject="Engineering")
    meta = read_docx_core_properties(docx)
    assert meta.get("title") == "My Doc"
    assert meta.get("author") == "John Doe"
    assert meta.get("subject") == "Engineering"


def test_read_core_properties_missing_core_xml(tmp_path):
    """Should return empty dict without crashing when docProps/core.xml is absent."""
    path = tmp_path / "empty.docx"
    with zipfile.ZipFile(str(path), "w") as z:
        z.writestr("dummy.txt", "nothing here")
    meta = read_docx_core_properties(str(path))
    assert meta == {}


def test_read_core_properties_not_a_zip(tmp_path):
    """Should return empty dict without crashing for a corrupt/non-zip file."""
    path = tmp_path / "bad.docx"
    path.write_bytes(b"this is not a zip file")
    meta = read_docx_core_properties(str(path))
    assert meta == {}


def test_read_core_properties_empty_fields(tmp_path):
    """Fields present in XML but with empty text should not appear in result."""
    core_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title></dc:title>
  <dc:creator>  </dc:creator>
</cp:coreProperties>"""
    path = tmp_path / "test.docx"
    with zipfile.ZipFile(str(path), "w") as z:
        z.writestr("docProps/core.xml", core_xml)
    meta = read_docx_core_properties(str(path))
    assert "title" not in meta
    assert "author" not in meta


# ---------------------------------------------------------------------------
# inject_pdf_metadata
# ---------------------------------------------------------------------------


def test_inject_pdf_metadata_basic(tmp_path):
    pytest.importorskip("pypdf")
    pdf = _make_minimal_pdf(tmp_path)
    result = inject_pdf_metadata(pdf, {"title": "Injected Title", "author": "Test Author"})
    assert result is True
    from pypdf import PdfReader
    reader = PdfReader(pdf)
    assert reader.metadata.get("/Title") == "Injected Title"
    assert reader.metadata.get("/Author") == "Test Author"


def test_inject_pdf_metadata_empty_meta_returns_false(tmp_path):
    """Empty meta should return False without touching the file."""
    pdf = _make_minimal_pdf(tmp_path)
    mtime_before = os.path.getmtime(pdf)
    result = inject_pdf_metadata(pdf, {})
    assert result is False
    assert os.path.getmtime(pdf) == mtime_before


def test_inject_pdf_metadata_no_tmp_file_left(tmp_path):
    """The .tmp.pdf file must be cleaned up after successful injection."""
    pytest.importorskip("pypdf")
    pdf = _make_minimal_pdf(tmp_path)
    inject_pdf_metadata(pdf, {"title": "Clean"})
    assert not os.path.exists(pdf + ".tmp.pdf")

"""Tests for DOCX file discovery."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_to_pdf import find_docx_files, is_docx


def test_is_docx_valid(tmp_path):
    f = tmp_path / "document.docx"
    f.write_bytes(b"fake")
    assert is_docx(str(f)) is True


def test_is_docx_temp_word_file_excluded(tmp_path):
    f = tmp_path / "~$document.docx"
    f.write_bytes(b"fake")
    assert is_docx(str(f)) is False


def test_is_docx_hidden_file_excluded(tmp_path):
    f = tmp_path / ".hidden.docx"
    f.write_bytes(b"fake")
    assert is_docx(str(f)) is False


def test_is_docx_wrong_extension(tmp_path):
    f = tmp_path / "document.pdf"
    f.write_bytes(b"fake")
    assert is_docx(str(f)) is False


def test_is_docx_case_insensitive(tmp_path):
    f = tmp_path / "document.DOCX"
    f.write_bytes(b"fake")
    assert is_docx(str(f)) is True


def test_find_docx_non_recursive(tmp_path):
    (tmp_path / "a.docx").write_bytes(b"x")
    (tmp_path / "b.docx").write_bytes(b"x")
    (tmp_path / "~$temp.docx").write_bytes(b"x")   # excluded
    (tmp_path / "other.pdf").write_bytes(b"x")      # excluded
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "c.docx").write_bytes(b"x")              # not visible non-recursively

    found = find_docx_files(str(tmp_path), recursive=False)
    names = sorted(os.path.basename(f) for f in found)
    assert names == ["a.docx", "b.docx"]


def test_find_docx_recursive(tmp_path):
    (tmp_path / "a.docx").write_bytes(b"x")
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "b.docx").write_bytes(b"x")
    deep = sub / "deep"
    deep.mkdir()
    (deep / "c.docx").write_bytes(b"x")

    found = find_docx_files(str(tmp_path), recursive=True)
    names = sorted(os.path.basename(f) for f in found)
    assert names == ["a.docx", "b.docx", "c.docx"]


def test_find_docx_empty_dir(tmp_path):
    found = find_docx_files(str(tmp_path), recursive=False)
    assert found == []


def test_find_docx_results_are_sorted(tmp_path):
    for name in ["z.docx", "a.docx", "m.docx"]:
        (tmp_path / name).write_bytes(b"x")
    found = find_docx_files(str(tmp_path), recursive=False)
    assert found == sorted(found)

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOCX to PDF converter (single-file) with optional double-click GUI mode.

Features:
- Multi-backend support: Microsoft Word (COM) with fallback to LibreOffice
- Parallel processing (LibreOffice) for faster batch conversions
- Logging to console and optional log file
- Optional progress bars (tqdm)
- Optional PDF validation (pypdf)
- Optional metadata injection (pypdf) for LibreOffice conversions
- Configuration file support (JSON)
- Double-click mode (no args): folder picker + simple prompts

Requirements:
  - Word backend: Windows + Microsoft Word + pywin32 (pip install pywin32)
  - LibreOffice backend: soffice on PATH
  - Optional: pypdf (pip install pypdf) for metadata injection/validation
  - Optional: tqdm (pip install tqdm) for progress bars

Usage:
  python docx_to_pdf.py --dir . --recursive --workers 4
  python docx_to_pdf.py --config config.json
  (Double-click) run without arguments and pick a folder
"""

from __future__ import annotations

import argparse
import dataclasses
import functools
import json
import logging
import os
import shutil
import subprocess
import sys
import traceback
import zipfile
import xml.etree.ElementTree as ET
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import asdict, dataclass
from datetime import datetime
from enum import Enum
from typing import Any, Dict, List, Literal, Optional, Tuple


# Optional imports (no hard deps)
try:
    import win32com.client  # type: ignore

    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False

try:
    from tqdm import tqdm  # type: ignore

    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False


# ============================================================================
# CONFIGURATION AND ENUMS
# ============================================================================


class Backend(Enum):
    AUTO = "auto"
    WORD = "word"
    LIBREOFFICE = "libreoffice"


class BookmarkMode(Enum):
    HEADINGS = "headings"
    WORD = "word"
    NONE = "none"


@dataclass(frozen=True)
class ConversionConfig:
    root_dir: str = "."
    recursive: bool = False
    overwrite: bool = False
    pdfa: bool = False
    bookmarks: Literal["headings", "word", "none"] = "headings"
    backend: Literal["auto", "word", "libreoffice"] = "auto"
    workers: int = 1
    validate_pdf: bool = True
    log_level: str = "INFO"
    log_file: Optional[str] = None

    def validate(self) -> None:
        if not isinstance(self.root_dir, str) or not self.root_dir:
            raise ValueError("root_dir must be a non-empty string")
        if self.bookmarks not in {m.value for m in BookmarkMode}:
            raise ValueError(f"Invalid bookmarks: {self.bookmarks}")
        if self.backend not in {b.value for b in Backend}:
            raise ValueError(f"Invalid backend: {self.backend}")
        if not isinstance(self.workers, int) or self.workers < 1:
            raise ValueError("workers must be >= 1")
        if self.log_level.upper() not in {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}:
            raise ValueError(f"Invalid log_level: {self.log_level}")

    @classmethod
    def from_file(cls, config_path: str) -> "ConversionConfig":
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        known = {f.name for f in dataclasses.fields(cls)}
        unknown = set(data) - known
        if unknown:
            get_logger().warning(f"Ignoring unknown config key(s): {', '.join(sorted(unknown))}")
        cfg = cls(**{k: v for k, v in data.items() if k in known})
        cfg.validate()
        return cfg

    def to_file(self, config_path: str) -> None:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(asdict(self), f, indent=2)


@dataclass(frozen=True)
class ConversionResult:
    docx_path: str
    pdf_path: str
    success: bool
    backend_used: str
    error_message: Optional[str] = None
    traceback: Optional[str] = None
    metadata_injected: bool = False
    validation_passed: bool = False


# ============================================================================
# LOGGING
# ============================================================================


LOGGER_NAME = "docx_to_pdf"


class ColoredFormatter(logging.Formatter):
    COLORS = {
        "DEBUG": "\033[36m",
        "INFO": "\033[32m",
        "WARNING": "\033[33m",
        "ERROR": "\033[31m",
        "CRITICAL": "\033[35m",
    }
    RESET = "\033[0m"

    def format(self, record: logging.LogRecord) -> str:
        if record.levelname in self.COLORS:
            record.levelname = f"{self.COLORS[record.levelname]}{record.levelname}{self.RESET}"
        return super().format(record)


def get_logger() -> logging.Logger:
    return logging.getLogger(LOGGER_NAME)


# Avoid unexpected "lastResort" output in multiprocessing workers.
get_logger().addHandler(logging.NullHandler())


def setup_logging(log_level: str = "INFO", log_file: Optional[str] = None) -> logging.Logger:
    logger = get_logger()
    logger.setLevel(getattr(logging, log_level.upper(), logging.INFO))
    logger.handlers.clear()
    logger.propagate = False

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(ColoredFormatter("%(levelname)-8s | %(message)s"))
    logger.addHandler(console_handler)

    if log_file:
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(
            logging.Formatter(
                "%(asctime)s | %(levelname)-8s | %(funcName)s | %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )
        )
        logger.addHandler(file_handler)

    return logger


# ============================================================================
# DISCOVERY
# ============================================================================


def is_docx(path: str) -> bool:
    name = os.path.basename(path)
    if name.startswith("~$"):
        return False
    if name.startswith("."):
        return False
    return os.path.isfile(path) and name.lower().endswith(".docx")


def find_docx_files(root: str, recursive: bool) -> List[str]:
    files: List[str] = []
    if recursive:
        for r, _, fns in os.walk(root):
            for fn in fns:
                p = os.path.join(r, fn)
                if is_docx(p):
                    files.append(p)
    else:
        for fn in os.listdir(root):
            p = os.path.join(root, fn)
            if is_docx(p):
                files.append(p)

    get_logger().info(f"Found {len(files)} DOCX file(s) in {root}")
    return sorted(files)


# ============================================================================
# METADATA (DOCX core.xml) + PDF injection
# ============================================================================


def read_docx_core_properties(docx_path: str) -> Dict[str, str]:
    meta: Dict[str, str] = {}
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            with z.open("docProps/core.xml") as f:
                tree = ET.parse(f)
                root = tree.getroot()

        ns = {
            "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "dc": "http://purl.org/dc/elements/1.1/",
            "dcterms": "http://purl.org/dc/terms/",
        }

        def get_text(tag: str) -> Optional[str]:
            el = root.find(tag, ns)
            return el.text.strip() if el is not None and el.text else None

        title = get_text("dc:title")
        subject = get_text("dc:subject") or get_text("cp:subject")
        creator = get_text("dc:creator")
        description = get_text("dc:description")
        keywords = get_text("cp:keywords")

        if title:
            meta["title"] = title
        if subject:
            meta["subject"] = subject
        if creator:
            meta["author"] = creator
        if description:
            meta["description"] = description
        if keywords:
            meta["keywords"] = keywords

        get_logger().debug(f"Extracted metadata from {os.path.basename(docx_path)}: {meta}")
    except Exception as e:
        get_logger().warning(f"Failed to extract metadata from {docx_path}: {e}")

    return meta


def inject_pdf_metadata(pdf_path: str, meta: Dict[str, str]) -> bool:
    try:
        from pypdf import PdfReader, PdfWriter  # type: ignore
    except ImportError:
        get_logger().debug("pypdf not available, skipping metadata injection")
        return False

    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        info: Dict[str, str] = {}
        if meta.get("title"):
            info["/Title"] = meta["title"]
        if meta.get("author"):
            info["/Author"] = meta["author"]
        if meta.get("subject"):
            info["/Subject"] = meta["subject"]
        if meta.get("keywords"):
            info["/Keywords"] = meta["keywords"]

        if not info:
            return False

        writer.add_metadata(info)
        tmp_path = pdf_path + ".tmp.pdf"
        with open(tmp_path, "wb") as f:
            writer.write(f)
        os.replace(tmp_path, pdf_path)

        get_logger().debug(f"Injected metadata into {os.path.basename(pdf_path)}: {list(info.keys())}")
        return True
    except Exception as e:
        get_logger().warning(f"Failed to inject metadata into {pdf_path}: {e}")
        tmp_path = pdf_path + ".tmp.pdf"
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        return False


# ============================================================================
# PDF VALIDATION
# ============================================================================


def validate_pdf(pdf_path: str) -> Tuple[bool, Optional[str]]:
    try:
        if not os.path.exists(pdf_path):
            return False, "File does not exist"

        # A minimal valid PDF (empty page, no content streams) is ~67 bytes.
        # 100 bytes is a safe lower bound that rejects truncated/corrupt output.
        _MIN_PDF_BYTES = 100

        file_size = os.path.getsize(pdf_path)
        if file_size == 0:
            return False, "File is empty (0 bytes)"
        if file_size < _MIN_PDF_BYTES:
            return False, f"File too small ({file_size} bytes, minimum {_MIN_PDF_BYTES})"

        try:
            from pypdf import PdfReader  # type: ignore

            reader = PdfReader(pdf_path)
            if len(reader.pages) == 0:
                return False, "PDF has no pages"
            get_logger().debug(
                f"PDF validation passed: {os.path.basename(pdf_path)} ({len(reader.pages)} pages, {file_size} bytes)"
            )
            return True, None
        except ImportError:
            get_logger().debug(f"PDF basic validation passed: {os.path.basename(pdf_path)} ({file_size} bytes)")
            return True, None
    except Exception as e:
        return False, str(e)


# ============================================================================
# BACKENDS
# ============================================================================


@functools.lru_cache(maxsize=1)
def is_word_available() -> bool:
    if not HAS_PYWIN32:
        return False
    try:
        app = win32com.client.Dispatch("Word.Application")
        app.Quit()
        return True
    except Exception as e:
        get_logger().debug(f"Word COM not available: {e}")
        return False


def _export_with_word_app(
    word: Any,
    docx_path: str,
    pdf_path: str,
    create_bookmarks: str,
    optimize_for_print: bool,
    pdfa: bool,
) -> None:
    """Export one document using an already-open Word.Application instance."""
    wdExportFormatPDF = 17
    wdExportOptimizeForPrint = 0
    wdExportOptimizeForOnScreen = 1
    wdExportAllDocument = 0
    wdExportDocumentContent = 0
    wdExportCreateNoBookmarks = 0
    wdExportCreateHeadingBookmarks = 1
    wdExportCreateWordBookmarks = 2

    optimize = wdExportOptimizeForPrint if optimize_for_print else wdExportOptimizeForOnScreen

    if create_bookmarks == "headings":
        cb = wdExportCreateHeadingBookmarks
    elif create_bookmarks == "word":
        cb = wdExportCreateWordBookmarks
    else:
        cb = wdExportCreateNoBookmarks

    doc = None
    try:
        get_logger().debug(f"Opening document: {docx_path}")
        doc = word.Documents.Open(docx_path, ReadOnly=True, AddToRecentFiles=False)
        doc.ExportAsFixedFormat(
            OutputFileName=pdf_path,
            ExportFormat=wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=optimize,
            Range=wdExportAllDocument,
            Item=wdExportDocumentContent,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=cb,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=bool(pdfa),
        )
        get_logger().debug(f"Word export completed: {pdf_path}")
    except Exception as e:
        raise RuntimeError(f"Word conversion error: {e}") from e
    finally:
        try:
            if doc:
                doc.Close(False)
        except Exception:
            pass


def convert_with_word(
    docx_path: str,
    pdf_path: str,
    create_bookmarks: str,
    optimize_for_print: bool,
    pdfa: bool,
) -> None:
    """Convert a single file, managing the Word application lifecycle."""
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        _export_with_word_app(word, docx_path, pdf_path, create_bookmarks, optimize_for_print, pdfa)
    finally:
        try:
            if word:
                word.Quit()
        except Exception:
            pass


def convert_with_libreoffice(docx_path: str, pdf_path: str) -> None:
    soffice = shutil.which("soffice") or shutil.which("soffice.bin") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("LibreOffice not found. Please install LibreOffice and ensure 'soffice' is in PATH.")

    outdir = os.path.dirname(pdf_path) or os.getcwd()
    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf:writer_pdf_Export",
        "--outdir",
        outdir,
        docx_path,
    ]
    get_logger().debug(f"Running LibreOffice: {' '.join(cmd)}")

    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        if proc.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed (exit code {proc.returncode}):\n{proc.stderr or proc.stdout}"
            )

        produced_pdf = os.path.join(outdir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
        if os.path.abspath(produced_pdf) != os.path.abspath(pdf_path):
            if not os.path.exists(produced_pdf):
                raise RuntimeError(f"LibreOffice output not found: {produced_pdf}")
            os.replace(produced_pdf, pdf_path)
        get_logger().debug(f"LibreOffice conversion completed: {pdf_path}")
    except subprocess.TimeoutExpired as e:
        raise RuntimeError("LibreOffice conversion timed out after 300 seconds") from e


# ============================================================================
# RUNNER (single + batch)
# ============================================================================


def convert_single_file(
    docx_path: str,
    pdf_path: str,
    backend: str,
    config: ConversionConfig,
    word_app: Optional[Any] = None,
) -> ConversionResult:
    metadata_injected = False
    validation_passed = False
    try:
        os.makedirs(os.path.dirname(pdf_path) or ".", exist_ok=True)

        if backend == "word":
            if word_app is not None:
                _export_with_word_app(
                    word_app,
                    docx_path=docx_path,
                    pdf_path=pdf_path,
                    create_bookmarks=config.bookmarks,
                    optimize_for_print=True,
                    pdfa=config.pdfa,
                )
            else:
                convert_with_word(
                    docx_path=docx_path,
                    pdf_path=pdf_path,
                    create_bookmarks=config.bookmarks,
                    optimize_for_print=True,
                    pdfa=config.pdfa,
                )
            metadata_injected = True
        elif backend == "libreoffice":
            convert_with_libreoffice(docx_path, pdf_path)
            meta = read_docx_core_properties(docx_path)
            if meta:
                metadata_injected = inject_pdf_metadata(pdf_path, meta)
        else:
            raise ValueError(f"Unknown backend: {backend}")

        if config.validate_pdf:
            is_valid, val_error = validate_pdf(pdf_path)
            if not is_valid:
                raise RuntimeError(f"PDF validation failed: {val_error}")
            validation_passed = True

        return ConversionResult(
            docx_path=docx_path,
            pdf_path=pdf_path,
            success=True,
            backend_used=backend,
            metadata_injected=metadata_injected,
            validation_passed=validation_passed,
        )
    except Exception as e:
        return ConversionResult(
            docx_path=docx_path,
            pdf_path=pdf_path,
            success=False,
            backend_used=backend,
            error_message=str(e),
            traceback=traceback.format_exc(),
            metadata_injected=metadata_injected,
            validation_passed=validation_passed,
        )


def resolve_backend(config: ConversionConfig) -> Tuple[str, bool]:
    word_available = is_word_available()
    backend = config.backend
    logger = get_logger()

    if backend == "auto":
        backend = "word" if word_available else "libreoffice"
        logger.info(f"Auto-selected backend: {backend}")

    if backend == "word" and not word_available:
        logger.warning("Microsoft Word not available, falling back to LibreOffice")
        backend = "libreoffice"

    if backend == "libreoffice" and config.pdfa:
        logger.warning("PDF/A is only supported with Word backend; ignoring --pdfa for LibreOffice")

    return backend, word_available


def _log_result(res: ConversionResult, root: str, logger: logging.Logger) -> None:
    rel_path = os.path.relpath(res.docx_path, root)
    if res.success:
        logger.info(f"OK   | {rel_path} ({res.backend_used})")
    else:
        logger.error(f"FAIL | {rel_path}: {res.error_message}")
        logger.debug(res.traceback or "")


def convert_batch(files: List[str], config: ConversionConfig, backend: str) -> List[ConversionResult]:
    logger = get_logger()
    results: List[ConversionResult] = []
    root = os.path.abspath(config.root_dir)

    tasks: List[Tuple[str, str]] = []
    for docx_path in files:
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
        if os.path.exists(pdf_path) and not config.overwrite:
            rel_path = os.path.relpath(docx_path, root)
            logger.info(f"SKIP | {rel_path} (PDF already exists)")
            continue
        tasks.append((docx_path, pdf_path))

    if not tasks:
        logger.warning("No files to convert")
        return results

    effective_workers = config.workers
    if backend == "word" and effective_workers > 1:
        logger.info("Word backend: forcing workers=1 for stability")
        effective_workers = 1

    logger.info(f"Converting {len(tasks)} file(s) using {effective_workers} worker(s)")

    # Parallel LibreOffice conversions
    if backend != "word" and effective_workers > 1:
        pbar = None
        if HAS_TQDM:
            pbar = tqdm(total=len(tasks), desc="Converting", unit="file")
        with ProcessPoolExecutor(max_workers=effective_workers) as executor:
            futures = [executor.submit(convert_single_file, docx, pdf, backend, config) for docx, pdf in tasks]
            for fut in as_completed(futures):
                res = fut.result()
                results.append(res)
                _log_result(res, root, logger)
                if pbar is not None:
                    pbar.update(1)
        if pbar is not None:
            pbar.close()
        return results

    # Sequential processing — Word: open one app instance for the whole batch
    word_app = None
    if backend == "word" and HAS_PYWIN32:
        try:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0
            logger.debug(f"Word.Application opened for batch of {len(tasks)} file(s)")
        except Exception as e:
            logger.warning(f"Failed to pre-open Word.Application, will open per file: {e}")

    iterator: Any = tasks
    pbar2 = None
    if HAS_TQDM:
        pbar2 = tqdm(tasks, desc="Converting", unit="file")
        iterator = pbar2
    try:
        for docx_path, pdf_path in iterator:
            res = convert_single_file(docx_path, pdf_path, backend, config, word_app)
            results.append(res)
            _log_result(res, root, logger)
    finally:
        if pbar2 is not None:
            pbar2.close()
        if word_app is not None:
            try:
                word_app.Quit()
            except Exception:
                pass
    return results


def print_summary(results: List[ConversionResult], backend: str) -> None:
    logger = get_logger()

    total = len(results)
    successes = sum(1 for r in results if r.success)
    failures = total - successes

    logger.info("=" * 60)
    logger.info("CONVERSION SUMMARY")
    logger.info("=" * 60)
    logger.info(f"Backend used:       {backend}")
    logger.info(f"Total files:        {total}")
    logger.info(f"Successful:         {successes}")
    logger.info(f"Failed:             {failures}")

    if results and successes > 0:
        metadata_count = sum(1 for r in results if r.success and r.metadata_injected)
        validated_count = sum(1 for r in results if r.success and r.validation_passed)
        logger.info(f"Metadata injected:  {metadata_count}/{successes}")
        logger.info(f"Validated:          {validated_count}/{successes}")
    logger.info("=" * 60)

    if failures > 0:
        logger.warning(f"\n{failures} file(s) failed to convert:")
        for r in results:
            if not r.success:
                logger.warning(f"  - {os.path.basename(r.docx_path)}: {r.error_message}")


# ============================================================================
# CLI + GUI
# ============================================================================


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Batch DOCX to PDF converter with optional parallelism and validation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python docx_to_pdf.py --dir . --recursive --workers 4
  python docx_to_pdf.py --config config.json --workers 8
  python docx_to_pdf.py --use word --bookmarks headings
        """,
    )

    parser.add_argument("--config", help="Load configuration from JSON file")
    parser.add_argument("--save-config", help="Save resulting configuration to JSON file and exit")

    # Use SUPPRESS so we can detect what the user explicitly passed and apply precedence over --config.
    parser.add_argument("--dir", dest="root_dir", default=argparse.SUPPRESS, help="Root directory (default: .)")
    parser.add_argument("--recursive", action="store_true", default=argparse.SUPPRESS, help="Recurse into subfolders")
    parser.add_argument("--overwrite", action="store_true", default=argparse.SUPPRESS, help="Overwrite existing PDFs")

    parser.add_argument("--pdfa", action="store_true", default=argparse.SUPPRESS, help="Export as PDF/A-1 (Word only)")
    parser.add_argument(
        "--bookmarks",
        choices=[m.value for m in BookmarkMode],
        default=argparse.SUPPRESS,
        help="Bookmark source (default: headings) when using Word",
    )
    parser.add_argument(
        "--use",
        dest="backend",
        choices=[b.value for b in Backend],
        default=argparse.SUPPRESS,
        help="Force a backend (default: auto)",
    )

    parser.add_argument(
        "--workers",
        type=int,
        default=argparse.SUPPRESS,
        help="Number of workers (default: 1; Word is forced to 1)",
    )
    parser.add_argument(
        "--no-validate",
        dest="validate_pdf",
        action="store_false",
        default=argparse.SUPPRESS,
        help="Skip PDF validation after conversion",
    )

    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default=argparse.SUPPRESS,
        help="Logging level (default: INFO)",
    )
    parser.add_argument("--log-file", default=argparse.SUPPRESS, help="Write logs to file")

    return parser


def build_config_from_args(args: argparse.Namespace) -> ConversionConfig:
    if getattr(args, "config", None):
        base = ConversionConfig.from_file(args.config)
    else:
        base = ConversionConfig()

    overrides: Dict[str, Any] = {}
    for f in dataclasses.fields(ConversionConfig):
        if hasattr(args, f.name):
            overrides[f.name] = getattr(args, f.name)

    cfg = dataclasses.replace(base, **overrides)
    cfg.validate()
    return cfg


def run_conversion(config: ConversionConfig) -> Tuple[List[ConversionResult], str]:
    logger = get_logger()

    root = os.path.abspath(config.root_dir)
    if not os.path.isdir(root):
        raise FileNotFoundError(f"Directory not found: {root}")

    files = find_docx_files(root, config.recursive)
    if not files:
        logger.warning("No DOCX files found")
        return [], config.backend

    backend, _ = resolve_backend(config)

    logger.info(f"Root directory:     {root}")
    logger.info(f"Recursive:          {config.recursive}")
    logger.info(f"Workers:            {config.workers}")
    logger.info(f"Backend:            {backend}")
    logger.info(f"Validate PDFs:      {config.validate_pdf}")
    logger.info(f"Files found:        {len(files)}")

    start = datetime.now()
    results = convert_batch(files, config, backend)
    end = datetime.now()

    print_summary(results, backend)

    duration = (end - start).total_seconds()
    logger.info(f"Total time:         {duration:.2f} seconds")
    success_count = sum(1 for r in results if r.success)
    if success_count > 0:
        logger.info(f"Average per file:   {duration / success_count:.2f} seconds")

    return results, backend


def run_gui_mode() -> int:
    """
    Double-click friendly mode.

    If tkinter is available, show folder picker and two yes/no prompts.
    If not available, fall back to minimal console prompts.
    """
    try:
        import tkinter as tk  # noqa: PLC0415
        from tkinter import filedialog, messagebox  # noqa: PLC0415

        root = tk.Tk()
        root.withdraw()
        try:
            root.attributes("-topmost", True)
        except Exception:
            pass

        try:
            folder = filedialog.askdirectory(title="Select folder containing DOCX files")
            if not folder:
                return 0

            recursive = messagebox.askyesno("DOCX to PDF", "Search subfolders too?", default="yes")
            overwrite = messagebox.askyesno("DOCX to PDF", "Overwrite existing PDFs?", default="no")

            log_file = os.path.join(folder, "conversion.log")
            cfg = ConversionConfig(
                root_dir=folder,
                recursive=bool(recursive),
                overwrite=bool(overwrite),
                backend="auto",
                workers=min(4, os.cpu_count() or 4),
                validate_pdf=True,
                log_level="INFO",
                log_file=log_file,
            )
            cfg.validate()

            setup_logging(cfg.log_level, cfg.log_file)
            get_logger().info("Running in double-click mode")

            results, backend_used = run_conversion(cfg)
            total = len(results)
            ok = sum(1 for r in results if r.success)
            fail = total - ok
            messagebox.showinfo(
                "DOCX to PDF",
                f"Done.\nBackend: {backend_used}\nTotal: {total}\nOK: {ok}\nFAIL: {fail}\n\nLog: {log_file}",
            )
            return 0 if fail == 0 else 1
        except Exception as e:
            # Try to write details to the log if possible.
            try:
                get_logger().error(str(e))
                get_logger().debug(traceback.format_exc())
            except Exception:
                pass
            messagebox.showerror("DOCX to PDF", f"Error:\n{e}")
            return 1
        finally:
            try:
                root.destroy()
            except Exception:
                pass
    except Exception:
        # Console fallback (e.g. tkinter not installed/available)
        print("DOCX to PDF (no GUI available).")
        folder = input("Folder containing DOCX files (blank to cancel): ").strip()
        if not folder:
            return 0
        recursive_in = input("Search subfolders too? [Y/n]: ").strip().lower()
        recursive = recursive_in != "n"
        overwrite_in = input("Overwrite existing PDFs? [y/N]: ").strip().lower()
        overwrite = overwrite_in == "y"

        log_file = os.path.join(folder, "conversion.log")
        cfg = ConversionConfig(
            root_dir=folder,
            recursive=recursive,
            overwrite=overwrite,
            backend="auto",
            workers=min(4, os.cpu_count() or 4),
            validate_pdf=True,
            log_level="INFO",
            log_file=log_file,
        )
        cfg.validate()

        setup_logging(cfg.log_level, cfg.log_file)
        get_logger().info("Running in double-click mode (console fallback)")

        results, _ = run_conversion(cfg)
        failures = sum(1 for r in results if not r.success)
        return 0 if failures == 0 else 1


def main(argv: Optional[List[str]] = None) -> int:
    if argv is None:
        argv = sys.argv[1:]

    # No-args mode: double-click UX.
    if len(argv) == 0:
        return run_gui_mode()

    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        cfg = build_config_from_args(args)
    except Exception as e:
        setup_logging("INFO", None)
        get_logger().error(f"Invalid configuration: {e}")
        return 2

    if getattr(args, "save_config", None):
        try:
            cfg.to_file(args.save_config)
            setup_logging(cfg.log_level, cfg.log_file)
            get_logger().info(f"Configuration saved to {args.save_config}")
            return 0
        except Exception as e:
            setup_logging("INFO", None)
            get_logger().error(f"Failed to save config file: {e}")
            return 1

    setup_logging(cfg.log_level, cfg.log_file)
    logger = get_logger()
    logger.info("=" * 60)
    logger.info("DOCX to PDF Batch Converter")
    logger.info("=" * 60)

    try:
        results, _ = run_conversion(cfg)
        failures = sum(1 for r in results if not r.success)
        return 0 if failures == 0 else 1
    except FileNotFoundError as e:
        logger.error(str(e))
        return 1
    except Exception as e:
        logger.error(str(e))
        logger.debug(traceback.format_exc())
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

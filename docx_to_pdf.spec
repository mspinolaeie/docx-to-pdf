# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for docx-to-pdf (Windows x64, --onefile)
#
# Builds two executables:
#   docx-to-pdf.exe     — console build  (CLI / scripting)
#   docx-to-pdf-gui.exe — windowed build (double-click, no terminal)
#
# Build:
#   pip install pyinstaller pyinstaller-hooks-contrib
#   pyinstaller docx_to_pdf.spec --clean --noconfirm

from PyInstaller.utils.hooks import collect_submodules

# pywin32 COM dispatch loads modules at runtime; list them explicitly so
# PyInstaller bundles the necessary pyd files.
_pywin32_hidden = [
    "win32com",
    "win32com.client",
    "win32com.client.dynamic",
    "win32com.client.gencache",
    "pywintypes",
    "win32api",
    "win32con",
]

_hidden = _pywin32_hidden + ["docx_to_pdf_gui"] + collect_submodules("pypdf") + collect_submodules("PySide6")

# ── Analysis (shared) ────────────────────────────────────────────────────────

a = Analysis(
    ["docx_to_pdf.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest", "_pytest", "pygments", "matplotlib", "numpy"],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

# ── Console build ────────────────────────────────────────────────────────────

exe_cli = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="docx-to-pdf",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    target_arch="x86_64",
)

# ── Windowed (GUI double-click) build ────────────────────────────────────────

exe_gui = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="docx-to-pdf-gui",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    target_arch="x86_64",
)

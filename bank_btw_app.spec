# PyInstaller-spec voor Kwartaal (pywebview + WebView2).
# Bouwen:  python -m PyInstaller --noconfirm bank_btw_app.spec
# Resultaat: dist/Kwartaal.exe  (1 bestand, dubbelklikken)
from PyInstaller.utils.hooks import collect_all

datas = [("webui", "webui")]           # de hele web-UI mee in de exe
binaries = []
hiddenimports = [
    "webview.platforms.edgechromium",   # WebView2-backend (lazy geladen)
    "clr",                              # pythonnet, nodig voor edgechromium
    # OCR/bonnen worden pas bij gebruik geimporteerd -> expliciet meenemen:
    "pytesseract", "pdfplumber", "fitz", "PIL", "PIL.Image",
    # eigen modules die dynamisch geladen kunnen worden:
    "boekhoud_export", "bonnen_matcher",
]

# Pakketten met eigen data/DLL's (WebView2-DLL's, pdfplumber-tabellen, PyMuPDF-libs).
for pkg in ("webview", "pdfplumber", "fitz"):
    try:
        d, b, h = collect_all(pkg)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception:
        pass

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter"],  # de app gebruikt geen tkinter (dat is de oude gui.py)
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Kwartaal",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,   # geen zwart terminalvenster
    icon="webui/app-icon.ico",
)

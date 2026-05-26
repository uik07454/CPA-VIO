# CPAuto.spec
# PyInstaller spec file for the CPAuto application.
#
# Build command:
#   pyinstaller CPAuto.spec
#
# Output: dist/CPAuto.exe  (single-file bundle)
#
# Files that must be distributed alongside the EXE (NOT bundled inside):
#   assets/CPA_template/*.xlsx   - Excel templates (user-selectable, may grow)
#   assets/splash_assets/splash.mp4 - Splash video
#   data/config.json             - User settings (read/write at runtime)
#   output/                      - Created automatically at runtime

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Collect hidden imports that PyInstaller may miss
hiddenimports = (
    collect_submodules("lxml")
    + collect_submodules("openpyxl")
    + collect_submodules("PIL")
    + [
        "PyQt6.QtCore",
        "PyQt6.QtGui",
        "PyQt6.QtWidgets",
        "PyQt6.QtMultimedia",
        "PyQt6.QtMultimediaWidgets",
        "PyQt6.QtNetwork",
        "openpyxl.cell.rich_text",
        "openpyxl.drawing.image",
        "openpyxl.drawing.spreadsheet_drawing",
        "openpyxl.utils",
        "openpyxl.utils.units",
        "openpyxl.styles",
        "difflib",
        "zipfile",
        "copy",
        "re",
    ]
)

# Unused PyQt6 modules — significantly reduces bundle size
_qt_excludes = [
    "PyQt6.QtBluetooth",
    "PyQt6.QtDBus",
    "PyQt6.QtDesigner",
    "PyQt6.QtHelp",
    "PyQt6.QtLocation",
    "PyQt6.QtNfc",
    "PyQt6.QtOpenGL",
    "PyQt6.QtOpenGLWidgets",
    "PyQt6.QtPositioning",
    "PyQt6.QtPrintSupport",
    "PyQt6.QtQml",
    "PyQt6.QtQuick",
    "PyQt6.QtQuickWidgets",
    "PyQt6.QtRemoteObjects",
    "PyQt6.QtSensors",
    "PyQt6.QtSerialPort",
    "PyQt6.QtSql",
    "PyQt6.QtSvg",
    "PyQt6.QtSvgWidgets",
    "PyQt6.QtTest",
    "PyQt6.QtWebChannel",
    "PyQt6.QtWebEngineCore",
    "PyQt6.QtWebEngineQuick",
    "PyQt6.QtWebEngineWidgets",
    "PyQt6.QtWebSockets",
    "PyQt6.QtXml",
    "PyQt6.Qt3DAnimation",
    "PyQt6.Qt3DCore",
    "PyQt6.Qt3DExtras",
    "PyQt6.Qt3DInput",
    "PyQt6.Qt3DLogic",
    "PyQt6.Qt3DRender",
    "PyQt6.QtCharts",
    "PyQt6.QtDataVisualization",
    "PyQt6.QtPdf",
    "PyQt6.QtPdfWidgets",
    "PyQt6.QtSpatialAudio",
    "PyQt6.QtStateMachine",
    "PyQt6.QtTextToSpeech",
]

a = Analysis(
    ["../gui.py"],
    pathex=[".."],
    binaries=[],
    datas=(
        collect_data_files("lxml")
        + collect_data_files("openpyxl")
    ),
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=(
        _qt_excludes
        + [
            # Heavy dev/science packages not needed at runtime
            "manim",
            "numpy",
            "matplotlib",
            "scipy",
            "tkinter",
            "unittest",
            "pytest",
            "IPython",
            "jupyter",
            "notebook",
            "setuptools",
            "pkg_resources",
            "distutils",
            "email",
            "html",
            "http",
            "xmlrpc",
            "ftplib",
            "imaplib",
            "poplib",
            "smtplib",
            "telnetlib",
            "nntplib",
            "pydoc",
            "doctest",
            "pdb",
            "profile",
            "cProfile",
            "timeit",
            "trace",
            "turtle",
            "curses",
        ]
    ),
    noarchive=False,
)

pyz = PYZ(a.pure)

# One-file EXE: everything packed into a single executable
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="CPAuto",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # no console window (GUI app)
    icon="../assets/DocGear.ico",
)

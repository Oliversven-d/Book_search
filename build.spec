# build.spec
# PyInstaller spec file for creating an executable

# Import the Analysis, PYZ, EXE and collect functions from PyInstaller
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT
from PyInstaller.config import CONF
from PyInstaller import log as logging

# Configuration
log = logging.getLogger(__name__)
block_cipher = None

# Analysis
a = Analysis(
    ['book_search.py'],  # Main script file
    pathex=[],
    binaries=[],
    datas=collect_data_files('nltk') + collect_data_files('janome') + collect_data_files('sklearn') + collect_data_files('gspread') + collect_data_files('oauth2client') + collect_data_files('PIL'),
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# PYZ (Python code archive)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# EXE (Executable)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='book_search_app',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Set to False if you want to hide the console window
)

# COLLECT (Package everything into one folder)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='book_search_app',
)

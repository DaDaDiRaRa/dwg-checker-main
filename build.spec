# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_data_files

datas, binaries, hiddenimports = [], [], []

# customtkinter — 테마/이미지 에셋 포함
for pkg in ('customtkinter', 'tkinterdnd2'):
    d, b, h = collect_all(pkg)
    datas += d; binaries += b; hiddenimports += h

# ezdxf — 폰트, 리소스 데이터
datas += collect_data_files('ezdxf')

# pdfminer — cmap 파일 (pdfplumber 의존)
datas += collect_data_files('pdfminer')

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports + [
        'ezdxf',
        'ezdxf.addons',
        'pdfplumber',
        'openpyxl',
        'openpyxl.styles',
        'pandas',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='도면검토기',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

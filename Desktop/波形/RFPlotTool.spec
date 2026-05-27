# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['rf_plot_ui.py'],
    pathex=[],
    binaries=[],
    datas=[('webui.html', '.'), ('dist/updater.exe', '.')],
    hiddenimports=['webview.platforms.edgechromium', 'clr_loader', 'clr_loader.ffi'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', '_tkinter', 'PyQt5', 'PySide6'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='RFPlotTool',
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
)

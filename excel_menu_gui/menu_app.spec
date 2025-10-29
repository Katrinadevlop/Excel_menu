# -*- mode: python ; coding: utf-8 -*-
import os
from pathlib import Path

# Определяем базовый путь
base_path = Path('.')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Включаем папку templates со всеми файлами
        ('templates', 'templates'),
        # Включаем любые другие важные файлы
        ('requirements.txt', '.'),
    ],
    hiddenimports=[
        # PySide6 модули
        'PySide6.QtCore',
        'PySide6.QtGui', 
        'PySide6.QtWidgets',
        # Excel библиотеки
        'openpyxl',
        'xlrd',
        'xlwings',
        # PowerPoint
        'pptx',
        'python-pptx',
        # Другие модули
        'PIL',
        'Pillow',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='MenuApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Отключаем консоль для GUI приложения
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # Красивая иконка приложения
    version=None,
)

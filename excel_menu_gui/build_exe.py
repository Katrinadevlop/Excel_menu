#!/usr/bin/env python3
"""
Скрипт для создания exe файла приложения "Работа с меню"
"""

import os
import sys
import subprocess
from pathlib import Path

def main():
    print("🚀 Начинаем сборку exe файла...")
    
    # Проверяем, что мы в правильной директории
    if not Path("main.py").exists():
        print("❌ Ошибка: файл main.py не найден в текущей директории!")
        return False
    
    # Проверяем PyInstaller
    print("📦 Проверяем PyInstaller...")
    try:
        import PyInstaller
        print("✅ PyInstaller найден")
    except ImportError:
        print("❌ PyInstaller не установлен!")
        print("💡 Установите его командой: pip install pyinstaller")
        return False
    
    # Создаем директорию для сборки, если её нет
    build_dir = Path("build")
    dist_dir = Path("dist")
    
    # Команда PyInstaller
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                    # Один exe файл
        "--windowed",                   # Без консоли (GUI приложение)
        "--name=MenuApp",               # Имя exe файла
        "--icon=app_icon.ico",          # Иконка (если есть)
        "--add-data=templates;templates",  # Добавляем папку templates
        "--hidden-import=openpyxl",     # Явно включаем openpyxl
        "--hidden-import=xlrd",         # Явно включаем xlrd
        "--hidden-import=PySide6",      # Явно включаем PySide6
        "--collect-all=PySide6",        # Собираем все модули PySide6
        "main.py"                       # Главный файл
    ]
    
    print("🔧 Запускаем PyInstaller...")
    print(f"Команда: {' '.join(cmd)}")
    
    try:
        # Запускаем PyInstaller
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✅ Сборка завершена успешно!")
        
        # Проверяем результат
        exe_path = dist_dir / "MenuApp.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"📁 Exe файл создан: {exe_path.absolute()}")
            print(f"📊 Размер: {size_mb:.1f} MB")
            return True
        else:
            print("❌ Exe файл не найден после сборки")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"❌ Ошибка сборки: {e}")
        if e.stdout:
            print("Вывод:", e.stdout)
        if e.stderr:
            print("Ошибки:", e.stderr)
        return False

def create_spec_file():
    """Создает spec файл для более точной настройки сборки"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('templates', 'templates')],
    hiddenimports=[
        'openpyxl',
        'xlrd', 
        'PySide6.QtCore',
        'PySide6.QtGui', 
        'PySide6.QtWidgets',
        'comparator',
        'presentation_handler',
        'template_linker',
        'theme'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

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
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico'
)
'''
    
    with open("MenuApp.spec", "w", encoding="utf-8") as f:
        f.write(spec_content)
    
    print("✅ Создан файл MenuApp.spec")

def create_icon():
    """Создает простую иконку для приложения"""
    try:
        from PIL import Image, ImageDraw, ImageFont
        
        # Создаем изображение 256x256
        size = 256
        img = Image.new('RGBA', (size, size), (255, 126, 95, 255))
        draw = ImageDraw.Draw(img)
        
        # Рисуем круг
        margin = 12
        draw.ellipse([margin, margin, size-margin, size-margin], 
                    fill=(253, 58, 105, 255), outline=(255, 255, 255, 230), width=6)
        
        # Добавляем букву М
        try:
            # Пытаемся использовать системный шрифт
            font = ImageFont.truetype("arial.ttf", 120)
        except:
            # Если не получилось, используем стандартный
            font = ImageFont.load_default()
        
        # Рисуем текст
        bbox = draw.textbbox((0, 0), "М", font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        x = (size - text_width) // 2
        y = (size - text_height) // 2 - 10
        
        draw.text((x, y), "М", fill=(255, 255, 255, 255), font=font)
        
        # Сохраняем как ICO
        img.save("app_icon.ico", format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
        print("✅ Создана иконка app_icon.ico")
        return True
        
    except ImportError:
        print("⚠️ Pillow не установлен, создаем exe без иконки")
        return False
    except Exception as e:
        print(f"⚠️ Не удалось создать иконку: {e}")
        return False

if __name__ == "__main__":
    print("=" * 50)
    print("🏗️  Сборка приложения 'Работа с меню' в exe")
    print("=" * 50)
    
    # Создаем иконку
    create_icon()
    
    # Создаем spec файл  
    create_spec_file()
    
    # Собираем exe
    success = main()
    
    if success:
        print("\n" + "=" * 50)
        print("🎉 Готово! Exe файл создан в папке dist/")
        print("📁 Путь: dist/MenuApp.exe")
        print("=" * 50)
    else:
        print("\n" + "=" * 50)
        print("❌ Сборка не удалась. Проверьте ошибки выше.")
        print("=" * 50)

@echo off
chcp 65001 >nul
echo ========================================
echo 📦 Установка библиотек для проекта
echo ========================================
echo.

REM Проверяем Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python не найден! Установите Python 3.8+ 
    pause
    exit /b 1
)

echo ✅ Python найден
echo.
echo 📥 Устанавливаем необходимые библиотеки...
echo.

REM Установка библиотек по одной для лучшего контроля
echo 🔹 Устанавливаем PySide6...
pip install PySide6>=6.4.0
if errorlevel 1 echo ⚠️  Ошибка установки PySide6

echo 🔹 Устанавливаем openpyxl...
pip install openpyxl>=3.1.0
if errorlevel 1 echo ⚠️  Ошибка установки openpyxl

echo 🔹 Устанавливаем xlrd...
pip install xlrd==1.2.0
if errorlevel 1 echo ⚠️  Ошибка установки xlrd

echo 🔹 Устанавливаем python-pptx...
pip install python-pptx>=0.6.21
if errorlevel 1 echo ⚠️  Ошибка установки python-pptx

echo 🔹 Устанавливаем PyInstaller...
pip install PyInstaller>=5.0
if errorlevel 1 echo ⚠️  Ошибка установки PyInstaller

echo 🔹 Устанавливаем Pillow (для иконки)...
pip install Pillow>=9.0.0
if errorlevel 1 echo ⚠️  Ошибка установки Pillow

echo.
echo ✅ Установка завершена!
echo 💡 Теперь можно запустить build.bat для сборки exe
echo.
pause

@echo off
chcp 65001 >nul
echo 🏗️ Простая сборка exe файла
echo.

REM Прямая команда PyInstaller
python -m PyInstaller ^
--onefile ^
--windowed ^
--name=MenuApp ^
--add-data=templates;templates ^
--hidden-import=openpyxl ^
--hidden-import=xlrd ^
--hidden-import=PySide6 ^
--collect-all=PySide6 ^
main.py

echo.
if exist "dist\MenuApp.exe" (
    echo ✅ Готово! Файл: dist\MenuApp.exe
) else (
    echo ❌ Что-то пошло не так
)
echo.
pause

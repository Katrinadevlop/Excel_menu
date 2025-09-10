@echo off
chcp 65001 >nul
echo ========================================
echo 🏗️  Сборка приложения "Работа с меню"
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
echo 💡 ВАЖНО: Убедитесь что установлены библиотеки:
echo    pip install PySide6 openpyxl xlrd python-pptx pyinstaller
echo.
echo 📦 Запускаем сборку...
echo.

REM Запускаем скрипт сборки
python build_exe.py

echo.
echo 🔄 Сборка завершена. Проверьте папку dist/
pause

@echo off
chcp 65001 >nul
echo.
echo ===================================
echo Тестирование Excel файла с меню
echo ===================================
echo.

python test_categories.py

echo.
echo Нажмите любую клавишу для выхода...
pause >nul

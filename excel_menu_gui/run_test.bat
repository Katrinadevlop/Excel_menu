@echo off
chcp 65001 > nul
echo Запуск теста извлечения блюд...
echo.
python quick_test.py
echo.
echo Если возникли проблемы, проверьте файл ПРОВЕРКА_ИСПРАВЛЕНИЙ.md
pause

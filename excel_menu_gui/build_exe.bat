@echo off
echo Creating executable for Menu Application...
echo.

REM Активируем виртуальную среду если она есть
if exist ".venv\Scripts\activate.bat" (
    echo Activating virtual environment...
    call .venv\Scripts\activate.bat
) else (
    echo Virtual environment not found, using system Python
)

REM Проверяем установлен ли PyInstaller
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

REM Удаляем старые файлы сборки
if exist "dist" (
    echo Removing old build files...
    rmdir /s /q dist
)
if exist "build" (
    rmdir /s /q build
)

REM Создаем .exe используя spec файл
echo Building executable...
pyinstaller menu_app.spec

if exist "dist\MenuApp.exe" (
    echo.
    echo SUCCESS! Executable created at: dist\MenuApp.exe
    echo.
    echo The executable is ready for installation on another computer.
    echo Copy the entire "dist" folder to the target computer.
    pause
) else (
    echo.
    echo ERROR: Build failed. Check the output above for errors.
    pause
)

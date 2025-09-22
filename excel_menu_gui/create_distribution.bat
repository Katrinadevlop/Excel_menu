@echo off
echo Creating distribution archive...
echo.

set "ARCHIVE_NAME=MenuApp_v1.0_%date:~6,4%-%date:~3,2%-%date:~0,2%.zip"

REM Проверяем наличие папки dist
if not exist "dist" (
    echo ERROR: dist folder not found. Please run build_exe.bat first.
    pause
    exit /b 1
)

REM Создаем архив с помощью PowerShell
echo Creating archive: %ARCHIVE_NAME%
powershell -Command "Compress-Archive -Path 'dist\*' -DestinationPath '%ARCHIVE_NAME%' -Force"

if exist "%ARCHIVE_NAME%" (
    echo.
    echo SUCCESS! Distribution archive created: %ARCHIVE_NAME%
    echo.
    echo Archive contents:
    echo - MenuApp.exe (Main application)
    echo - templates/ folder (Excel and PowerPoint templates)
    echo - install.bat (Installation script)
    echo - README.txt (User instructions)
    echo.
    echo Archive size: 
    for %%A in ("%ARCHIVE_NAME%") do echo %%~zA bytes
    echo.
    echo This archive can now be sent to other computers for installation.
    echo Recipients should extract the archive and run install.bat as administrator.
) else (
    echo.
    echo ERROR: Failed to create archive.
)

pause

@echo off
echo Installing Menu Application...
echo.

REM Создаем папку для приложения
set "INSTALL_DIR=%ProgramFiles%\MenuApp"
if not exist "%INSTALL_DIR%" (
    echo Creating installation directory...
    mkdir "%INSTALL_DIR%" 2>nul
    if errorlevel 1 (
        echo WARNING: Cannot create directory in Program Files.
        set "INSTALL_DIR=%USERPROFILE%\MenuApp"
        echo Installing to user directory: %INSTALL_DIR%
        mkdir "%INSTALL_DIR%" 2>nul
    )
)

REM Копируем exe файл и templates
echo Copying application files...
copy "MenuApp.exe" "%INSTALL_DIR%\" >nul 2>&1
if exist "templates" (
    xcopy /E /I /Y "templates" "%INSTALL_DIR%\templates" >nul 2>&1
)

REM Создаем ярлык на рабочем столе
set "DESKTOP=%USERPROFILE%\Desktop"
set "SHORTCUT=%DESKTOP%\Menu Application.lnk"

echo Creating desktop shortcut...
powershell -Command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%SHORTCUT%'); $Shortcut.TargetPath = '%INSTALL_DIR%\MenuApp.exe'; $Shortcut.WorkingDirectory = '%INSTALL_DIR%'; $Shortcut.Description = 'Menu Application'; $Shortcut.Save()"

if exist "%SHORTCUT%" (
    echo.
    echo Installation completed successfully!
    echo Application installed to: %INSTALL_DIR%
    echo Desktop shortcut created: %SHORTCUT%
    echo.
    echo You can now run the application from the desktop shortcut.
) else (
    echo.
    echo Installation completed, but shortcut creation failed.
    echo You can run the application from: %INSTALL_DIR%\MenuApp.exe
)

echo.
pause

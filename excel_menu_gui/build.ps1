# PowerShell script для сборки exe файла
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "🏗️  Сборка приложения 'Работа с меню'" -ForegroundColor Cyan  
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Проверяем Python
try {
    $pythonVersion = python --version 2>&1
    if ($LASTEXITCODE -eq 0) {
Write-Host "✅ Python найден: $pythonVersion" -ForegroundColor Green
    } else {
        throw "Python не найден"
    }
} catch {
    Write-Host "❌ Python не найден! Установите Python 3.8+" -ForegroundColor Red
    Read-Host "Нажмите Enter для выхода"
    exit 1
}

# Проверяем main.py
if (-not (Test-Path "main.py")) {
    Write-Host "❌ Файл main.py не найден!" -ForegroundColor Red
    Read-Host "Нажмите Enter для выхода"  
    exit 1
}

Write-Host "💡 ВАЖНО: Убедитесь что установлены библиотеки:" -ForegroundColor Yellow
Write-Host "   pip install PySide6 openpyxl xlrd python-pptx pyinstaller" -ForegroundColor Cyan
Write-Host ""
Write-Host "📦 Запускаем сборку..." -ForegroundColor Yellow
Write-Host ""

# Запускаем скрипт сборки
try {
    python build_exe.py
    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "🎉 Сборка завершена успешно!" -ForegroundColor Green
        Write-Host "📁 Проверьте папку dist/MenuApp.exe" -ForegroundColor Green
    } else {
        Write-Host "❌ Сборка завершилась с ошибкой" -ForegroundColor Red
    }
} catch {
    Write-Host "❌ Ошибка выполнения: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Read-Host "Нажмите Enter для выхода"

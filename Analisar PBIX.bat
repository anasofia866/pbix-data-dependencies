@echo off
chcp 65001 > nul

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ❌ Python is not installed!
    echo    Please run first: "1_INSTALAR_PYTHON.bat"
    echo.
    pause
    exit /b 1
)

:: Launch tool
start "" python "%~dp0pbix_analyzer.py"

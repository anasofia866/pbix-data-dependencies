@echo off
chcp 65001 > nul
echo.
echo ╔══════════════════════════════════════════════════╗
echo ║     Installer — PBIX Analyzer                   ║
echo ╚══════════════════════════════════════════════════╝
echo.

:: Check if Python is already installed
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ Python is already installed!
    python --version
    goto install_deps
)

echo Python not found. Installing via Windows Store (winget)...
echo.
echo WARNING: An installation window may appear. Please wait.
echo.

winget install Python.Python.3.12 --accept-source-agreements --accept-package-agreements

if %errorlevel% neq 0 (
    echo.
    echo ❌ Automatic installation failed.
    echo.
    echo Please install Python manually:
    echo   1. Open Microsoft Store
    echo   2. Search for "Python 3.12"
    echo   3. Install it
    echo   4. Run this file again
    echo.
    pause
    exit /b 1
)

echo.
echo ✅ Python installed successfully!
echo    (you may need to close and reopen this window)
echo.

:install_deps
echo Installing required libraries...
python -m pip install openpyxl --quiet --upgrade
if %errorlevel% neq 0 (
    echo.
    echo ⚠️  Problem installing openpyxl. Trying with pip...
    pip install openpyxl
)

echo.
echo ════════════════════════════════════════════════════
echo  ✅ INSTALLATION COMPLETE!
echo  You can now use: "Analisar PBIX.bat"
echo ════════════════════════════════════════════════════
echo.
pause

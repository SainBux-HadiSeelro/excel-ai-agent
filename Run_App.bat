@echo off
title Excel AI Agent - 1-Click Launcher
color 0A
cd /d "%~dp0"

echo ========================================================
echo        📊 Excel AI Agent - Autonomous Spreadsheet Engine
echo ========================================================
echo.
echo [1/3] Checking Python environment...

:: Check if virtual environment exists
if exist ".venv\Scripts\activate.bat" (
    echo [OK] Using virtual environment (.venv)...
    call ".venv\Scripts\activate.bat"
) else (
    echo [INFO] Virtual environment not found, checking system Python...
    where python >nul 2>nul
    if %ERRORLEVEL% NEQ 0 (
        color 0C
        echo [ERROR] Python install nahi hai! 
        echo Please install Python 3.10+ from https://www.python.org/
        echo Make sure to check "Add Python to PATH" during installation.
        echo.
        pause
        exit /b 1
    )
)

echo [2/3] Checking dependencies...
python -c "import streamlit, openpyxl, formulas" >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [INFO] Installing required packages (first time setup)...
    pip install -r requirements.txt
    if %ERRORLEVEL% NEQ 0 (
        color 0C
        echo [ERROR] Dependencies install karte waqt error aaya.
        pause
        exit /b 1
    )
)

echo [3/3] Starting Excel AI Copilot Web App...
echo.
echo ========================================================
echo  🚀 App aapke browser mein open ho rahi hai...
echo  (Is window ko band mat karein jab tak aap app use kar rahe hain)
echo ========================================================
echo.

python -m streamlit run app.py --server.headless=false

if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo.
    echo [ERROR] App run hone mein masla hua.
    pause
)

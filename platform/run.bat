@echo off
echo ============================================
echo   CPMS Evaluation Platform - Setup ^& Launch
echo ============================================
echo.

cd /d "%~dp0"

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python 3.9+ from https://www.python.org
    pause
    exit /b 1
)

:: Create virtual environment if it doesn't exist
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

:: Activate venv and install dependencies
echo Installing dependencies...
call venv\Scripts\activate.bat
pip install -r requirements.txt --quiet

:: Create folders
if not exist "instance" mkdir instance
if not exist "uploads" mkdir uploads

:: Run the application
echo.
echo ============================================
echo   Starting CPMS Evaluation Platform
echo   Open http://127.0.0.1:5000 in your browser
echo ============================================
echo.
python app.py

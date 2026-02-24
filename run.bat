@echo off
chcp 65001 >nul 2>&1
title 5040 Agents - Delivery Performance Analysis

echo ══════════════════════════════════════════════════════════
echo   5040 Agents - Delivery Performance Report Generator
echo ══════════════════════════════════════════════════════════
echo.

REM ─── تنظیم مسیر پروژه ───
cd /d "%~dp0"
cd 5040Agents-add_bat

REM ─── بررسی وجود پایتون ───
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found! Please install Python and add it to PATH.
    echo Download: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [INFO] Python found:
python --version
echo.

REM --- Setup venv ---
if not exist "venv" (
    echo [..] Creating virtual environment...
    python -m venv venv
    echo [OK] Virtual environment created.
)

call venv\Scripts\activate.bat

REM ─── نصب وابستگی‌ها (در صورت نیاز) ───
echo [INFO] Checking and installing dependencies...
pip install openpyxl pandas numpy jdatetime --quiet
if %errorlevel% neq 0 (
    echo [WARNING] Some dependencies may have failed to install.
    echo [INFO] Trying to continue anyway...
)
echo [OK] Dependencies are ready.
echo.

REM ─── اجرای برنامه ───
echo [INFO] Running the analysis...
echo ──────────────────────────────────────────────────────────
echo.

if "%~1"=="" (
    python main.py
) else (
    python main.py "%~1"
)

echo.
if %errorlevel% equ 0 (
    echo ══════════════════════════════════════════════════════════
    echo   [DONE] Process completed successfully!
    echo ══════════════════════════════════════════════════════════
) else (
    echo ══════════════════════════════════════════════════════════
    echo   [FAILED] An error occurred during execution.
    echo ══════════════════════════════════════════════════════════
)

echo.
pause

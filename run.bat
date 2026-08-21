@echo off
title InvenScan - Smart Inventory System
color 0B
cd /d "%~dp0"

echo ================================================================
echo           STARTING INVENSCAN SMART INVENTORY SYSTEM
echo ================================================================
echo.

:: Check for Python
where python >nul 2>nul
if %errorlevel% neq 0 (
    if exist ".venv\Scripts\python.exe" (
        set "PY_CMD=.venv\Scripts\python.exe"
    ) else (
        color 0C
        echo [ERROR] Python is not installed or not in PATH!
        echo Please install Python 3.10+ from https://www.python.org/
        echo.
        pause
        exit /b 1
    )
) else (
    set "PY_CMD=python"
)

echo [OK] Python found. Launching server...
echo.

%PY_CMD% app.py

if %errorlevel% neq 0 (
    echo.
    echo Server stopped.
    pause
)

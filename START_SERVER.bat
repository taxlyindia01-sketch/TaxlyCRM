@echo off
title Taxly CRM - Backend Server
color 0A
echo.
echo  ████████╗ █████╗ ██╗  ██╗██╗  ██╗   ██╗
echo     ██╔══╝██╔══██╗╚██╗██╔╝██║  ╚██╗ ██╔╝
echo     ██║   ███████║ ╚███╔╝ ██║   ╚████╔╝
echo     ██║   ██╔══██║ ██╔██╗ ██║    ╚██╔╝
echo     ██║   ██║  ██║██╔╝ ██╗███████╗██║
echo     ╚═╝   ╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝╚═╝
echo.
echo  Taxly CRM Backend Startup
echo  ─────────────────────────────────────────
echo.

cd /d "%~dp0backend"

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.10+ from python.org
    pause
    exit /b 1
)

echo [1/3] Installing / updating Python packages...
pip install -r requirements.txt -q --disable-pip-version-check
if errorlevel 1 (
    echo [WARN] Some packages may have failed. Trying to continue...
)

echo [2/3] Packages ready.
echo [3/3] Starting Taxly backend on http://127.0.0.1:8000 ...
echo.
echo  ✅ Once started, open VS Code Live Server on frontend/index.html
echo  ✅ OR visit: http://localhost:8000  (backend serves frontend too)
echo.
echo  Press Ctrl+C to stop the server
echo  ─────────────────────────────────────────
echo.

uvicorn server:app --host 127.0.0.1 --port 8000 --reload

if errorlevel 1 (
    echo.
    echo [ERROR] Server failed to start. Common fixes:
    echo   1. Port 8000 already in use - close other apps using port 8000
    echo   2. Database error - check backend/.env DATABASE_URL
    echo   3. Missing package - run: pip install -r requirements.txt
    echo.
    pause
)

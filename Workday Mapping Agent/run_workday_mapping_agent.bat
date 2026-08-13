@echo off
setlocal

echo ============================================================
echo   PeopleSoft ^-^> Workday Migration Agent
echo   Accenture HR Data Migration
echo ============================================================
echo.

REM ── Paths ────────────────────────────────────────────────────
set "APP_DIR=C:\Users\tejas.raj.srivastava\Desktop\ps_to_workday_migration\ps_to_workday"
set "HTML_FILE=C:\Users\tejas.raj.srivastava\Desktop\Automation\Agentic AI Demo\Agentic AI Demo\PS_to_Workday_Migration_Agent.html"
set "PORT=8502"

REM ── Install / upgrade dependencies ───────────────────────────
echo [1/3] Checking dependencies...
pip install streamlit>=1.35.0 pandas>=2.0.0 openpyxl>=3.1.0 -q --disable-pip-version-check
if %errorlevel% neq 0 (
    echo ERROR: pip install failed. Check your Python environment.
    pause
    exit /b 1
)
echo       Dependencies OK.
echo.

REM ── Option A — Open HTML Agent (no server needed) ────────────
echo [2/3] Browser-based HTML Agent (no server required):
echo       %HTML_FILE%
echo.
set /p OPEN_HTML=Open HTML Agent in Chrome now? (Y/N):
if /i "%OPEN_HTML%"=="Y" (
    start "" "%HTML_FILE%"
    echo       Opened in default browser.
    echo.
)

REM ── Option B — Launch Streamlit App ──────────────────────────
echo [3/3] Starting Streamlit web app on port %PORT%...
echo       URL: http://localhost:%PORT%
echo.
echo       Press Ctrl+C to stop the server.
echo ============================================================
echo.

cd /d "%APP_DIR%"
streamlit run app.py --server.port %PORT% --browser.gatherUsageStats false --server.headless false

pause
endlocal

@echo off
setlocal
cd /d "%~dp0"
set "INVENTORY_APP_TEST=1"

REM Start Streamlit locally (offline, localhost only) on port 3000.
start "" python -m streamlit run "app.py" --server.port 3000 --global.developmentMode false --server.headless false

REM Give server a moment to boot, then open browser.
timeout /t 3 /nobreak >nul
start "" "http://localhost:3000"

endlocal

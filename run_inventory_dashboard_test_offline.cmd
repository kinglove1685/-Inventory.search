@echo off
setlocal

cd /d "%~dp0"

set "PY_CMD="
set "INVENTORY_APP_TEST=1"
set "INVENTORY_APP_PAGE_TITLE=Inventory Search TEST"
set "INVENTORY_APP_DASHBOARD_TITLE=Inventory Search Dashboard TEST"
set "STREAMLIT_SERVER_HEADLESS=false"
set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"

where python >nul 2>&1
if not errorlevel 1 set "PY_CMD=python"
if not defined PY_CMD (
    where py >nul 2>&1
    if not errorlevel 1 set "PY_CMD=py -3"
)

if not defined PY_CMD (
    echo [ERROR] Python runtime not found.
    goto :end
)

if not exist "app.py" (
    echo [ERROR] app.py not found in current folder.
    goto :end
)

echo Starting Streamlit dashboard (TEST mode)...
%PY_CMD% -m streamlit run app.py --global.developmentMode false --server.port 3000 --server.headless false

:end
echo.
pause
endlocal

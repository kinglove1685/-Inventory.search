@echo off
setlocal

cd /d "%~dp0"

set "PY_CMD="

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

if not exist "build_inventory_dummy.py" (
    echo [ERROR] Missing script: build_inventory_dummy.py
    goto :end
)

echo ============================================
echo Inventory dummy one-click builder
echo ============================================

%PY_CMD% "build_inventory_dummy.py"
if errorlevel 1 (
    echo [FAILED] Could not create dummy workbook.
    goto :end
)

echo.
echo [DONE] inventory_dummy.xlsx created.

:end
echo.
pause
endlocal

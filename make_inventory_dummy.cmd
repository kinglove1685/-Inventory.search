@echo off
setlocal

cd /d "%~dp0"

set "SCRIPT=extract_inventory_sheet.py"
set "OUT=inventory_dummy.xlsx"
set "SRC=%~1"
set "PY_CMD="

if not defined SRC (
    for %%F in (*.xlsx) do (
        set "NAME=%%~nxF"
        if /I not "%%~nxF"=="%OUT%" if /I not "%%~nF"=="~$%%~nF" (
            set "SRC=%%~fF"
            goto :found
        )
    )
)

:found
echo ============================================
echo Inventory dummy builder
echo Source: "%SRC%"
echo Output: "%CD%\%OUT%"
echo ============================================

if not exist "%SCRIPT%" (
    echo [ERROR] Script not found: "%SCRIPT%"
    goto :end
)

if not defined SRC (
    echo [ERROR] No source Excel found in this folder.
    echo.
    echo Usage:
    echo 1. Double-click this file in the project folder.
    echo 2. Or drag and drop a source .xlsx file onto this .cmd file.
    goto :end
)

if not exist "%SRC%" (
    echo [ERROR] Source Excel not found: "%SRC%"
    goto :end
)

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

%PY_CMD% "%SCRIPT%" "%SRC%" -o "%OUT%"
if errorlevel 1 (
    echo [FAILED] Could not create dummy file.
    goto :end
)

echo.
echo [DONE] Dummy file created:
echo "%CD%\%OUT%"

:end
echo.
pause
endlocal

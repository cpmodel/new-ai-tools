@echo off
REM ===========================================================================
REM  2_build_tool.bat  -  THE TOOL
REM
REM  Builds the interactive workbook from the panel produced by step 1.
REM    output\global_data_long.csv  +  config\  ->  output\Global_fiscal_indicators.xlsx
REM
REM  Run 1_build_database.bat first. This step does not read any raw source file,
REM  so you can re-run it on its own after editing config\indicators.csv - for
REM  example to move an indicator between blocks or change a label.
REM ===========================================================================
setlocal
pushd "%~dp0"

echo.
echo  === Locating Python ===
set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY where python >nul 2>&1 && set "PY=python"
if not defined PY (
    echo  ERROR: no Python found on PATH. Run 0_setup.bat first.
    goto :fail
)

echo  === Checking inputs ===
if not exist "%~dp0output\global_data_long.csv" (
    echo  ERROR: output\global_data_long.csv not found.
    echo  Run 1_build_database.bat first.
    goto :fail
)
if not exist "%~dp0config\indicators.csv" (
    echo  ERROR: config\indicators.csv not found.
    goto :fail
)

echo.
echo  === Building the workbook ===
%PY% "%~dp0scripts\build_dashboard.py" ^
     --out "%~dp0output" ^
     --config "%~dp0config"
if errorlevel 1 goto :fail

echo.
echo  === Done ===
echo  output\Global_fiscal_indicators.xlsx
echo.
echo  Open it, go to the Dashboard sheet, and pick a country in the yellow cell.
echo  NOTE: Excel may need to recalculate on first open. If the Dashboard looks
echo  empty, press Ctrl+Alt+F9 to force a full recalculation.
popd
endlocal
pause
exit /b 0

:fail
echo.
echo  *** BUILD FAILED - see the messages above ***
popd
endlocal
pause
exit /b 1

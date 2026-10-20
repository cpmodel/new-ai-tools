@echo off
REM ===========================================================================
REM  1_build_database.bat  -  PREPROCESSING
REM
REM  Reads every raw source file and builds the indicator panel.
REM    raw_data\  +  config\indicators.csv   ->   output\global_data_*.csv
REM
REM  Optional argument:  ccdr    restrict to the 98 published-CCDR economies
REM      1_build_database.bat ccdr
REM
REM  Takes a couple of minutes: the CPAT .xlsb alone is ~59,000 rows.
REM ===========================================================================
setlocal
pushd "%~dp0"

set "EXTRA="
if /I "%~1"=="ccdr" set "EXTRA=--ccdr-only"

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
if not exist "%~dp0raw_data" (
    echo  ERROR: raw_data folder not found next to this file.
    goto :fail
)
if not exist "%~dp0config\indicators.csv" (
    echo  ERROR: config\indicators.csv not found.
    goto :fail
)
if not exist "%~dp0output" mkdir "%~dp0output"

echo.
echo  === Building the panel %EXTRA% ===
%PY% "%~dp0scripts\build_global_indicators.py" ^
     --raw "%~dp0raw_data" ^
     --out "%~dp0output" ^
     --config "%~dp0config" %EXTRA%
if errorlevel 1 goto :fail

echo.
echo  === Also writing the CPAT sheets for the R route ===
%PY% "%~dp0scripts\build_global_indicators.py" ^
     --raw "%~dp0raw_data" ^
     --out "%~dp0output" ^
     --config "%~dp0config" --emit-cpat-xlsx >nul
REM  not fatal if this fails - only the R pipeline needs it

echo.
echo  === Done. Files written to output\ ===
dir /b "%~dp0output\*.csv"
echo.
echo  Next: run 2_build_tool.bat
popd
endlocal
pause
exit /b 0

:fail
echo.
echo  *** BUILD FAILED - see the messages above ***
echo  A "PIPELINE/CONFIG MISMATCH" message means config\indicators.csv lists an
echo  indicator the pipeline does not produce. Fix the config or add the extraction.
popd
endlocal
pause
exit /b 1

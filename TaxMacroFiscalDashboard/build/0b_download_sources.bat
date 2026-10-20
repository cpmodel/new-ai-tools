@echo off
REM ===========================================================================
REM  0b_download_sources.bat  -  fetch the sources that have a direct URL
REM
REM  Reads config\sources.csv. Files land in raw_data_downloaded\, NOT raw_data\,
REM  so nothing you already have is overwritten.
REM
REM  Six of the nineteen sources can be fetched automatically. The rest need an
REM  interactive query or sit behind a JavaScript page - the script prints those
REM  as a checklist with the settings to choose.
REM
REM    0b_download_sources.bat            fetch
REM    0b_download_sources.bat list       show the plan without downloading
REM    0b_download_sources.bat force      re-fetch files already present
REM ===========================================================================
setlocal
pushd "%~dp0"

set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY where python >nul 2>&1 && set "PY=python"
if not defined PY (
    echo  ERROR: no Python found on PATH. Run 0_setup.bat first.
    goto :fail
)

set "EXTRA="
if /I "%~1"=="list"  set "EXTRA=--list"
if /I "%~1"=="force" set "EXTRA=--force"

%PY% "%~dp0scripts\download_sources.py" ^
     --config "%~dp0config" ^
     --dest "%~dp0raw_data_downloaded" %EXTRA%

echo.
echo  Files are in raw_data_downloaded\ - compare them against raw_data\
echo  before swapping any in.
popd
endlocal
pause
exit /b 0

:fail
popd
endlocal
pause
exit /b 1

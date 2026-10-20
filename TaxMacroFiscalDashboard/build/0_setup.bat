@echo off
REM ===========================================================================
REM  0_setup.bat  -  install the Python packages the build needs
REM  Run this once. Then run 1_build_database.bat, then 2_build_tool.bat.
REM ===========================================================================
setlocal
pushd "%~dp0"

echo.
echo  === Checking for Python ===
set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY where python >nul 2>&1 && set "PY=python"
if not defined PY (
    echo  ERROR: no Python found on PATH.
    echo  Install Python 3.9 or later from python.org, ticking "Add to PATH".
    goto :fail
)
echo  Using: %PY%
%PY% --version

echo.
echo  === Installing packages from requirements.txt ===
%PY% -m pip install --upgrade pip
%PY% -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 goto :fail

echo.
echo  === Done. Next: run 1_build_database.bat ===
popd
endlocal
pause
exit /b 0

:fail
echo.
echo  *** SETUP FAILED - see the messages above ***
popd
endlocal
pause
exit /b 1

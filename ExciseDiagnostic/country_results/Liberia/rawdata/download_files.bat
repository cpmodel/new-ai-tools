@echo off
rem Downloads every link in download_list.txt into a "rawdata" subfolder.
rem Keep this file and download_list.txt in the same folder.
rem Uses curl, which is built into Windows 10 and 11.

set "LIST=%~dp0download_list.txt"
set "OUT=%~dp0rawdata"

if not exist "%LIST%" (
    echo Cannot find download_list.txt next to this file.
    pause
    exit /b 1
)

if not exist "%OUT%" mkdir "%OUT%"
cd /d "%OUT%"

for /f "usebackq delims=" %%u in ("%LIST%") do (
    echo Downloading %%~nxu
    curl -L -O "%%u"
)

echo.
echo Done. Files are in %OUT%
pause

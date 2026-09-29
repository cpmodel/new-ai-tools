@echo off
setlocal

rem === Settings: must match start-day.bat ===
set "REPO_DIR=C:\Users\wb547395\Repos\new-ai-tools"

echo === End of day: new-ai-tools ===
echo.

cd /d "%REPO_DIR%" || goto :fail

echo --- Changes today ---
git status -s
echo.

rem Stage everything (new, changed, deleted files)
git add -A || goto :fail

rem Skip the commit if nothing is staged
git diff --cached --quiet
if %errorlevel%==0 (
    echo Nothing new to commit.
    goto :push
)

set "MSG="
set /p "MSG=Commit message (press Enter for default): "
if "%MSG%"=="" set "MSG=End of day: %date% %time:~0,5%"

git commit -m "%MSG%" || goto :fail

:push
echo.
echo Syncing with GitHub before pushing...
git pull --rebase --autostash || goto :fail

echo.
echo Pushing...
git push || goto :fail

echo.
echo --- Final status ---
git status -sb
echo.
echo Done. Work is safe on GitHub.
pause
exit /b 0

:fail
echo.
echo *** Something went wrong. Read the message above - nothing further was done. ***
pause
exit /b 1

@echo off
setlocal

rem === Settings: change REPO_DIR to where you keep the repo ===
set "REPO_URL=https://github.com/sjstretton/new-ai-tools.git"
set "REPO_DIR=C:\Users\wb547395\Repos\new-ai-tools"

echo === Start of day: new-ai-tools ===
echo.

rem Clone the repo if it isn't there yet
if not exist "%REPO_DIR%\.git" (
    echo Repo not found at %REPO_DIR% - cloning...
    git clone "%REPO_URL%" "%REPO_DIR%" || goto :fail
)

cd /d "%REPO_DIR%" || goto :fail

echo Fetching latest from GitHub...
git fetch --all --prune || goto :fail

echo.
echo Pulling (local uncommitted changes are stashed and restored)...
git pull --rebase --autostash || goto :fail

echo.
echo --- Current status ---
git status -sb

echo.
echo --- Last 5 commits ---
git log --oneline -5

echo.
echo All set. Have a good day.
pause
exit /b 0

:fail
echo.
echo *** Something went wrong. Read the message above before carrying on. ***
pause
exit /b 1

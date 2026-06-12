@echo off
echo ============================================================
echo  SmartPayer PDF Matcher - Push to GitHub
echo ============================================================
echo.

:: Change to the project directory
cd /d "%~dp0"

:: Check git is available
git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] git is not installed or not in PATH.
    echo         Download it from https://git-scm.com/downloads
    pause
    exit /b 1
)

:: Prompt for GitHub repo URL if not already set
set /p REPO_URL="Enter your GitHub repository URL (e.g. https://github.com/yourname/smartpayer-pdf-matcher): "

if "%REPO_URL%"=="" (
    echo [ERROR] No URL entered. Exiting.
    pause
    exit /b 1
)

:: Initialise git repo if not already done
if not exist ".git\" (
    echo.
    echo [INFO] Initialising git repository...
    git init
    git branch -M main
)

:: Create .gitignore to exclude PDFs, temp files, and logs
if not exist ".gitignore" (
    echo [INFO] Creating .gitignore...
    (
        echo # Temporary Word lock files
        echo ~$*
        echo.
        echo # Conversion logs
        echo Generated_Letters/conversion_log_*.txt
        echo.
        echo # Generated PDFs ^(optional: remove these lines to include PDFs^)
        echo # Generated_Letters/*.pdf
        echo.
        echo # Python cache
        echo __pycache__/
        echo *.pyc
    ) > .gitignore
)

:: Stage all files
echo.
echo [INFO] Staging files...
git add .
git status

:: Commit
echo.
echo [INFO] Creating commit...
git commit -m "Initial commit: SmartPayer PDF Converter"

:: Set remote (remove existing 'origin' if present)
git remote remove origin >nul 2>&1
git remote add origin %REPO_URL%

:: Push
echo.
echo [INFO] Pushing to %REPO_URL% ...
git push -u origin main

if errorlevel 1 (
    echo.
    echo [ERROR] Push failed.
    echo         - Make sure the repository exists on GitHub ^(create it first at https://github.com/new^)
    echo         - Make sure you are authenticated ^(run: git config --global credential.helper manager^)
    echo         - If the repo has existing content, pull first: git pull origin main --allow-unrelated-histories
) else (
    echo.
    echo [SUCCESS] Project pushed to GitHub successfully!
    echo          %REPO_URL%
)

echo.
pause

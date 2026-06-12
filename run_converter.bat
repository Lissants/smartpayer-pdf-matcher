@echo off
echo ============================================================
echo  SmartPayer PDF Converter - Run
echo ============================================================
echo.

:: Change to the script's directory so relative paths work correctly
cd /d "%~dp0"

:: Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not found in PATH.
    echo         Run install_prerequisites.bat first.
    pause
    exit /b 1
)

:: Check pywin32 is installed
python -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] pywin32 is not installed.
    echo         Run install_prerequisites.bat first.
    pause
    exit /b 1
)

:: Check the script exists
if not exist "convert_missing_pdfs.py" (
    echo [ERROR] convert_missing_pdfs.py not found in:
    echo         %~dp0
    pause
    exit /b 1
)

:: Check Generated_Letters folder exists
if not exist "Generated_Letters\" (
    echo [ERROR] Generated_Letters folder not found!
    echo         Expected location: %~dp0Generated_Letters\
    echo         Make sure this script is in the same folder as Generated_Letters\
    echo         and that the folder name is spelled exactly 'Generated_Letters'.
    pause
    exit /b 1
)

echo [INFO] Starting conversion...
echo.
python convert_missing_pdfs.py

echo.
if errorlevel 1 (
    echo [ERROR] Script exited with an error. Check the log in Generated_Letters\.
) else (
    echo [INFO] Done. Check the log file in Generated_Letters\ for details.
)

pause

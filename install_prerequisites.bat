@echo off
echo ============================================================
echo  SmartPayer PDF Converter - Install Prerequisites
echo ============================================================
echo.

:: Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo         Please install Python from https://www.python.org/downloads/
    echo         and ensure "Add Python to PATH" is checked during install.
    pause
    exit /b 1
)

echo [INFO] Python found:
python --version
echo.

:: Upgrade pip silently
echo [INFO] Upgrading pip...
python -m pip install --upgrade pip --quiet

:: Install pywin32
echo [INFO] Installing pywin32...
python -m pip install pywin32

if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install pywin32. Check your internet connection
    echo         or try running this file as Administrator.
    pause
    exit /b 1
)

:: Run post-install script (registers COM components)
echo.
echo [INFO] Running pywin32 post-install setup...
python -c "import subprocess, sys, pathlib; scripts = pathlib.Path(sys.exec_prefix) / 'Scripts'; subprocess.run([sys.executable, str(scripts / 'pywin32_postinstall.py'), '-install'], check=True)" 2>nul
if errorlevel 1 (
    echo [WARN] Post-install step had a warning (may be OK if already registered).
)

echo.
echo ============================================================
echo  All prerequisites installed successfully.
echo ============================================================
pause

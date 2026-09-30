@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo   Acme PACP/LACP Checker - Build Script
echo ==========================================
echo.

where py >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found.
    echo Install Python 3.12, 3.13, or 3.14 from python.org.
    echo Make sure the Python launcher is enabled during installation.
    echo.
    pause
    exit /b 1
)

set "PY_CMD="
py -3.14 -c "import sys" >nul 2>nul && set "PY_CMD=py -3.14"
if not defined PY_CMD py -3.13 -c "import sys" >nul 2>nul && set "PY_CMD=py -3.13"
if not defined PY_CMD py -3.12 -c "import sys" >nul 2>nul && set "PY_CMD=py -3.12"
if not defined PY_CMD set "PY_CMD=py"

echo Using Python:
%PY_CMD% --version
if errorlevel 1 goto :fail

echo.
echo Recreating clean local virtual environment...
if exist ".venv" rmdir /s /q ".venv"
%PY_CMD% -m venv .venv
if errorlevel 1 goto :fail

echo Installing/updating required packages...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :fail
".venv\Scripts\python.exe" -m pip install --no-cache-dir -r requirements.txt
if errorlevel 1 goto :fail

echo.
echo Cleaning previous build output...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo Building application...
".venv\Scripts\python.exe" -m PyInstaller --clean --noconfirm pacp_lacp_checker.spec
if errorlevel 1 goto :fail

echo.
echo ==========================================
echo BUILD COMPLETE
echo Your new EXE is here:
echo %CD%\dist\pacp_lacp_checker.exe
echo ==========================================
echo.
pause
exit /b 0

:fail
echo.
echo BUILD FAILED. Review the error messages above.
echo.
pause
exit /b 1

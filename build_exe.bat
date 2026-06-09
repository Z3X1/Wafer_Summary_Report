@echo off
setlocal
title Build Wafer_Summary_Report.exe
echo ================================================
echo   Wafer_Summary_Report - Auto Build Tool
echo ================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "PY_FILE=%SCRIPT_DIR%Wafer_Summary_Report.py"

if not exist "%PY_FILE%" (
    echo [ERROR] Wafer_Summary_Report.py not found.
    echo         Place build_exe.bat and Wafer_Summary_Report.py in the same folder.
    pause
    exit /b 1
)

echo [1/4] Checking Python...
set "PYTHON="

python --version >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON=python"
    goto :FOUND_PYTHON
)

py --version >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON=py"
    goto :FOUND_PYTHON
)

echo       Python not found. Downloading Python 3.11.9...
set "PY_INST=%TEMP%\python_install.exe"
set "PY_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"

powershell -NoProfile -ExecutionPolicy Bypass -Command "='SilentlyContinue'; Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%PY_INST%'"

if not exist "%PY_INST%" (
    echo [ERROR] Download failed. Check your internet connection.
    pause
    exit /b 1
)

echo       Installing Python silently...
"%PY_INST%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
del "%PY_INST%" >nul 2>&1

set "PATH=%LOCALAPPDATA%\Programs\Python\Python311;%LOCALAPPDATA%\Programs\Python\Python311\Scripts;%PATH%"

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python installed but not found. Reopen this window and try again.
    pause
    exit /b 1
)
set "PYTHON=python"

:FOUND_PYTHON
%PYTHON% --version
echo       Python OK.

echo [2/4] Installing packages...
%PYTHON% -m pip install --upgrade pip -q --no-warn-script-location
%PYTHON% -m pip install openpyxl tkinterdnd2 pyinstaller -q --no-warn-script-location
if %errorlevel% neq 0 (
    echo [ERROR] Package install failed. Check your internet connection.
    pause
    exit /b 1
)
echo       Packages OK.

echo [3/4] Building exe (may take 1-3 minutes)...
cd /d "%SCRIPT_DIR%"

%PYTHON% -m PyInstaller --onefile --windowed --name "Wafer_Summary_Report" --collect-all tkinterdnd2 --hidden-import openpyxl --hidden-import openpyxl.styles --hidden-import openpyxl.utils --hidden-import xml.etree.ElementTree --clean --noconfirm "Wafer_Summary_Report.py"

if %errorlevel% neq 0 (
    echo [ERROR] PyInstaller failed. Screenshot the error above.
    pause
    exit /b 1
)

echo [4/4] Copying exe and cleaning up...

if exist "%SCRIPT_DIR%dist\Wafer_Summary_Report.exe" (
    copy /Y "%SCRIPT_DIR%dist\Wafer_Summary_Report.exe" "%SCRIPT_DIR%Wafer_Summary_Report.exe" >nul
    echo       Done: %SCRIPT_DIR%Wafer_Summary_Report.exe
) else (
    echo [WARNING] exe not found in dist folder.
)

rmdir /s /q "%SCRIPT_DIR%dist"  >nul 2>&1
rmdir /s /q "%SCRIPT_DIR%build" >nul 2>&1
del /q "%SCRIPT_DIR%Wafer_Summary_Report.spec" >nul 2>&1

echo.
echo ================================================
echo   Done! Double-click Wafer_Summary_Report.exe
echo   No Python needed on target machine.
echo ================================================
echo.
pause

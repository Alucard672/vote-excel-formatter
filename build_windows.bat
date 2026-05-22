@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set LOG_FILE=%cd%\build_windows.log
if exist "%LOG_FILE%" del "%LOG_FILE%"

echo [1/5] Checking Python...
py -3 --version >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto no_python

echo [2/5] Creating virtual environment...
if not exist ".venv" py -3 -m venv .venv >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo [3/5] Installing dependencies...
call ".venv\Scripts\python.exe" -m pip install --upgrade pip >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
call ".venv\Scripts\python.exe" -m pip install -r requirements.txt >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo [4/5] Building exe...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
call ".venv\Scripts\python.exe" -m PyInstaller --noconfirm Excel整理.spec >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo [5/5] Done.
echo.
echo Build finished:
echo %cd%\dist\Excel整理.exe
echo.
pause
exit /b 0

:no_python
echo.
echo Python was not found. Please install Python 3.10 or newer from python.org.
echo Log file: %LOG_FILE%
pause
exit /b 1

:failed
echo.
echo Build failed. Please send this log file:
echo %LOG_FILE%
echo.
type "%LOG_FILE%"
pause
exit /b 1

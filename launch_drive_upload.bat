@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "EXE_NAME=Drive Upload Launcher.exe"
set "ALT_EXE=drive_file_processor_gui.exe"

if exist "%SCRIPT_DIR%%EXE_NAME%" (
    start "" "%SCRIPT_DIR%%EXE_NAME%"
    exit /b 0
)

if exist "%SCRIPT_DIR%%ALT_EXE%" (
    start "" "%SCRIPT_DIR%%ALT_EXE%"
    exit /b 0
)

where pyw >nul 2>nul
if %ERRORLEVEL%==0 (
    start "" pyw "%SCRIPT_DIR%drive_file_processor_gui.py"
    exit /b 0
)

where python >nul 2>nul
if %ERRORLEVEL%==0 (
    start "" python "%SCRIPT_DIR%drive_file_processor_gui.py"
    exit /b 0
)

echo Python was not found on this computer.
echo Install Python or use the packaged .exe version of the launcher.
pause

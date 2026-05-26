@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
if exist "%~dp0drive_file_processor_gui.py" goto found_project
if defined DRIVE_UPLOAD_APP_DIR (
    set "SCRIPT_DIR=%DRIVE_UPLOAD_APP_DIR%"
    if not "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR%\"
    if exist "%SCRIPT_DIR%drive_file_processor_gui.py" goto found_project
)
set "SCRIPT_DIR=%USERPROFILE%\Downloads\drive_upload_automation-main\drive_upload_automation-main\"
if exist "%SCRIPT_DIR%drive_file_processor_gui.py" goto found_project
set "SCRIPT_DIR=%USERPROFILE%\Downloads\drive_upload_automation-main\"
if exist "%SCRIPT_DIR%drive_file_processor_gui.py" goto found_project
set "SCRIPT_DIR=%USERPROFILE%\Desktop\drive_upload_automation-main\drive_upload_automation-main\"
if exist "%SCRIPT_DIR%drive_file_processor_gui.py" goto found_project
set "SCRIPT_DIR=%USERPROFILE%\Desktop\drive_upload_automation-main\"
if exist "%SCRIPT_DIR%drive_file_processor_gui.py" goto found_project

echo Could not find drive_file_processor_gui.py next to this launcher.
echo.
echo If this launcher is on your Desktop, set DRIVE_UPLOAD_APP_DIR to the project folder first:
echo setx DRIVE_UPLOAD_APP_DIR "C:\path\to\drive_upload_automation-main\drive_upload_automation-main"
echo.
pause
exit /b 1

:found_project
set "EXE_NAME=Drive Upload Launcher.exe"
set "ALT_EXE=drive_file_processor_gui.exe"
set "VENV_PYTHONW=%SCRIPT_DIR%.venv\Scripts\pythonw.exe"
set "VENV_PYTHON=%SCRIPT_DIR%.venv\Scripts\python.exe"

if exist "%SCRIPT_DIR%%EXE_NAME%" (
    start "" "%SCRIPT_DIR%%EXE_NAME%"
    exit /b 0
)

if exist "%SCRIPT_DIR%%ALT_EXE%" (
    start "" "%SCRIPT_DIR%%ALT_EXE%"
    exit /b 0
)

if exist "%VENV_PYTHONW%" (
    start "" "%VENV_PYTHONW%" "%SCRIPT_DIR%drive_file_processor_gui.py"
    exit /b 0
)

if exist "%VENV_PYTHON%" (
    start "" "%VENV_PYTHON%" "%SCRIPT_DIR%drive_file_processor_gui.py"
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
echo.
echo If dependencies are missing, run these commands in PowerShell:
echo python -m venv .venv
echo .\.venv\Scripts\python.exe -m pip install -r requirements.txt
pause

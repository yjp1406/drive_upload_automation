@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

python -m pip install --upgrade pip
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --onefile --console --name "drive_file_processor" drive_file_processor.py
python -m PyInstaller --noconfirm --onefile --windowed --name "Drive Upload Launcher" drive_file_processor_gui.py

echo.
echo Build complete. The executables are in the dist folder.
pause

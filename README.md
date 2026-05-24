# File Processing and Google Drive Upload Automation

This script recursively scans a local folder, mirrors the same folder hierarchy in Google Drive, uploads supported files into the matching Drive folders, makes each file publicly viewable, and writes an Excel index of successful uploads.

## Supported Files

- `.pdf`
- `.jpg`
- `.jpeg`
- `.png`
- `.mp4`

## Files

- `drive_file_processor.py` - main automation script
- `drive_file_processor_gui.py` - desktop launcher for non-technical users
- `launch_drive_upload.bat` - one-click Windows launcher
- `build_drive_upload_exe.bat` - Windows packaging helper for a standalone exe
- `requirements.txt` - Python dependencies

## Setup

1. Install dependencies:

```bash
pip install -r requirements.txt
```

2. Create an OAuth client in Google Cloud and download the file as `credentials.json`.
3. Place `credentials.json` next to the script.
4. Create a `.env` file in the project root and add `ROOT_FOLDER`, `DRIVE_FOLDER_ID`, and `OUTPUT_FILE`.

If you want the desktop launcher, make sure your Python install includes `tkinter` as well. On many systems it is already included with Python.
For the Windows launcher or exe, keep `credentials.json` in the same folder as the launcher or packaged executable.

## Usage

The script loads `.env` automatically, so you only need to run:

```bash
python3 drive_file_processor.py
```

To use the desktop UI instead:

```bash
python3 drive_file_processor_gui.py
```

In the UI, fill in the root folder, Drive folder ID, and output file, then press `Run Upload`.
Only those three fields are shown. The credentials, logs, state file, and chunk size use built-in defaults.

On Windows, you can also double-click `launch_drive_upload.bat`.

Example `.env`:

```env
ROOT_FOLDER="/path/to/Root"
DRIVE_FOLDER_ID="your_google_drive_folder_id"
OUTPUT_FILE="uploaded_files.xlsx"
UPLOAD_STATE_FILE="upload_state.json"
```

Optional environment variables:

- `ROOT_FOLDER`
- `DRIVE_FOLDER_ID`
- `OUTPUT_FILE`
- `GOOGLE_CREDENTIALS_FILE`
- `GOOGLE_TOKEN_FILE`
- `UPLOAD_LOG_FILE`
- `UPLOAD_ERROR_LOG_FILE`
- `UPLOAD_STATE_FILE`
- `DRIVE_CHUNK_SIZE`
- `drive_file_processor_gui.py` also saves the last used UI values in `.drive_file_processor_gui.json`

## Output

- Excel report with `File Name` and `Drive Link`
- Excel report also includes `Status`, `Progress`, `Error`, and `Updated At`
- Main run log file
- Error-only log file for failures
- Checkpoint file used to resume interrupted uploads
- OAuth token cache stored in `token.pickle`

## Notes

- The first run opens a local browser flow for OAuth login.
- Uploaded files are shared as "Anyone with the link can view".
- The script creates a top-level Drive folder named after `ROOT_FOLDER` inside `DRIVE_FOLDER_ID`, then mirrors all subfolders under it.
- Empty local folders are created on Drive too, so the hierarchy stays aligned.
- If you stop the script and run it again, completed files are skipped and interrupted uploads resume from the saved checkpoint.
- If the internet drops during an upload or permission update, the script waits and continues from the same file when the connection returns.
- Duplicate Drive files are skipped automatically.
- The launcher shows a progress bar and a concise summary of uploaded, duplicate, and failed files.

## Windows Packaging

To build a packaged Windows version, run `build_drive_upload_exe.bat` on a machine that already has Python installed.
It installs `pyinstaller` and creates both `drive_file_processor.exe` and `Drive Upload Launcher.exe` in the `dist` folder.
Keep `credentials.json` beside those executables when you distribute them.

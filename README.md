# File Processing and Google Drive Upload Automation

This script recursively scans a local folder, mirrors the same folder hierarchy in Google Drive, uploads supported files into the matching Drive folders, makes each file publicly viewable, and writes an Excel index of successful uploads.

## Supported Files

- `.pdf`
- `.jpg`
- `.jpeg`
- `.mp4`

## Files

- `drive_file_processor.py` - main automation script
- `requirements.txt` - Python dependencies

## Setup

1. Install dependencies:

```bash
pip install -r requirements.txt
```

2. Create an OAuth client in Google Cloud and download the file as `credentials.json`.
3. Place `credentials.json` next to the script.
4. Create a `.env` file in the project root and add `ROOT_FOLDER`, `DRIVE_FOLDER_ID`, and `OUTPUT_FILE`.

## Usage

The script loads `.env` automatically, so you only need to run:

```bash
python3 drive_file_processor.py
```

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
- Duplicate Drive files are skipped automatically.

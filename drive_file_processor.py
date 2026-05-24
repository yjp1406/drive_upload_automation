"""Process local files, upload them to Google Drive, and generate an Excel report.

Supported extensions:
    - .pdf
    - .jpg
    - .jpeg
    - .png
    - .mp4

The script:
    1. Recursively scans a root folder.
    2. Mirrors the local folder hierarchy in Google Drive.
    3. Uploads each supported file into the matching Drive folder.
    4. Grants public view permission.
    5. Writes an Excel index of uploaded files.

Authentication uses OAuth 2.0 with credentials.json and token.pickle.
"""

from __future__ import annotations

import argparse
import json
import mimetypes
import logging
import os
import pickle
import socket
import time
import pytz
import dotenv
from collections import OrderedDict
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterator, Sequence

dotenv.load_dotenv()

import pandas as pd
from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import HttpRequest, MediaFileUpload
from httplib2 import HttpLib2Error

SUPPORTED_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png", ".mp4"}
SCOPES = ["https://www.googleapis.com/auth/drive"]
DRIVE_FOLDER_MIME_TYPE = "application/vnd.google-apps.folder"
DEFAULT_CREDENTIALS_FILE = "credentials.json"
DEFAULT_TOKEN_FILE = "token.pickle"
DEFAULT_OUTPUT_FILE = "uploaded_files.xlsx"
DEFAULT_LOG_FILE = "upload.log"
DEFAULT_ERROR_LOG_FILE = "upload_errors.log"
DEFAULT_STATE_FILE = "upload_state.json"
DEFAULT_CHUNK_SIZE = 10 * 1024 * 1024
UI_EVENT_PREFIX = "UI_EVENT "

ERROR_LOGGER = logging.getLogger("upload_errors")

STATUS_PENDING = "pending"
STATUS_UPLOADING = "uploading"
STATUS_UPLOADED_PRIVATE = "uploaded_private"
STATUS_UPLOADED = "uploaded"
STATUS_DUPLICATE = "duplicate"
STATUS_FAILED = "failed"
STATUS_PERMISSION_FAILED = "permission_failed"
STATUS_PAUSED = "paused"
STATUS_WAITING = "waiting_for_connection"


@dataclass
class UploadRecord:
    source_path: str
    renamed_file_name: str
    status: str = STATUS_PENDING
    progress: str = "0%"
    drive_link: str = ""
    drive_file_id: str = ""
    error: str = ""
    updated_at: str = ""
    resumable_request_json: str = ""

    def to_excel_row(self) -> dict[str, str]:
        return {
            "Source Path": self.source_path,
            "File Name": self.renamed_file_name,
            "Status": self.status,
            "Progress": self.progress,
            "Drive Link": self.drive_link,
            "Error": self.error,
            "Updated At": self.updated_at,
        }

    @staticmethod
    def from_dict(data: dict) -> "UploadRecord":
        return UploadRecord(
            source_path=data["source_path"],
            renamed_file_name=data["renamed_file_name"],
            status=data.get("status", STATUS_PENDING),
            progress=data.get("progress", "0%"),
            drive_link=data.get("drive_link", ""),
            drive_file_id=data.get("drive_file_id", ""),
            error=data.get("error", ""),
            updated_at=data.get("updated_at", ""),
            resumable_request_json=data.get("resumable_request_json", ""),
        )


class UploadTracker:
    def __init__(self, root_folder: Path, output_file: Path, state_file: Path) -> None:
        self.root_folder = root_folder
        self.output_file = output_file
        self.state_file = state_file
        self.records: "OrderedDict[str, UploadRecord]" = OrderedDict()
        self._load_state()

    def _load_state(self) -> None:
        if not self.state_file.exists():
            return

        try:
            payload = json.loads(self.state_file.read_text(encoding="utf-8"))
            for item in payload.get("records", []):
                record = UploadRecord.from_dict(item)
                self.records[record.source_path] = record
        except Exception:
            logging.exception("Failed to load state file %s; starting fresh", self.state_file)

    def ensure_records(self, files: Sequence[Path]) -> None:
        for file_path in files:
            source_path = str(file_path.relative_to(self.root_folder))
            file_name = drive_file_name(file_path)
            if source_path not in self.records:
                self.records[source_path] = UploadRecord(
                    source_path=source_path,
                    renamed_file_name=file_name,
                )
            else:
                record = self.records[source_path]
                if record.renamed_file_name != file_name:
                    record.renamed_file_name = file_name
                    if record.status in {
                        STATUS_PENDING,
                        STATUS_UPLOADING,
                        STATUS_FAILED,
                        STATUS_PERMISSION_FAILED,
                        STATUS_PAUSED,
                    }:
                        record.status = STATUS_PENDING
                        record.progress = "0%"
                        record.error = ""
                        record.drive_file_id = ""
                        record.drive_link = ""
                        record.resumable_request_json = ""

    def get(self, file_path: Path) -> UploadRecord:
        source_path = str(file_path.relative_to(self.root_folder))
        return self.records[source_path]

    def set_record(self, record: UploadRecord) -> None:
        self.records[record.source_path] = record

    def touch(self, record: UploadRecord) -> None:
        record.updated_at = ist_now()

    def save_state(self) -> None:
        payload = {
            "root_folder": str(self.root_folder),
            "records": [asdict(record) for record in self.records.values()],
        }
        self.state_file.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = self.state_file.with_suffix(self.state_file.suffix + ".tmp")
        tmp_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        tmp_path.replace(self.state_file)

    def save_excel(self) -> None:
        self.output_file.parent.mkdir(parents=True, exist_ok=True)
        df = pd.DataFrame([record.to_excel_row() for record in self.records.values()])
        if df.empty:
            df = pd.DataFrame(
                columns=["Source Path", "File Name", "Status", "Progress", "Drive Link", "Error", "Updated At"]
            )
        tmp_path = self.output_file.with_name(f"{self.output_file.stem}.tmp{self.output_file.suffix}")
        df.to_excel(tmp_path, index=False)
        tmp_path.replace(self.output_file)

    def flush(self) -> None:
        self.save_state()
        self.save_excel()


def ist_now() -> str:
    ist = pytz.timezone("Asia/Kolkata")
    return datetime.now(ist).isoformat(timespec="seconds")


def emit_ui_event(event_type: str, **payload) -> None:
    if os.environ.get("UI_PROGRESS") != "1":
        return
    message = {"type": event_type, **payload}
    print(f"{UI_EVENT_PREFIX}{json.dumps(message, separators=(',', ':'))}", flush=True)


def is_transient_connection_error(exc: Exception) -> bool:
    if isinstance(exc, (HttpLib2Error, TimeoutError, socket.timeout, socket.gaierror, ConnectionError, OSError)):
        return True
    message = str(exc).lower()
    transient_markers = [
        "connection reset",
        "connection aborted",
        "broken pipe",
        "timed out",
        "temporary failure in name resolution",
        "network is unreachable",
        "connection refused",
        "dns",
        "name or service not known",
        "failed to establish a new connection",
        "remote end closed connection",
    ]
    return any(marker in message for marker in transient_markers)


def is_transient_http_error(exc: HttpError) -> bool:
    return getattr(exc.resp, "status", None) in {429, 500, 502, 503, 504}


def configure_logging(log_file: Path, error_log_file: Path) -> None:
    log_file.parent.mkdir(parents=True, exist_ok=True)
    error_log_file.parent.mkdir(parents=True, exist_ok=True)

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)

    root_logger.addHandler(stream_handler)
    root_logger.addHandler(file_handler)

    error_handler = logging.FileHandler(error_log_file, encoding="utf-8")
    error_handler.setLevel(logging.ERROR)
    error_handler.setFormatter(formatter)

    ERROR_LOGGER.handlers.clear()
    ERROR_LOGGER.setLevel(logging.ERROR)
    ERROR_LOGGER.propagate = False
    ERROR_LOGGER.addHandler(error_handler)

    logging.captureWarnings(True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Upload supported files from a folder tree to Google Drive."
    )
    parser.add_argument(
        "--credentials-file",
        default=os.environ.get("GOOGLE_CREDENTIALS_FILE", DEFAULT_CREDENTIALS_FILE),
        help="OAuth client secrets JSON file.",
    )
    parser.add_argument(
        "--token-file",
        default=os.environ.get("GOOGLE_TOKEN_FILE", DEFAULT_TOKEN_FILE),
        help="Pickle file used to cache OAuth tokens.",
    )
    parser.add_argument(
        "--log-file",
        default=os.environ.get("UPLOAD_LOG_FILE", DEFAULT_LOG_FILE),
        help="Main run log file.",
    )
    parser.add_argument(
        "--error-log-file",
        default=os.environ.get("UPLOAD_ERROR_LOG_FILE", DEFAULT_ERROR_LOG_FILE),
        help="Error-only log file.",
    )
    parser.add_argument(
        "--state-file",
        default=os.environ.get("UPLOAD_STATE_FILE", DEFAULT_STATE_FILE),
        help="Checkpoint file used to resume interrupted uploads.",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=int(os.environ.get("DRIVE_CHUNK_SIZE", DEFAULT_CHUNK_SIZE)),
        help="Upload chunk size in bytes for resumable uploads.",
    )
    args = parser.parse_args()

    args.root_folder = os.environ.get("ROOT_FOLDER")
    args.drive_folder_id = os.environ.get("DRIVE_FOLDER_ID")
    args.output_file = os.environ.get("OUTPUT_FILE", DEFAULT_OUTPUT_FILE)
    args.state_file = os.environ.get("UPLOAD_STATE_FILE", DEFAULT_STATE_FILE)

    if not args.root_folder:
        parser.error("ROOT_FOLDER environment variable is required")
    if not args.drive_folder_id:
        parser.error("DRIVE_FOLDER_ID environment variable is required")

    return args


def load_credentials(credentials_file: Path, token_file: Path) -> Credentials:
    creds: Credentials | None = None
    if token_file.exists():
        try:
            with token_file.open("rb") as fh:
                creds = pickle.load(fh)
        except Exception:
            logging.exception("Cached token file %s is unreadable; starting fresh", token_file)
            creds = None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except RefreshError as exc:
            logging.warning("Saved token was rejected by Google (%s); re-authenticating", exc)
            creds = None
            if token_file.exists():
                token_file.unlink()

    if creds is None or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(str(credentials_file), SCOPES)
        creds = flow.run_local_server(port=0)

    token_file.parent.mkdir(parents=True, exist_ok=True)
    with token_file.open("wb") as fh:
        pickle.dump(creds, fh)

    return creds


def is_supported_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS


def iter_supported_files(root_folder: Path) -> Iterator[Path]:
    for current_root, _, files in os.walk(root_folder):
        current_root_path = Path(current_root)
        for file_name in files:
            candidate = current_root_path / file_name
            if is_supported_file(candidate):
                yield candidate


def drive_file_name(file_path: Path) -> str:
    return file_path.name


def build_drive_service(creds: Credentials):
    return build("drive", "v3", credentials=creds)


def build_drive_link(file_id: str) -> str:
    return f"https://drive.google.com/file/d/{file_id}/view"


def escape_drive_query_value(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def find_existing_drive_folder(service, parent_folder_id: str, folder_name: str) -> str | None:
    query = (
        f"name = '{escape_drive_query_value(folder_name)}' "
        f"and mimeType = '{DRIVE_FOLDER_MIME_TYPE}' "
        f"and '{escape_drive_query_value(parent_folder_id)}' in parents "
        "and trashed = false"
    )
    response = service.files().list(
        q=query,
        spaces="drive",
        pageSize=10,
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    folders = response.get("files", [])
    if not folders:
        return None
    if len(folders) > 1:
        logging.warning(
            "Multiple Drive folders matched %s in parent %s; using the first one",
            folder_name,
            parent_folder_id,
        )
    return folders[0]["id"]


def create_drive_folder(service, parent_folder_id: str, folder_name: str) -> str:
    response = service.files().create(
        body={
            "name": folder_name,
            "mimeType": DRIVE_FOLDER_MIME_TYPE,
            "parents": [parent_folder_id],
        },
        fields="id, name",
        supportsAllDrives=True,
    ).execute()
    return response["id"]


def ensure_drive_folder(service, parent_folder_id: str, folder_name: str) -> str:
    existing_folder_id = find_existing_drive_folder(service, parent_folder_id, folder_name)
    if existing_folder_id:
        return existing_folder_id
    return create_drive_folder(service, parent_folder_id, folder_name)


def ensure_drive_folder_path(
    service,
    drive_root_folder_id: str,
    root_folder_name: str,
    relative_folder_path: Path,
    folder_cache: dict[tuple[str, ...], str],
) -> str:
    current_key: tuple[str, ...] = ()
    if current_key not in folder_cache:
        folder_cache[current_key] = ensure_drive_folder(service, drive_root_folder_id, root_folder_name)

    current_folder_id = folder_cache[current_key]
    if relative_folder_path == Path("."):
        return current_folder_id

    for part in relative_folder_path.parts:
        if part == ".":
            continue
        current_key = (*current_key, part)
        if current_key not in folder_cache:
            folder_cache[current_key] = ensure_drive_folder(service, current_folder_id, part)
        current_folder_id = folder_cache[current_key]

    return current_folder_id


def find_existing_drive_file(service, drive_folder_id: str, file_name: str) -> str | None:
    query = (
        f"name = '{escape_drive_query_value(file_name)}' "
        f"and '{escape_drive_query_value(drive_folder_id)}' in parents "
        "and trashed = false"
    )
    response = service.files().list(
        q=query,
        spaces="drive",
        pageSize=10,
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = response.get("files", [])
    if not files:
        return None
    if len(files) > 1:
        logging.warning(
            "Multiple Drive files matched %s in folder %s; using the first one",
            file_name,
            drive_folder_id,
        )
    return files[0]["id"]


def build_upload_request(
    service,
    file_path: Path,
    file_name: str,
    drive_folder_id: str,
    chunk_size: int,
):
    metadata = {
        "name": file_name,
        "parents": [drive_folder_id],
    }
    mimetype = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
    media = MediaFileUpload(
        str(file_path),
        mimetype=mimetype,
        resumable=True,
        chunksize=chunk_size,
    )
    return service.files().create(
        body=metadata,
        media_body=media,
        fields="id, name",
        supportsAllDrives=True,
    )


def upload_or_resume_file(
    service,
    tracker: UploadTracker,
    record: UploadRecord,
    file_path: Path,
    drive_folder_id: str,
    chunk_size: int,
) -> None:
    if record.status in {STATUS_UPLOADED, STATUS_DUPLICATE}:
        return

    if record.status in {STATUS_UPLOADED_PRIVATE, STATUS_PERMISSION_FAILED} and record.drive_file_id:
        logging.info("Retrying permission update for %s", file_path)
        make_file_public(service=service, file_id=record.drive_file_id, file_path=file_path)
        record.status = STATUS_UPLOADED
        record.error = ""
        record.progress = "100%"
        record.updated_at = ist_now()
        tracker.flush()
        return

    duplicate_backoff_seconds = 5
    while True:
        try:
            existing_file_id = find_existing_drive_file(service, drive_folder_id, record.renamed_file_name)
            break
        except HttpError as exc:
            if is_transient_http_error(exc):
                logging.warning(
                    "Drive lookup failed while checking duplicates for %s. Waiting %s seconds before retrying.",
                    file_path,
                    duplicate_backoff_seconds,
                )
                ERROR_LOGGER.error(
                    "Drive lookup failed while checking duplicates for %s",
                    file_path,
                    exc_info=True,
                )
                record.status = STATUS_WAITING
                record.error = f"Waiting for connection: {exc}"
                record.updated_at = ist_now()
                tracker.flush()
                time.sleep(duplicate_backoff_seconds)
                duplicate_backoff_seconds = min(duplicate_backoff_seconds * 2, 60)
                continue
            raise
        except Exception as exc:
            if is_transient_connection_error(exc):
                logging.warning(
                    "Connection lost while checking duplicates for %s. Waiting %s seconds before retrying.",
                    file_path,
                    duplicate_backoff_seconds,
                )
                ERROR_LOGGER.error(
                    "Connection lost while checking duplicates for %s",
                    file_path,
                    exc_info=True,
                )
                record.status = STATUS_WAITING
                record.error = f"Waiting for connection: {exc}"
                record.updated_at = ist_now()
                tracker.flush()
                time.sleep(duplicate_backoff_seconds)
                duplicate_backoff_seconds = min(duplicate_backoff_seconds * 2, 60)
                continue
            raise

    if existing_file_id:
        record.status = STATUS_DUPLICATE
        record.drive_file_id = existing_file_id
        record.drive_link = build_drive_link(existing_file_id)
        record.progress = "100%"
        record.error = ""
        record.updated_at = ist_now()
        tracker.flush()
        logging.info("Skipping duplicate file already in Drive: %s", file_path)
        return

    logging.info("Uploading %s as %s", file_path, record.renamed_file_name)
    fresh_request = build_upload_request(
        service=service,
        file_path=file_path,
        file_name=record.renamed_file_name,
        drive_folder_id=drive_folder_id,
        chunk_size=chunk_size,
    )

    if record.resumable_request_json:
        request = HttpRequest.from_json(
            record.resumable_request_json,
            fresh_request.http,
            fresh_request.postproc,
        )
    else:
        request = fresh_request

    record.status = STATUS_UPLOADING
    record.error = ""
    record.updated_at = ist_now()
    tracker.flush()

    response = None
    network_backoff_seconds = 5
    try:
        while response is None:
            try:
                status, response = request.next_chunk(num_retries=3)
            except KeyboardInterrupt:
                record.status = STATUS_PAUSED
                record.resumable_request_json = request.to_json()
                record.updated_at = ist_now()
                tracker.flush()
                raise
            except HttpError as exc:
                if exc.resp.status == 404 and record.resumable_request_json:
                    logging.warning(
                        "Resumable session expired for %s; restarting that file from the beginning",
                        file_path,
                    )
                    record.status = STATUS_PENDING
                    record.progress = "0%"
                    record.resumable_request_json = ""
                    record.updated_at = ist_now()
                    tracker.flush()
                    upload_or_resume_file(
                        service=service,
                        tracker=tracker,
                        record=record,
                        file_path=file_path,
                        drive_folder_id=drive_folder_id,
                        chunk_size=chunk_size,
                    )
                    return
                raise
            except Exception as exc:
                if is_transient_connection_error(exc):
                    logging.warning(
                        "Connection lost while uploading %s. Waiting %s seconds before retrying.",
                        file_path,
                        network_backoff_seconds,
                    )
                    ERROR_LOGGER.error(
                        "Connection lost while uploading %s; will resume when network returns",
                        file_path,
                        exc_info=True,
                    )
                    record.status = STATUS_WAITING
                    record.error = f"Waiting for connection: {exc}"
                    record.resumable_request_json = request.to_json()
                    record.updated_at = ist_now()
                    tracker.flush()
                    emit_ui_event(
                        "upload_waiting",
                        file=record.source_path,
                        status=record.status,
                        progress=record.progress,
                        message=record.error,
                    )
                    time.sleep(network_backoff_seconds)
                    network_backoff_seconds = min(network_backoff_seconds * 2, 60)
                    continue
                raise

            if status is not None:
                record.status = STATUS_UPLOADING
                record.error = ""
                record.progress = f"{int(status.progress() * 100)}%"
                record.resumable_request_json = request.to_json()
                record.updated_at = ist_now()
                tracker.flush()
                network_backoff_seconds = 5

        file_id = response["id"]
        record.drive_file_id = file_id
        record.drive_link = build_drive_link(file_id)
        record.status = STATUS_UPLOADED_PRIVATE
        record.progress = "100%"
        record.resumable_request_json = ""
        record.updated_at = ist_now()
        tracker.flush()

        make_file_public(service=service, file_id=file_id, file_path=file_path)
        record.status = STATUS_UPLOADED
        record.error = ""
        record.updated_at = ist_now()
        tracker.flush()
    except Exception as exc:
        if record.status not in {STATUS_PAUSED, STATUS_DUPLICATE, STATUS_UPLOADED}:
            if record.drive_file_id and record.status == STATUS_UPLOADED_PRIVATE:
                record.status = STATUS_PERMISSION_FAILED
            else:
                record.status = STATUS_FAILED
            record.error = str(exc)
            record.updated_at = ist_now()
            tracker.flush()
        raise


def make_file_public(
    service,
    file_id: str,
    file_path: Path,
) -> None:
    backoff_seconds = 5
    while True:
        try:
            service.permissions().create(
                fileId=file_id,
                body={"type": "anyone", "role": "reader"},
                fields="id",
                supportsAllDrives=True,
            ).execute()
            return
        except HttpError as exc:
            if is_transient_http_error(exc):
                logging.warning(
                    "Permission update failed for %s. Waiting %s seconds before retrying.",
                    file_path,
                    backoff_seconds,
                )
                ERROR_LOGGER.error(
                    "Permission update failed for %s; will retry after connection returns",
                    file_path,
                    exc_info=True,
                )
                time.sleep(backoff_seconds)
                backoff_seconds = min(backoff_seconds * 2, 60)
                continue
            raise
        except Exception as exc:  # pragma: no cover - defensive guard
            if is_transient_connection_error(exc):
                logging.warning(
                    "Connection lost while updating permissions for %s. Waiting %s seconds before retrying.",
                    file_path,
                    backoff_seconds,
                )
                ERROR_LOGGER.error(
                    "Connection lost while updating permissions for %s; will retry after connection returns",
                    file_path,
                    exc_info=True,
                )
                time.sleep(backoff_seconds)
                backoff_seconds = min(backoff_seconds * 2, 60)
                continue
            logging.exception("Unexpected permission failure for %s", file_path)
            ERROR_LOGGER.exception("Unexpected permission failure for %s", file_path)
            raise


def process_files(
    root_folder: Path,
    drive_folder_id: str,
    output_file: Path,
    token_file: Path,
    credentials_file: Path,
    log_file: Path,
    error_log_file: Path,
    state_file: Path,
    chunk_size: int,
) -> int:
    configure_logging(log_file, error_log_file)

    if not root_folder.exists() or not root_folder.is_dir():
        raise FileNotFoundError(f"Root folder does not exist or is not a directory: {root_folder}")
    if not credentials_file.exists():
        raise FileNotFoundError(f"Missing OAuth credentials file: {credentials_file}")

    logging.info("Authenticating with Google Drive")
    creds = load_credentials(credentials_file, token_file)
    service = build_drive_service(creds)

    files = list(iter_supported_files(root_folder))
    tracker = UploadTracker(root_folder=root_folder, output_file=output_file, state_file=state_file)
    tracker.ensure_records(files)
    tracker.flush()
    emit_ui_event("run_started", scanned=len(files))

    folder_cache: dict[tuple[str, ...], str] = {}
    for current_root, _, _ in os.walk(root_folder):
        current_root_path = Path(current_root)
        relative_folder_path = current_root_path.relative_to(root_folder)
        ensure_drive_folder_path(
            service=service,
            drive_root_folder_id=drive_folder_id,
            root_folder_name=root_folder.name,
            relative_folder_path=relative_folder_path,
            folder_cache=folder_cache,
        )

    try:
        for file_path in files:
            record = tracker.get(file_path)
            destination_folder_id = ensure_drive_folder_path(
                service=service,
                drive_root_folder_id=drive_folder_id,
                root_folder_name=root_folder.name,
                relative_folder_path=file_path.parent.relative_to(root_folder),
                folder_cache=folder_cache,
            )
            upload_or_resume_file(
                service=service,
                tracker=tracker,
                record=record,
                file_path=file_path,
                drive_folder_id=destination_folder_id,
                chunk_size=chunk_size,
            )
            emit_ui_event(
                "file_complete",
                source_path=record.source_path,
                status=record.status,
                drive_link=record.drive_link,
            )
    except KeyboardInterrupt:
        logging.warning("Interrupted by user; progress has been saved.")
        emit_ui_event("run_interrupted")
        tracker.flush()
        return 130
    except Exception as exc:
        emit_ui_event("run_failed", error=str(exc))
        raise

    tracker.flush()
    uploaded_count = sum(1 for record in tracker.records.values() if record.status == STATUS_UPLOADED)
    duplicate_count = sum(1 for record in tracker.records.values() if record.status == STATUS_DUPLICATE)
    failed_count = sum(
        1 for record in tracker.records.values() if record.status in {STATUS_FAILED, STATUS_PERMISSION_FAILED}
    )
    emit_ui_event(
        "run_completed",
        scanned=len(files),
        uploaded=uploaded_count,
        duplicates=duplicate_count,
        failed=failed_count,
    )
    logging.info(
        "Completed. Scanned %s supported files, uploaded %s successfully, report written to %s",
        len(files),
        uploaded_count,
        output_file,
    )

    return 0


def main() -> int:
    script_dir = Path(__file__).resolve().parent

    args = parse_args()
    root_folder = Path(args.root_folder).expanduser().resolve()
    output_file = Path(args.output_file).expanduser().resolve()
    token_file = Path(args.token_file).expanduser().resolve()
    credentials_file = Path(args.credentials_file).expanduser().resolve()
    log_file = Path(args.log_file).expanduser().resolve()
    error_log_file = Path(args.error_log_file).expanduser().resolve()
    state_file = Path(args.state_file).expanduser().resolve()

    return process_files(
        root_folder=root_folder,
        drive_folder_id=args.drive_folder_id,
        output_file=output_file,
        token_file=token_file,
        credentials_file=credentials_file,
        log_file=log_file,
        error_log_file=error_log_file,
        state_file=state_file,
        chunk_size=args.chunk_size,
    )


if __name__ == "__main__":
    raise SystemExit(main())

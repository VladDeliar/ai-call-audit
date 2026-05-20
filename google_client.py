"""Авторизація OAuth та робота з Google Drive (читання + вивантаження).

OAuth (на відміну від service account) дозволяє створювати файли у вашому
Google Drive — вони належать вам і використовують вашу квоту сховища.
"""
import io
import os

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

import config

SCOPES = ["https://www.googleapis.com/auth/drive"]

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _credentials():
    """Отримати OAuth-облікові дані (з кешу token.json або через вхід у браузері)."""
    creds = None
    if os.path.exists(config.OAUTH_TOKEN):
        creds = Credentials.from_authorized_user_file(config.OAUTH_TOKEN, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(config.OAUTH_CLIENT):
                raise FileNotFoundError(
                    f"Не знайдено OAuth client JSON: {config.OAUTH_CLIENT}. "
                    "Створіть OAuth Client ID (Desktop app) у Google Cloud Console."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                config.OAUTH_CLIENT, SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open(config.OAUTH_TOKEN, "w", encoding="utf-8") as fh:
            fh.write(creds.to_json())
    return creds


class GoogleClient:
    """Обгортка над Drive API (читання + вивантаження)."""

    def __init__(self):
        self.drive = build("drive", "v3", credentials=_credentials(),
                           cache_discovery=False)

    # ---------- Читання ----------

    def list_folder(self, folder_id):
        """Усі (не в кошику) файли в папці."""
        files, page_token = [], None
        query = f"'{folder_id}' in parents and trashed = false"
        while True:
            resp = (
                self.drive.files()
                .list(
                    q=query,
                    fields="nextPageToken, files(id, name, mimeType)",
                    pageSize=1000,
                    pageToken=page_token,
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True,
                )
                .execute()
            )
            files.extend(resp.get("files", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
        return files

    def list_audio(self, folder_id):
        """Список аудіофайлів у папці."""
        result = []
        for f in self.list_folder(folder_id):
            name = f["name"].lower()
            if f["mimeType"].startswith(config.AUDIO_MIME_PREFIX) or name.endswith(
                config.AUDIO_EXTENSIONS
            ):
                result.append(f)
        return result

    def download_file(self, file_id, local_path):
        """Завантажити бінарний файл Drive на диск."""
        request = self.drive.files().get_media(fileId=file_id)
        with io.FileIO(local_path, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        return local_path

    def export_spreadsheet(self, sheet_id, local_path):
        """Експортувати Google Sheet як .xlsx на диск."""
        request = self.drive.files().export_media(fileId=sheet_id, mimeType=XLSX_MIME)
        with io.FileIO(local_path, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        return local_path

    # ---------- Вивантаження ----------

    def _find_in_folder(self, folder_id, name):
        """ID файлу з такою назвою в папці (або None)."""
        safe = name.replace("'", "\\'")
        resp = (
            self.drive.files()
            .list(
                q=f"'{folder_id}' in parents and name = '{safe}' and trashed = false",
                fields="files(id)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute()
        )
        files = resp.get("files", [])
        return files[0]["id"] if files else None

    def upload(self, local_path, folder_id, name=None, mime_type=None):
        """Вивантажити файл у папку Drive. Якщо файл з такою назвою вже є —
        оновити його вміст (без дублікатів). Повертає id файлу."""
        name = name or os.path.basename(local_path)
        media = MediaFileUpload(local_path, mimetype=mime_type, resumable=False)
        existing = self._find_in_folder(folder_id, name)
        if existing:
            self.drive.files().update(
                fileId=existing, media_body=media, supportsAllDrives=True
            ).execute()
            return existing
        created = (
            self.drive.files()
            .create(
                body={"name": name, "parents": [folder_id]},
                media_body=media,
                fields="id",
                supportsAllDrives=True,
            )
            .execute()
        )
        return created["id"]

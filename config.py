"""Завантаження налаштувань з .env."""
import os
import sys

from dotenv import load_dotenv

# Вивід кирилиці в консоль Windows (типове кодування cp1252 не підтримує її).
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

load_dotenv()


def _required(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        sys.exit(f"[config] Не задано обов'язкову змінну {name} у .env")
    return value


SOURCE_FOLDER_ID = _required("SOURCE_FOLDER_ID")
SOURCE_SHEET_ID = _required("SOURCE_SHEET_ID")
REPORT_FOLDER_ID = _required("REPORT_FOLDER_ID")
GEMINI_API_KEY = _required("GEMINI_API_KEY")

# OAuth: JSON клієнта (Desktop app) та кеш токена.
OAUTH_CLIENT = os.getenv("OAUTH_CLIENT", "credentials/oauth_client.json").strip()
OAUTH_TOKEN = os.getenv("OAUTH_TOKEN", "credentials/token.json").strip()

# Модель Gemini для транскрипції аудіо та аналізу дзвінків.
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash").strip()

try:
    MAX_FILES = int(os.getenv("MAX_FILES", "0"))
except ValueError:
    MAX_FILES = 0

# Локальна тека для тимчасових файлів (аудіо + транскрипти).
WORK_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "work_files")
os.makedirs(WORK_DIR, exist_ok=True)

AUDIO_EXTENSIONS = (".mp3", ".wav", ".m4a", ".ogg", ".flac", ".aac")
AUDIO_MIME_PREFIX = "audio/"

"""Тестовий скрипт: перевірка доступу та вивід усіх файлів з Google Drive.

Запуск: python test_drive.py
Виводить:
  - усі файли, доступні service account;
  - вміст вихідної та робочої папок із .env (якщо задані).
"""
import config
from google_client import GoogleClient


def list_all_files(gc):
    """Усі файли (не в кошику), доступні service account."""
    files, page_token = [], None
    while True:
        resp = (
            gc.drive.files()
            .list(
                q="trashed = false",
                fields="nextPageToken, files(id, name, mimeType, parents)",
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


def _print_files(files):
    if not files:
        print("    (порожньо)")
        return
    for f in files:
        print(f"    - {f['name']}  [{f['mimeType']}]  id={f['id']}")


def main():
    gc = GoogleClient()

    print("=== Усі файли, доступні service account ===")
    all_files = list_all_files(gc)
    print(f"Знайдено {len(all_files)} файл(ів):")
    _print_files(all_files)

    print(f"\n=== Вміст вихідної папки ({config.SOURCE_FOLDER_ID}) ===")
    try:
        _print_files(gc.list_folder(config.SOURCE_FOLDER_ID))
    except Exception as e:  # noqa: BLE001
        print(f"    Помилка доступу: {e}")

    print(f"\n=== Вміст робочої папки ({config.WORK_FOLDER_ID}) ===")
    try:
        _print_files(gc.list_folder(config.WORK_FOLDER_ID))
    except Exception as e:  # noqa: BLE001
        print(f"    Помилка доступу: {e}")


if __name__ == "__main__":
    main()

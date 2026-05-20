"""Транскрипція аудіо через Gemini API."""
import time

import config

_client = None

PROMPT = (
    "Транскрибуй цей аудіозапис телефонної розмови менеджера автосервісу з "
    "клієнтом. Мова — українська. Поверни ЛИШЕ дослівний текст розмови, без "
    "коментарів, без позначок часу та без імен учасників."
)


def _get_client():
    global _client
    if _client is None:
        from google import genai

        _client = genai.Client(api_key=config.GEMINI_API_KEY)
    return _client


def transcribe(audio_path):
    """Транскрибувати аудіофайл через Gemini. Повертає текст."""
    client = _get_client()

    uploaded = client.files.upload(file=audio_path)
    # Дочекатися, поки Gemini опрацює завантажений файл.
    while uploaded.state.name == "PROCESSING":
        time.sleep(2)
        uploaded = client.files.get(name=uploaded.name)
    if uploaded.state.name == "FAILED":
        raise RuntimeError(f"Gemini не зміг обробити файл: {audio_path}")

    try:
        resp = client.models.generate_content(
            model=config.GEMINI_MODEL,
            contents=[uploaded, PROMPT],
        )
    finally:
        try:
            client.files.delete(name=uploaded.name)
        except Exception:  # noqa: BLE001 — прибирання не критичне
            pass

    return (resp.text or "").strip()

"""Аналіз і класифікація колонок через Gemini API."""
import json

import config

_client = None


def _get_client():
    global _client
    if _client is None:
        from google import genai

        _client = genai.Client(api_key=config.GEMINI_API_KEY)
    return _client


def _generate_json(system_prompt, payload):
    """Запит до Gemini з відповіддю у форматі JSON."""
    from google.genai import types

    resp = _get_client().models.generate_content(
        model=config.GEMINI_MODEL,
        contents=json.dumps(payload, ensure_ascii=False),
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            response_mime_type="application/json",
            temperature=0,
        ),
    )
    return json.loads(resp.text)


SYSTEM_PROMPT = """Ти — аудитор якості телефонних дзвінків автосервісу.
Отримуєш транскрипт дзвінка менеджера з клієнтом (українською) та опис колонок
оцінкової таблиці. Проаналізуй розмову та поверни СТРОГО JSON-об'єкт.

Формат відповіді:
{"values": {"<назва колонки>": <значення>}, "comment_is_negative": <true|false>}

Для КОЖНОЇ колонки зі списку `columns` заповни значення відповідно до її `kind`:
- "binary"  -> 1, якщо критерій виконано в розмові, інакше 0.
- "score"   -> ціле число 0–10: загальна оцінка роботи менеджера.
- "choice"  -> обери РІВНО ОДИН варіант із масиву `options` цієї колонки —
              найближчий за змістом до обговореного в дзвінку. Поверни текст
              варіанта точно як у списку.
- "date"    -> дата (і час) запису на сервіс. Якщо домовлено й про час —
              формат "РРРР-ММ-ДД ГГ:ХХ", якщо лише про день — "РРРР-ММ-ДД".
              Дату обчисли на основі дати дзвінка (поле `call_date`) і згадок
              у розмові («завтра», «у п'ятницю о 13:00», «15-го» тощо).
              Якщо запис не призначено — порожній рядок "".
- "text"    -> короткий текст українською. Для колонки коментаря опиши якість
              розмови; якщо менеджер відповідав некоректно, грубо, помилково
              або не вирішив питання клієнта — чітко назви проблему.

`comment_is_negative` = true, якщо дзвінок пройшов НЕ ОК, менеджер відпрацював
погано/некоректно і це потребує реагування; інакше false.

Використовуй рівно ті назви колонок, що передані в `columns`."""


CLASSIFY_PROMPT = """Ти аналізуєш структуру оцінкової таблиці аудиту дзвінків
автосервісу. Дано список колонок (index + назва) та назви груп над ними.
Визнач тип КОЖНОЇ колонки:
- "binary" — критерій оцінки «так/ні» (заповнюється 1 або 0): чи виконав
  менеджер певну дію (привітання/представлення, уточнення кузова/року/пробігу,
  пропозиція діагностики, дотримання інструкцій, прощання тощо).
- "score"  — підсумкова числова оцінка роботи менеджера.
- "choice" — вибір типу/виду робіт зі списку.
- "text"   — довільний текст або метадані (дата, телефон, філія, менеджер,
  тип звернення, результат, запчастини, коментар, опис недотриманих рекомендацій).

Поверни СТРОГО JSON: {"kinds": {"<index>": "<kind>"}} для всіх переданих індексів."""


def classify_columns(columns_info, group_headers):
    """Класифікує колонки таблиці. columns_info — список {"index","name"}.
    Повертає {index(int): kind}."""
    payload = {"columns": columns_info, "group_headers": group_headers}
    kinds = _generate_json(CLASSIFY_PROMPT, payload).get("kinds", {})
    result = {}
    for k, v in kinds.items():
        try:
            result[int(k)] = v if v in ("binary", "score", "choice", "text") else "text"
        except (ValueError, TypeError):
            continue
    return result


def analyze(transcript, columns, call_date=None):
    """Аналізує транскрипт.

    columns   — список {"name": str, "kind": "binary|score|choice|date|text",
                "options": list|None}.
    call_date — дата дзвінка (рядок РРРР-ММ-ДД) для обчислення дат запису.
    Повертає {"values": {name: value}, "comment_is_negative": bool}.
    """
    payload = {"columns": columns, "transcript": transcript,
               "call_date": call_date or ""}
    data = _generate_json(SYSTEM_PROMPT, payload)

    values = data.get("values", {}) or {}
    normalized = {}
    for col in columns:
        name, kind = col["name"], col["kind"]
        raw = values.get(name, "")
        if kind == "binary":
            normalized[name] = 1 if str(raw).strip() in ("1", "true", "True") else 0
        elif kind == "score":
            try:
                normalized[name] = int(float(raw))
            except (ValueError, TypeError):
                normalized[name] = ""
        else:
            normalized[name] = str(raw).strip() if raw not in (None, "") else ""

    return {
        "values": normalized,
        "comment_is_negative": bool(data.get("comment_is_negative", False)),
    }

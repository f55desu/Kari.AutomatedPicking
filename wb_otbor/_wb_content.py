"""
Низкоуровневые функции WB Content API.
Инлайн-копия из wb-photo-report/wb_photo_report.py — сделано для того,
чтобы пакет был самодостаточным и легко собирался в exe через PyInstaller.
"""
from __future__ import annotations

import json
import os
import time
from pathlib import Path
from urllib import error, request


API_URL = "https://content-api.wildberries.ru/content/v2/get/cards/list"
REQUEST_DELAY_SECONDS = 0.65
MAX_RETRIES = 4
TOKEN_ENV_NAMES = (
    "WB_API_TOKEN",
    "WB_CONTENT_API_TOKEN",
    "WB_TOKEN",
)
RETRYABLE_STATUS_CODES = {408, 429, 500, 502, 503, 504}


class WBContentAPIError(Exception):
    """Ошибка при работе с WB Content API."""


def load_env_file(env_path: Path) -> None:
    """Читает переменные окружения из .env (если есть) и кладёт в os.environ."""
    env_path = Path(env_path)
    if not env_path.exists():
        return
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            os.environ.setdefault(key, value)


def get_api_token() -> str:
    """Возвращает токен WB из переменных окружения."""
    for env_name in TOKEN_ENV_NAMES:
        token = os.environ.get(env_name, "").strip()
        if token:
            return token
    expected = ", ".join(TOKEN_ENV_NAMES)
    raise WBContentAPIError(
        f"Не найден токен WB API. Укажите его в .env в одной из переменных: {expected}"
    )


def wb_post_json(token: str, payload: dict, timeout: int = 60) -> dict:
    """
    POST-запрос к WB Content API с ретраями и разбором ответа.
    """
    body = json.dumps(payload).encode("utf-8")
    req = request.Request(
        API_URL,
        data=body,
        headers={
            "Authorization": token,
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST",
    )

    last_error_message = "неизвестная ошибка"

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            with request.urlopen(req, timeout=timeout) as response:
                response_text = response.read().decode("utf-8")
        except error.HTTPError as exc:
            details = exc.read().decode("utf-8", errors="ignore")
            last_error_message = f"WB API HTTP {exc.code}. Детали: {details or exc.reason}"
            if exc.code in RETRYABLE_STATUS_CODES and attempt < MAX_RETRIES:
                time.sleep(max(REQUEST_DELAY_SECONDS, attempt * 1.5))
                continue
            raise WBContentAPIError(last_error_message) from exc
        except error.URLError as exc:
            last_error_message = f"Не удалось обратиться к WB API: {exc.reason}"
            if attempt < MAX_RETRIES:
                time.sleep(max(REQUEST_DELAY_SECONDS, attempt * 1.5))
                continue
            raise WBContentAPIError(last_error_message) from exc

        try:
            parsed = json.loads(response_text)
        except json.JSONDecodeError as exc:
            raise WBContentAPIError("WB API вернул ответ не в формате JSON.") from exc

        if parsed.get("error") is True:
            error_text = parsed.get("errorText") or "без текста ошибки"
            last_error_message = f"WB API сообщил об ошибке: {error_text}"
            if "timeout" in error_text.lower() and attempt < MAX_RETRIES:
                time.sleep(max(REQUEST_DELAY_SECONDS, attempt * 1.5))
                continue
            raise WBContentAPIError(last_error_message)

        return parsed

    raise WBContentAPIError(last_error_message)


def count_photos(card: dict) -> int:
    """Возвращает количество фото в объекте карточки."""
    photos = card.get("photos")
    if isinstance(photos, list):
        return len(photos)
    return 0

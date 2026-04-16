"""
Отправка готового xlsx-файла в Telegram.

Поведение:
  1. Пробуем прямое подключение к api.telegram.org.
  2. Если прямой доступ закрыт — ищем рабочий HTTP-прокси
     (из публичного списка monosans/proxy-list).
  3. Отправляем документ через python-telegram-bot с правильно
     прокинутым прокси (HTTPXRequest) — либо напрямую.

Файл .env читается из config.ENV_FILE (абсолютный путь), а не из CWD —
это важно для запуска из планировщика задач / exe.
"""
from __future__ import annotations

import asyncio
import inspect
import json
import logging
import os
from pathlib import Path
from urllib import request as urllib_request, error as urllib_error

from telegram import Bot
from telegram.request import HTTPXRequest

from . import config
from ._wb_content import load_env_file


logger = logging.getLogger(__name__)

# === Загрузка токена ===
# Делаем это на импорте, но БЕЗ падения при отсутствии — пусть падает
# только при реальной попытке отправки, с внятным сообщением.
load_env_file(config.ENV_FILE)


def _clean_token(value: str) -> str:
    """Убирает возможные кавычки вокруг значения."""
    v = (value or '').strip()
    if len(v) >= 2 and v[0] == v[-1] and v[0] in ("'", '"'):
        v = v[1:-1]
    return v


TOKEN = _clean_token(os.environ.get("TELEGRAM_BOT_TOKEN", ""))

# Дополнительно: пробуем второе имя переменной на случай, если в .env
# токен под другим ключом.
if not TOKEN:
    TOKEN = _clean_token(os.environ.get("TG_BOT_TOKEN", ""))


# === Прокси-обвязка через stdlib ===

PROXY_LIST_URL = "https://raw.githubusercontent.com/monosans/proxy-list/main/proxies/http.txt"
TELEGRAM_PING_URL_TEMPLATE = "https://api.telegram.org/bot{token}/getMe"


def _is_telegram_reachable_direct(token: str, timeout: int = 5) -> bool:
    """True, если api.telegram.org отвечает без прокси."""
    try:
        req = urllib_request.Request(TELEGRAM_PING_URL_TEMPLATE.format(token=token))
        with urllib_request.urlopen(req, timeout=timeout) as resp:
            return bool(json.loads(resp.read()).get("ok"))
    except Exception:
        return False


def _get_free_proxies() -> list[str]:
    """Скачивает публичный список HTTP-прокси."""
    try:
        req = urllib_request.Request(PROXY_LIST_URL, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib_request.urlopen(req, timeout=10) as resp:
            data = resp.read().decode('utf-8')
            proxies = [line.strip() for line in data.splitlines() if line.strip()]
            logger.info("Получено %d прокси из публичного списка.", len(proxies))
            return proxies
    except Exception as exc:
        logger.error("Не удалось скачать список прокси: %s", exc)
        return []


def _find_working_proxy(token: str, max_tries: int = 50) -> str | None:
    """Ищет первый прокси, через который Telegram отдаёт getMe=ok."""
    proxies = _get_free_proxies()
    ping_url = TELEGRAM_PING_URL_TEMPLATE.format(token=token)

    for i, p in enumerate(proxies[:max_tries], start=1):
        proxy_url = f"http://{p}"
        try:
            handler = urllib_request.ProxyHandler({'https': proxy_url, 'http': proxy_url})
            opener = urllib_request.build_opener(handler)
            with opener.open(urllib_request.Request(ping_url), timeout=5) as resp:
                if json.loads(resp.read()).get("ok"):
                    logger.info("Рабочий прокси найден (#%d): %s", i, proxy_url)
                    return proxy_url
        except Exception:
            continue

    logger.error("Среди первых %d прокси рабочего не нашлось.", max_tries)
    return None


# === HTTPXRequest с прокси, совместимый между версиями PTB ===

def _make_httpx_request(proxy_url: str | None = None) -> HTTPXRequest:
    """
    Создаёт HTTPXRequest с прокси (если задан).
    В разных версиях python-telegram-bot параметр называется
    по-разному: proxy / proxy_url. Выбираем автоматически.
    """
    kwargs = dict(read_timeout=120, write_timeout=120, connect_timeout=30)
    if proxy_url:
        sig = inspect.signature(HTTPXRequest)
        if 'proxy' in sig.parameters:
            kwargs['proxy'] = proxy_url
        elif 'proxy_url' in sig.parameters:
            kwargs['proxy_url'] = proxy_url
        else:
            # Фолбэк: проставляем через env vars (httpx их подхватит)
            os.environ['HTTPS_PROXY'] = proxy_url
            os.environ['HTTP_PROXY'] = proxy_url
    return HTTPXRequest(**kwargs)


# === Публичная точка входа ===

def telegram_sendFile(file_path, chat_id, message: str = "") -> bool:
    """
    Отправляет файл в Telegram chat_id. Возвращает True при успехе.
    """
    if not TOKEN:
        logger.error(
            "TELEGRAM_BOT_TOKEN не найден в %s. "
            "Проверьте, что переменная задана в .env.",
            config.ENV_FILE,
        )
        return False

    file_path = Path(file_path)
    if not file_path.exists():
        logger.error("Файл для отправки не существует: %s", file_path)
        return False

    # 1. Пробуем прямое подключение
    logger.info("Проверяю прямое подключение к Telegram API...")
    proxy_url: str | None = None
    if _is_telegram_reachable_direct(TOKEN):
        logger.info("Прямое подключение работает — прокси не нужен.")
    else:
        logger.warning("Прямое подключение к Telegram закрыто, ищу прокси...")
        proxy_url = _find_working_proxy(TOKEN)
        if not proxy_url:
            logger.error("Отправка невозможна: нет ни прямого доступа, ни рабочего прокси.")
            return False

    # 2. Отправляем через python-telegram-bot
    async def _send() -> None:
        req = _make_httpx_request(proxy_url)
        bot = Bot(token=TOKEN, request=req)
        try:
            with open(file_path, 'rb') as f:
                await bot.send_document(
                    chat_id=chat_id,
                    document=f,
                    caption=message,
                    filename=file_path.name,
                )
        finally:
            try:
                await bot.shutdown()
            except Exception:
                pass

    try:
        asyncio.run(_send())
        logger.info("Файл %s отправлен в chat %s.", file_path.name, chat_id)
        return True
    except Exception as exc:
        logger.error("Ошибка при отправке в Telegram: %s", exc)
        return False

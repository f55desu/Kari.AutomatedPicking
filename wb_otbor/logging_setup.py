"""
Централизованное логирование: пишет в консоль + ротирующийся файл logs/wb_otbor.log.

Во всех модулях пакета:
    from .logging_setup import get_logger
    logger = get_logger(__name__)
    logger.info(...)  # DEBUG/INFO/WARNING/ERROR/EXCEPTION

setup_logging() вызывается единожды при старте (в run_gui.py / run_otbor.py).
"""
from __future__ import annotations

import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

from . import config


LOG_DIR = config.BASE_DIR / 'logs'
LOG_FILE = LOG_DIR / 'wb_otbor.log'
_CONFIGURED = False


def setup_logging(console_level: int = logging.INFO,
                  file_level: int = logging.DEBUG) -> logging.Logger:
    """
    Инициализирует корневой логгер 'wb_otbor'. Идемпотентно: повторные вызовы
    не плодят обработчики.
    """
    global _CONFIGURED

    root = logging.getLogger('wb_otbor')
    root.setLevel(logging.DEBUG)

    if _CONFIGURED:
        return root

    # На всякий случай чистим предыдущие handlers
    for h in root.handlers[:]:
        root.removeHandler(h)

    fmt = logging.Formatter(
        '%(asctime)s [%(levelname)-7s] %(name)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )

    # --- Консольный handler ---
    try:
        ch = logging.StreamHandler(sys.stdout)
        ch.setLevel(console_level)
        ch.setFormatter(fmt)
        root.addHandler(ch)
    except Exception:
        pass  # нет stdout (windowed exe без консоли) — молча пропускаем

    # --- Файловый handler (с ротацией) ---
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        fh = RotatingFileHandler(
            LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=5,
            encoding='utf-8', delay=True,
        )
        fh.setLevel(file_level)
        fh.setFormatter(fmt)
        root.addHandler(fh)
    except Exception as exc:
        # Если не удалось открыть файл (права, сетевая шара) — ругаемся в консоль,
        # но не падаем.
        sys.stderr.write(f"[logging_setup] Не удалось открыть лог-файл {LOG_FILE}: {exc}\n")

    _CONFIGURED = True
    root.info("=" * 70)
    root.info(f"Лог-файл: {LOG_FILE}")
    return root


def get_logger(name: str) -> logging.Logger:
    """
    Возвращает логгер как поддерево 'wb_otbor'. Для любого имени модуля
    автоматически получаем иерархическое имя вроде 'wb_otbor.pipeline'.
    """
    if name.startswith('wb_otbor'):
        return logging.getLogger(name)
    # короткие имена вроде 'pipeline'
    return logging.getLogger(f'wb_otbor.{name}')

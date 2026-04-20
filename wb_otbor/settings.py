"""
Пользовательские настройки (settings.json, рядом с exe / в корне проекта).

Содержит:
  - Фильтры справочника (multi-select)
  - Числовые параметры отчёта (период, пороги Технички)
  - Кэш уникальных значений фильтров (подгружается по кнопке в GUI)

Файл settings.json создаётся автоматически при первом сохранении.
Если файла нет — возвращаются значения из DEFAULTS (совместимо со старым поведением).
"""
from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any

from . import config
from .logging_setup import get_logger


logger = get_logger('settings')
SETTINGS_FILE = config.BASE_DIR / 'settings.json'

# Категориальные фильтры (multi-select из справочника GuideAssortiment)
FILTER_KEYS = [
    'Бизнес-группа',
    'Розничный отдел',
    'Группа',
    'Сезон',
    'Бренд',
    'Ответственный за группу',
    'Коллекция',
]

# Маппинг ключа фильтра → SQL-колонка в [DBReport].[dbo].[GuideAssortiment] q
FILTER_SQL_COLUMN = {
    'Бизнес-группа':           'q.businessgroupru',
    'Розничный отдел':         'q.DEPARTMENTIDRU',
    'Группа':                  "CONCAT(q.retailgroup,' ',q.grpnameru)",
    'Сезон':                   'q.KAR_SEASONCODERU',
    'Бренд':                   'q.trademark',
    'Ответственный за группу': 'q.buyer',
    'Коллекция':               'q.KAR_ACTUALCOLLECTION',
}

# Числовые параметры (дефолты)
NUMERIC_DEFAULTS = {
    'period_days': 7,
    'offset_from_today': 1,
    'min_stock': 10,
    'min_distrib_percent': 20,
    'min_shows': 1200,
}

DEFAULTS = {
    'filters': {
        'Бизнес-группа':           ['Одежда для детей', 'Одежда и аксессуары'],
        'Розничный отдел':         [],   # пустой список = не фильтровать
        'Группа':                  [],
        'Сезон':                   [],
        'Бренд':                   [],
        'Ответственный за группу': [],
        'Коллекция':               ['2026SS'],
    },
    **NUMERIC_DEFAULTS,
    'unique_values_cache':   {},           # кэш значений из БД (заполняется по кнопке)
    'unique_values_updated': None,         # ISO-timestamp последнего обновления кэша
}


def load() -> dict:
    """Загружает настройки из settings.json (или DEFAULTS, если файла нет)."""
    if not SETTINGS_FILE.exists():
        logger.debug(f"{SETTINGS_FILE} отсутствует — используем DEFAULTS.")
        return deepcopy(DEFAULTS)
    try:
        data = json.loads(SETTINGS_FILE.read_text(encoding='utf-8'))
    except Exception:
        logger.exception(f"Не удалось распарсить {SETTINGS_FILE} — используем DEFAULTS.")
        return deepcopy(DEFAULTS)

    # Мержим с дефолтами — новые ключи появятся без потерь пользовательских значений
    result = deepcopy(DEFAULTS)
    if 'filters' in data and isinstance(data['filters'], dict):
        for k in FILTER_KEYS:
            if k in data['filters'] and isinstance(data['filters'][k], list):
                result['filters'][k] = list(data['filters'][k])
    for k in NUMERIC_DEFAULTS:
        if k in data:
            try:
                result[k] = int(data[k])
            except (TypeError, ValueError):
                pass
    if isinstance(data.get('unique_values_cache'), dict):
        result['unique_values_cache'] = data['unique_values_cache']
    if data.get('unique_values_updated'):
        result['unique_values_updated'] = data['unique_values_updated']
    return result


def save(data: dict) -> Path:
    """Сохраняет настройки в settings.json. Возвращает путь."""
    try:
        SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding='utf-8',
        )
        active = sum(1 for v in data.get('filters', {}).values() if v)
        logger.info(f"Настройки сохранены в {SETTINGS_FILE} "
                    f"(активных фильтров: {active})")
        return SETTINGS_FILE
    except Exception:
        logger.exception(f"Не удалось сохранить настройки в {SETTINGS_FILE}")
        raise


# --- Удобные геттеры ---

def get_filters() -> dict[str, list[str]]:
    return load()['filters']


def get_numeric(key: str) -> int:
    s = load()
    return int(s.get(key, NUMERIC_DEFAULTS[key]))


def get_unique_values() -> dict[str, list[str]]:
    """Кэш уникальных значений (из БД), по ключу-фильтру."""
    return load().get('unique_values_cache', {})


def update_unique_values(values: dict[str, list[str]]) -> None:
    """Обновляет кэш уникальных значений."""
    from datetime import datetime
    data = load()
    data['unique_values_cache'] = values
    data['unique_values_updated'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    save(data)


# --- Построение SQL WHERE из фильтров ---

def build_filter_where(filters: dict[str, list[str]], indent: str = '      ') -> str:
    """
    Превращает dict фильтров в AND-чейн вида:
        q.businessgroupru IN (N'...', N'...')
        AND q.DEPARTMENTIDRU IN (N'...')
    Пустой список = фильтр не применяется.
    Возвращает пустую строку если ни одного активного фильтра.
    """
    parts = []
    for key, values in filters.items():
        if not values:
            continue
        col = FILTER_SQL_COLUMN.get(key)
        if not col:
            continue
        # экранируем одинарные кавычки в значениях
        escaped = [str(v).replace("'", "''") for v in values]
        in_list = ', '.join(f"N'{v}'" for v in escaped)
        parts.append(f"{col} IN ({in_list})")
    if not parts:
        return ''
    return ('\n' + indent + 'AND ').join(parts)

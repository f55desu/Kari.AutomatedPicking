"""Работа с WB Content API + xlsx-кэш количества фото."""
import time
from pathlib import Path
from openpyxl import Workbook, load_workbook

from . import config
from ._wb_content import (
    load_env_file,
    get_api_token,
    wb_post_json,
    count_photos,
    REQUEST_DELAY_SECONDS,
)
from .logging_setup import get_logger


logger = get_logger('wb_api')


def save_photo_cache(photo_counts: dict, cache_path: Path = None) -> Path:
    """Сохраняет dict {str(nmID) -> raw photo count} в xlsx. Возвращает путь."""
    cache_path = Path(cache_path) if cache_path else config.PHOTO_CACHE_FILE
    try:
        cache_path.parent.mkdir(parents=True, exist_ok=True)

        wb_cache = Workbook()
        ws = wb_cache.active
        ws.title = "photo_counts"
        ws.append(["Артикул WB", "Количество фото"])

        def _sort_key(pair):
            k = pair[0]
            return int(k) if k.isdigit() else 0

        for nm_id, count in sorted(photo_counts.items(), key=_sort_key):
            ws.append([nm_id, int(count)])
        wb_cache.save(cache_path)
        logger.info(f"Фото-кэш сохранён: {cache_path.resolve()} ({len(photo_counts)} записей)")
        return cache_path.resolve()
    except Exception:
        logger.exception(f"Не удалось сохранить фото-кэш в {cache_path}")
        raise


def load_photo_cache(cache_path: Path = None) -> dict:
    """Загружает dict {str(nmID) -> raw photo count} из xlsx."""
    cache_path = cache_path or config.PHOTO_CACHE_FILE
    if not cache_path.exists():
        logger.error(f"Фото-кэш не найден: {cache_path}")
        raise FileNotFoundError(f"Кэш не найден: {cache_path}")
    try:
        wb_cache = load_workbook(cache_path)
        ws = wb_cache.active
        result = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            nm_id = str(row[0]).strip()
            count = int(row[1]) if row[1] is not None else 0
            result[nm_id] = count
        logger.info(f"Фото-кэш загружен: {cache_path} ({len(result)} записей)")
        return result
    except Exception:
        logger.exception(f"Ошибка при чтении фото-кэша {cache_path}")
        raise


def fetch_photos_from_api(target_nm_ids, timeout: int = 60, log=print) -> dict:
    """
    Batch-загрузка через cursor-пагинацию /content/v2/get/cards/list.

    WB Content API v2 НЕ поддерживает массив nmID в одном запросе.
    Единственный фильтр — textSearch (1 значение). Cursor-пагинация
    по всему каталогу — максимально быстрый легитимный метод.
    """
    load_env_file(config.ENV_FILE)
    token = get_api_token()

    target_set = set()
    for x in target_nm_ids:
        if x is None:
            continue
        s = str(x).strip()
        if not s or s.lower() == 'nan':
            continue
        try:
            target_set.add(str(int(float(s))))
        except (TypeError, ValueError):
            continue

    results = {}
    cursor_payload = {"limit": 100}
    total_fetched = 0
    page = 0

    log(f"Целевых артикулов WB: {len(target_set)}")
    start = time.time()

    while True:
        page += 1
        payload = {"settings": {"cursor": cursor_payload, "filter": {"withPhoto": -1}}}

        data = wb_post_json(token, payload, timeout)
        cards = data.get("cards", [])
        if not cards:
            break

        for card in cards:
            nm_id_str = str(card.get("nmID", ""))
            if nm_id_str in target_set:
                results[nm_id_str] = count_photos(card)

        total_fetched += len(cards)
        cursor_data = data.get("cursor", {})
        total_cards = cursor_data.get("total", "?")

        log(f"  Стр. {page}: +{len(cards)} карточек "
            f"(просмотрено: {total_fetched}/{total_cards}), "
            f"совпадений: {len(results)}/{len(target_set)}")

        if len(results) >= len(target_set):
            log(f"  Все {len(target_set)} артикулов найдены.")
            break

        next_updated = cursor_data.get("updatedAt")
        next_nmid = cursor_data.get("nmID")
        if not next_updated and not next_nmid:
            break

        cursor_payload = {
            "limit": 100,
            "updatedAt": next_updated or "",
            "nmID": next_nmid or 0,
        }
        time.sleep(REQUEST_DELAY_SECONDS)

    elapsed = time.time() - start
    not_found = target_set - set(results.keys())
    log(f"Итого: {page} запросов за {elapsed:.1f} сек. "
        f"Найдено: {len(results)}/{len(target_set)}.")
    if not_found:
        log(f"Не найдено: {len(not_found)} артикулов (проставится 0).")
        for nm in not_found:
            results[nm] = 0

    return results


def get_photo_counts(target_nm_ids, use_cache: bool = False, log=print) -> dict:
    """
    Основная точка входа.

      use_cache=True  + кэш существует  → читаем из xlsx-кэша, без обращения к API.
      use_cache=True  + кэша НЕТ         → API + сохранение кэша (защита от ошибки).
      use_cache=False                    → API + сохранение кэша.
    """
    cache_path = config.PHOTO_CACHE_FILE

    if use_cache:
        if cache_path.exists():
            log(f"Читаем фото-кэш: {cache_path}")
            logger.info(f"get_photo_counts: используем кэш {cache_path}")
            try:
                return load_photo_cache()
            except Exception as exc:
                logger.exception("Кэш повреждён — падаем на API")
                log(f"Кэш повреждён ({exc}), запрашиваем через API.")
        else:
            log(f"Кэш не найден ({cache_path}) — запрашиваем через WB API.")
            logger.info(f"Кэш отсутствует, use_cache=True → переключаемся на API")

    log("Запрос количества фото через WB API...")
    logger.info(f"get_photo_counts: WB API для {len(target_nm_ids)} nmID")
    try:
        results = fetch_photos_from_api(target_nm_ids, log=log)
    except Exception:
        logger.exception("Ошибка fetch_photos_from_api")
        raise

    # Сохраняем кэш БЕЗУСЛОВНО после каждого обращения к API
    try:
        saved_to = save_photo_cache(results)
        log(f"✓ Фото-кэш сохранён: {saved_to} ({len(results)} записей)")
    except Exception as exc:
        log(f"⚠ ВНИМАНИЕ: не удалось сохранить фото-кэш: {exc}")
        logger.exception("Не удалось сохранить фото-кэш — продолжаем без него")

    return results

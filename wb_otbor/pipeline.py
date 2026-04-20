"""Оркестратор: SQL → WB API → Excel → Telegram. С полным логированием."""
import time
import traceback
from pathlib import Path
import pandas as pd
from datetime import datetime, timedelta, date

from . import config
from . import sql_loader
from . import wb_api
from . import excel_writer
from . import telegram_bot_sender
from .logging_setup import get_logger


logger = get_logger('pipeline')


class PipelineStageError(Exception):
    """Ошибка выполнения одной из стадий пайплайна."""


def _log_both(gui_log, logger_level, msg):
    """Пишет в GUI-лог (через log=callable) и в файл (через logging)."""
    try:
        gui_log(msg)
    except Exception:
        pass
    logger_level(msg)


def run_full_pipeline(use_photo_cache: bool = False,
                      output_dir: Path = None,
                      log=print) -> Path:
    """
    Полный цикл: SQL → WB API → Excel → Telegram.
    Каждая стадия обёрнута в try/except: ошибки попадают и в GUI-лог, и в файл.
    """
    t_start = time.time()
    output_path = None

    logger.info("=" * 60)
    logger.info(f"Pipeline старт. use_photo_cache={use_photo_cache}, output_dir={output_dir}")

    # -------------------- Стадия 0: проверка шаблона -------------------
    try:
        if not config.TEMPLATE_FILE.exists():
            raise FileNotFoundError(
                f"Шаблон не найден: {config.TEMPLATE_FILE}. "
                f"Положите xlsx-шаблон в корень проекта."
            )
        _log_both(log, logger.info, f"Шаблон найден: {config.TEMPLATE_FILE}")
    except Exception as exc:
        _log_both(log, logger.error, f"СТАДИЯ 0 (проверка шаблона) FAIL: {exc}")
        logger.exception("Traceback стадии 0:")
        raise

    # -------------------- Стадия 1: SQL --------------------------------
    try:
        _log_both(log, logger.info, "=" * 60)
        _log_both(log, logger.info, "[1/4] Загрузка данных из SQL...")
        start_date, end_date = sql_loader.get_date_range()
        _log_both(log, logger.info,
                  f"Период: {start_date.strftime('%d.%m.%Y')} — {end_date.strftime('%d.%m.%Y')}")
        df_base = sql_loader.load_base_dataframe(start_date, end_date)
        _log_both(log, logger.info, f"[1/4] SQL OK: получено {len(df_base)} строк.")
    except Exception as exc:
        _log_both(log, logger.error, f"[1/4] SQL FAIL: {exc}")
        logger.exception("Traceback стадии 1:")
        raise PipelineStageError(f"SQL: {exc}") from exc

    # -------------------- Стадия 2: WB API (фото) ----------------------
    try:
        _log_both(log, logger.info, "=" * 60)
        _log_both(log, logger.info, "[2/4] Получение количества фото...")
        nm_ids = df_base['Артикул WB'].dropna().unique().tolist()
        photo_counts = wb_api.get_photo_counts(
            nm_ids,
            use_cache=use_photo_cache,
            log=lambda m: _log_both(log, logger.info, m),
        )

        def _photo_minus2(wb_id):
            if pd.isna(wb_id) or str(wb_id).strip() == '':
                return 0
            try:
                key = str(int(float(wb_id)))
            except (TypeError, ValueError):
                return 0
            raw = photo_counts.get(key, 0)
            return max(raw - 2, 0)

        df_base['Количество фото (-2 от скрипта)'] = df_base['Артикул WB'].apply(_photo_minus2)

        # Сортировка по Показам по убыванию
        df_base = df_base.sort_values('Показы', ascending=False).reset_index(drop=True)
        _log_both(log, logger.info, "[2/4] WB API OK (сортировка по Показам DESC применена).")
    except Exception as exc:
        _log_both(log, logger.error, f"[2/4] WB API FAIL: {exc}")
        logger.exception("Traceback стадии 2:")
        raise PipelineStageError(f"WB API: {exc}") from exc

    # -------------------- Стадия 3: Excel ------------------------------
    try:
        _log_both(log, logger.info, "=" * 60)
        _log_both(log, logger.info, "[3/4] Формирование Excel-файла...")
        output_dir = Path(output_dir) if output_dir else config.BASE_DIR
        output_dir.mkdir(parents=True, exist_ok=True)
        final_date = end_date + timedelta(days=1)
        output_path = output_dir / f'Отчёт WB о выбросах CTR, CR на {final_date.strftime("%d.%m")}.xlsx'
        excel_writer.write_otbor_file(
            df_base=df_base,
            end_date=end_date,
            template_path=config.TEMPLATE_FILE,
            output_path=output_path,
            log=lambda m: _log_both(log, logger.info, m),
        )
        _log_both(log, logger.info, f"[3/4] Excel OK: {output_path}")
    except Exception as exc:
        _log_both(log, logger.error, f"[3/4] Excel FAIL: {exc}")
        logger.exception("Traceback стадии 3:")
        raise PipelineStageError(f"Excel: {exc}") from exc

    # -------------------- Стадия 4: Telegram ---------------------------
    try:
        _log_both(log, logger.info, "=" * 60)
        _log_both(log, logger.info, "[4/4] Отправка в Telegram...")
        tg_id = getattr(config, 'TALDYKIN_ID', None) or getattr(config, 'F55_ID', None)
        tg_id2 = getattr(config, 'ANALYTICS_AUTO', None)
        if tg_id is None:
            raise ValueError("В config.py не задан chat_id (TALDYKIN_ID / F55_ID).")
        ok = telegram_bot_sender.telegram_sendFile(
            file_path=output_path, chat_id=tg_id, message=""
        )
        # ok = telegram_bot_sender.telegram_sendFile(
        #     file_path=output_path, chat_id=tg_id2, message=""
        # )
        if ok:
            _log_both(log, logger.info, f"[4/4] Telegram OK (chat_id={tg_id}).")
        else:
            _log_both(log, logger.warning,
                      f"[4/4] Telegram: отправка не удалась (смотрите лог). Файл всё равно создан.")
    except Exception as exc:
        # Отправка в Telegram — не критичная стадия: логируем, но не поднимаем
        _log_both(log, logger.warning,
                  f"[4/4] Telegram FAIL: {exc}. Файл создан, но не отправлен.")
        logger.exception("Traceback стадии 4:")

    elapsed = time.time() - t_start
    _log_both(log, logger.info, "=" * 60)
    _log_both(log, logger.info, f"ГОТОВО за {elapsed:.1f} сек. Файл: {output_path}")
    return output_path

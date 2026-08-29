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
from . import manager_mapping
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

    # --- Подтягиваем ФИО менеджера из 'Распределение категорий.xlsx' ---
    # Вставляем колонку СРАЗУ ПОСЛЕ 'Ответственный за группу', чтобы сохранить
    # порядок столбцов в итоговом Excel-файле.
    try:
        mapping = manager_mapping.load_mapping()
        df_base['ФИО менеджера'] = df_base['Группа'].apply(
            lambda g: manager_mapping.get_manager(g, mapping)
        )
        # Перемещаем колонку на позицию после 'Ответственный за группу'
        cols = list(df_base.columns)
        cols.remove('ФИО менеджера')
        insert_after = cols.index('Ответственный за группу') + 1
        cols.insert(insert_after, 'ФИО менеджера')
        df_base = df_base[cols]

        n_unknown = (df_base['ФИО менеджера'] == manager_mapping.UNKNOWN).sum()
        n_known = len(df_base) - n_unknown
        _log_both(log, logger.info,
                  f"ФИО менеджера подтянуто: {n_known}/{len(df_base)} строк, "
                  f"не определено — {n_unknown}.")
    except Exception as exc:
        # Не критично — просто проставим 'Не определено' везде и идём дальше
        _log_both(log, logger.warning,
                  f"Не удалось подтянуть ФИО менеджера: {exc}. "
                  f"Все строки получат '{manager_mapping.UNKNOWN}'.")
        logger.exception("Traceback шага подтягивания ФИО менеджера:")
        if 'ФИО менеджера' not in df_base.columns:
            df_base['ФИО менеджера'] = manager_mapping.UNKNOWN
            cols = list(df_base.columns)
            cols.remove('ФИО менеджера')
            insert_after = cols.index('Ответственный за группу') + 1
            cols.insert(insert_after, 'ФИО менеджера')
            df_base = df_base[cols]

    # -------------------- Стадия 2: WB API (фото + дата создания) ------
    try:
        _log_both(log, logger.info, "=" * 60)
        _log_both(log, logger.info, "[2/4] Получение количества фото и даты создания...")
        nm_ids = df_base['Артикул WB'].dropna().unique().tolist()
        photo_data = wb_api.get_photo_data(
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
            rec = photo_data.get(key, {})
            return max(int(rec.get('photos', 0)) - 2, 0)

        def _format_created_at(wb_id):
            """ISO-дата 'YYYY-MM-DDTHH:MM:SSZ' → 'DD.MM.YYYY' (с ведущими нулями).
            Если nmID не найден или дата пустая — возвращаем пустую строку."""
            if pd.isna(wb_id) or str(wb_id).strip() == '':
                return ''
            try:
                key = str(int(float(wb_id)))
            except (TypeError, ValueError):
                return ''
            iso = (photo_data.get(key) or {}).get('created_at', '') or ''
            if not iso:
                return ''
            # Пытаемся разобрать. Формат WB: '2024-01-15T10:30:00Z' или с микросекундами.
            from datetime import datetime as _dt
            for fmt in ('%Y-%m-%dT%H:%M:%S.%fZ', '%Y-%m-%dT%H:%M:%SZ',
                        '%Y-%m-%dT%H:%M:%S.%f', '%Y-%m-%dT%H:%M:%S'):
                try:
                    d = _dt.strptime(iso, fmt)
                    return d.strftime('%d.%m.%Y')
                except ValueError:
                    continue
            # Если не распарсили — оставляем как есть (для отладки)
            return iso

        df_base['Количество фото (-2 от скрипта)'] = df_base['Артикул WB'].apply(_photo_minus2)
        df_base['Дата создания на WB']             = df_base['Артикул WB'].apply(_format_created_at)

        # Сортировка по Показам по убыванию
        df_base = df_base.sort_values('Показы', ascending=False).reset_index(drop=True)

        n_with_date = (df_base['Дата создания на WB'] != '').sum()
        _log_both(log, logger.info,
                  f"[2/4] WB API OK. Дата создания подтянута для "
                  f"{n_with_date}/{len(df_base)} строк.")
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
        tg_id3 = getattr(config, 'ANALYTICS_AUTO2', None)
        if tg_id is None:
            raise ValueError("В config.py не задан chat_id (TALDYKIN_ID / F55_ID).")
        ok = telegram_bot_sender.telegram_sendFile(
            file_path=output_path, chat_id=tg_id, message=""
        )
        # ok = telegram_bot_sender.telegram_sendFile(
        #     file_path=output_path, chat_id=tg_id2, message=""
        # )
        ok = telegram_bot_sender.telegram_sendFile(
            file_path=output_path, chat_id=tg_id3, message=""
        )
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

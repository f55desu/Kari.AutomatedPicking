"""Оркестратор: SQL → WB API → Excel."""
import time
from pathlib import Path
import pandas as pd

from . import config
from . import sql_loader
from . import wb_api
from . import excel_writer
from . import telegram_bot_sender

def run_full_pipeline(use_photo_cache: bool = False,
                      output_dir: Path = None,
                      log=print) -> Path:
    """
    Полный цикл: загрузка данных из SQL, получение фото, запись в Excel.

    Параметры:
      use_photo_cache — читать количество фото из photo_cache.xlsx
                        вместо запроса к WB API (для отладки).
      output_dir      — куда сохранять итоговый файл.
                        По умолчанию — рядом с шаблоном.
      log             — функция логирования.

    Возвращает путь к созданному файлу.
    """
    t_start = time.time()

    # 1. SQL
    log("=" * 60)
    log("[1/4] Загрузка данных из SQL...")
    start_date, end_date = sql_loader.get_date_range()
    log(f"Период: {start_date.strftime('%d.%m.%Y')} — {end_date.strftime('%d.%m.%Y')}")
    df_base = sql_loader.load_base_dataframe(start_date, end_date)
    log(f"Получено {len(df_base)} строк из SQL.")

    # 2. WB API (фото)
    log("=" * 60)
    log("[2/4] Получение количества фото...")
    nm_ids = df_base['Артикул WB'].dropna().unique().tolist()
    photo_counts = wb_api.get_photo_counts(
        nm_ids,
        use_cache=use_photo_cache,
        log=log,
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
    log("Данные отсортированы по Показам (по убыванию).")

    # 3. Excel
    log("=" * 60)
    log("[3/4] Формирование Excel-файла...")
    output_dir = Path(output_dir) if output_dir else config.BASE_DIR
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f'Отбор {end_date.strftime("%d.%m")}.xlsx'
    excel_writer.write_otbor_file(
        df_base=df_base,
        end_date=end_date,
        template_path=config.TEMPLATE_FILE,
        output_path=output_path,
        log=log,
    )
    # 4. Telegram
    log("=" * 60)
    log("[4/4] Отправка в Telegram...")
    telegram_bot_sender.telegram_sendFile(file_path=output_path, chat_id=config.TALDYKIN_ID, message=f"Ваш файл отбора WB на {end_date.strftime('%d.%m')}")
    # telegram_bot_sender.telegram_sendFile(file_path=output_path, chat_id=config.ANALYTICS_AUTO, message=f"Ваш файл отбора WB на {end_date.strftime('%d.%m')}")

    elapsed = time.time() - t_start
    log("=" * 60)
    log(f"ГОТОВО за {elapsed:.1f} сек. Файл: {output_path}")
    return output_path

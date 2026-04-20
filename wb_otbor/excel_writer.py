"""Запись итогового файла Отбор на основе xlsx-шаблона."""
from copy import copy
from datetime import datetime
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from . import config
from . import settings as app_settings
from .logging_setup import get_logger


logger = get_logger('excel_writer')


def _append_applied_filters_sheet(wb, end_date, output_date: str):
    """
    Добавляет второй лист 'Применённые фильтры' с текущими настройками
    (фильтры справочника + числовые параметры + даты периода).
    """
    s = app_settings.load()

    # Удаляем старый лист, если есть (при пересохранении)
    for existing in list(wb.sheetnames):
        if existing == 'Применённые фильтры':
            del wb[existing]

    ws = wb.create_sheet('Применённые фильтры')

    thin = Side(style='thin', color='BFBFBF')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill('solid', fgColor='1F4E78')
    header_font = Font(bold=True, color='FFFFFF', name='Calibri')
    center = Alignment(horizontal='center', vertical='center')
    left = Alignment(horizontal='left', vertical='top', wrap_text=True)

    def _write_header(row, title):
        cell = ws.cell(row=row, column=1, value=title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
        ws.cell(row=row, column=2).fill = header_fill
        ws.cell(row=row, column=2).border = border
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)

    # --- Шапка с датой формирования ---
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    ws['A1'] = 'Параметры отчёта WB'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws.merge_cells('A1:B1')
    ws['A2'] = f'Дата формирования:'
    ws['B2'] = now

    # --- Период отчёта ---
    period_days = s.get('period_days', 7)
    offset_days = s.get('offset_from_today', 1)
    from datetime import timedelta
    start_date = end_date - timedelta(days=period_days - 1)
    ws['A3'] = 'Период:'
    ws['B3'] = f'{start_date.strftime("%d.%m.%Y")} — {end_date.strftime("%d.%m.%Y")}'
    ws['A4'] = 'Длина периода, дней:'
    ws['B4'] = period_days
    ws['A5'] = 'Сдвиг от сегодня, дней:'
    ws['B5'] = offset_days

    # --- Пороги Технички ---
    _write_header(7, 'Пороги «Технички» / «Отбора»')
    rows_th = [
        ('Мин. остаток (шт.):', s.get('min_stock', 10)),
        ('Мин. дистрибуция (%):', s.get('min_distrib_percent', 20)),
        ('Мин. показы:', s.get('min_shows', 1200)),
    ]
    for i, (lbl, val) in enumerate(rows_th, start=8):
        ws.cell(row=i, column=1, value=lbl)
        ws.cell(row=i, column=2, value=val)

    # --- Фильтры справочника ---
    start_row = 12
    _write_header(start_row, 'Фильтры справочника')
    row = start_row + 1

    filters = s.get('filters', {})
    for key in app_settings.FILTER_KEYS:
        values = filters.get(key, [])
        ws.cell(row=row, column=1, value=key).font = Font(bold=True)
        if values:
            val_text = '\n'.join(values) if len(values) <= 30 else \
                       '\n'.join(values[:30]) + f'\n…ещё {len(values) - 30}'
            ws.cell(row=row, column=2, value=val_text)
        else:
            ws.cell(row=row, column=2, value='(все — фильтр не задан)').font = \
                Font(italic=True, color='999999')
        ws.cell(row=row, column=2).alignment = left
        row += 1

    # Ширина колонок
    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['B'].width = 80

    # Заморозим шапку
    ws.freeze_panes = 'A2'

    logger.info("Второй лист 'Применённые фильтры' добавлен.")


def write_otbor_file(df_base: pd.DataFrame, end_date, template_path: Path,
                     output_path: Path, log=print) -> Path:
    """
    Заполняет шаблон данными из df_base и сохраняет под output_path.
    Второй лист 'Применённые фильтры' с текущими настройками добавляется автоматически.
    """
    logger.info(f"write_otbor_file: template={template_path}, output={output_path}, "
                f"rows={len(df_base)}")

    # --- Проверка входных данных ---
    if not Path(template_path).exists():
        logger.error(f"Шаблон не найден: {template_path}")
        raise FileNotFoundError(f"Шаблон не найден: {template_path}")

    required_cols = {'Артикул', 'Показы', 'Клики', 'Заказы',
                     'Количество фото (-2 от скрипта)'}
    missing = required_cols - set(df_base.columns)
    if missing:
        logger.error(f"В df_base отсутствуют столбцы: {missing}")
        raise ValueError(f"В df_base отсутствуют обязательные столбцы: {missing}")

    try:
        wb = load_workbook(template_path)
        ws = wb.active
    except Exception:
        logger.exception(f"Не удалось открыть шаблон {template_path}")
        raise

    output_date = end_date.strftime('%d.%m')

    # Сохраняем стили из шаблонной строки 3
    row3_styles = {}
    for col in range(1, 25):
        cell = ws.cell(row=3, column=col)
        row3_styles[col] = {
            'font': copy(cell.font),
            'fill': copy(cell.fill),
            'border': copy(cell.border),
            'alignment': copy(cell.alignment),
            'number_format': cell.number_format,
        }

    # Очищаем все строки данных
    for row in range(3, ws.max_row + 1):
        for col in range(1, 25):
            ws.cell(row=row, column=col).value = None

    # Имена столбцов из df_base
    остаток_col = [c for c in df_base.columns if 'Остаток' in c][0]
    distrib_col = [c for c in df_base.columns if 'Дистрибуция' in c][0]

    # Пороги для Технички/Отбора — читаем из settings.json
    min_stock = app_settings.get_numeric('min_stock')
    min_distrib = app_settings.get_numeric('min_distrib_percent')
    min_shows = app_settings.get_numeric('min_shows')

    last_data_row = 2 + len(df_base)

    log(f"Запись {len(df_base)} строк в Excel...")

    for idx, row_data in df_base.reset_index(drop=True).iterrows():
        r = 3 + idx

        # A–I: справочные данные
        ws.cell(row=r, column=1).value = str(row_data['Артикул'])
        ws.cell(row=r, column=2).value = (
            str(int(float(row_data['Артикул WB'])))
            if pd.notna(row_data['Артикул WB']) else ''
        )
        ws.cell(row=r, column=3).value = row_data['Бизнес-группа']
        ws.cell(row=r, column=4).value = row_data['Розничный отдел']
        ws.cell(row=r, column=5).value = row_data['Группа']
        ws.cell(row=r, column=6).value = row_data['Сезон']
        ws.cell(row=r, column=7).value = row_data['Бренд']
        ws.cell(row=r, column=8).value = row_data['Ответственный за группу']
        ws.cell(row=r, column=9).value = row_data['Коллекция']

        # J–L: маркетинг
        ws.cell(row=r, column=10).value = int(row_data['Показы'])
        ws.cell(row=r, column=11).value = int(row_data['Клики'])
        ws.cell(row=r, column=12).value = int(row_data['Заказы'])

        # M, N: формулы CTR и Конверсии
        ws.cell(row=r, column=13).value = f'=IFERROR(K{r}/J{r},0)'
        ws.cell(row=r, column=14).value = f'=IFERROR(L{r}/K{r},0)'

        # O: Остаток
        ws.cell(row=r, column=15).value = int(row_data[остаток_col])

        # P: Дистрибуция (как decimal для формата 0.0%)
        ws.cell(row=r, column=16).value = row_data[distrib_col] / 100

        # Q и R: Техничка и Отбор для поиска
        stock = row_data[остаток_col]
        distrib = row_data[distrib_col]
        shows = row_data['Показы']

        if stock < min_stock or distrib < min_distrib:
            ws.cell(row=r, column=18).value = 'Мало остатка'
            ws.cell(row=r, column=17).value = 0
        elif shows < min_shows:
            ws.cell(row=r, column=18).value = 'Мало показов'
            ws.cell(row=r, column=17).value = 0
        else:
            ws.cell(row=r, column=18).value = (
                f'=IF(OR(M{r}<$S$1*0.5,N{r}<$T$1*0.5),'
                f'"код для проверки","нормальный код")'
            )
            ws.cell(row=r, column=17).value = 1

        # S–W: формулы отбора
        ws.cell(row=r, column=19).value = (
            f'=IF(M{r}<$S$1*0.5,"Низшая четверть",'
            f'IF(M{r}<$S$1,"Вторая четверть","Больше половины"))'
        )
        ws.cell(row=r, column=20).value = (
            f'=IF(N{r}<$T$1*0.5,"Низшая четверть",'
            f'IF(N{r}<$T$1,"Вторая четверть","Больше половины"))'
        )
        ws.cell(row=r, column=21).value = (
            f'=IF(OR(M{r}<SUMIFS(K:K,Q:Q,1,E:E,E{r})/SUMIFS(J:J,Q:Q,1,E:E,E{r})*0.5,'
            f'N{r}<SUMIFS(L:L,Q:Q,1,E:E,E{r})/SUMIFS(K:K,Q:Q,1,E:E,E{r})*0.5),'
            f'"код для проверки","нормальный код")'
        )
        ws.cell(row=r, column=22).value = (
            f'=IF(M{r}<SUMIFS(K:K,Q:Q,1,E:E,E{r})/SUMIFS(J:J,Q:Q,1,E:E,E{r})*0.5,'
            f'"Низшая четверть",IF(M{r}<SUMIFS(K:K,Q:Q,1,E:E,E{r})/SUMIFS(J:J,Q:Q,1,E:E,E{r}),'
            f'"Вторая четверть","Больше половины"))'
        )
        ws.cell(row=r, column=23).value = (
            f'=IF(N{r}<SUMIFS(L:L,Q:Q,1,E:E,E{r})/SUMIFS(K:K,Q:Q,1,E:E,E{r})*0.5,'
            f'"Низшая четверть",IF(N{r}<SUMIFS(L:L,Q:Q,1,E:E,E{r})/SUMIFS(K:K,Q:Q,1,E:E,E{r}),'
            f'"Вторая четверть","Больше половины"))'
        )

        # X: Количество фото (-2)
        ws.cell(row=r, column=24).value = int(row_data['Количество фото (-2 от скрипта)'])

        # Применяем стили из шаблонной строки 3
        for col in range(1, 25):
            cell = ws.cell(row=r, column=col)
            style = row3_styles.get(col)
            if style:
                cell.font = copy(style['font'])
                cell.fill = copy(style['fill'])
                cell.border = copy(style['border'])
                cell.alignment = copy(style['alignment'])
                cell.number_format = style['number_format']

    # Обновляем формулы строки 1 под новый диапазон
    ws['J1'] = f'=SUBTOTAL(9,J3:J{last_data_row})'
    ws['K1'] = f'=SUBTOTAL(9,K3:K{last_data_row})'
    ws['L1'] = f'=SUBTOTAL(9,L3:L{last_data_row})'
    ws['M1'] = '=K1/J1'
    ws['N1'] = '=L1/K1'
    ws['S1'] = '=SUMIF(Q:Q,1,K:K)/SUMIF(Q:Q,1,J:J)'
    ws['T1'] = '=SUMIF(Q:Q,1,L:L)/SUMIF(Q:Q,1,K:K)'

    # Обновляем заголовок столбца Остаток
    ws.cell(row=2, column=15).value = f'Остаток на {output_date}'

    # Добавляем второй лист с применёнными фильтрами
    try:
        _append_applied_filters_sheet(wb, end_date, output_date)
    except Exception as exc:
        log(f"⚠ Не удалось добавить лист с фильтрами: {exc}")
        logger.exception("Ошибка при добавлении листа 'Применённые фильтры'")
        # Не падаем — основной лист важнее

    try:
        wb.save(output_path)
        logger.info(f"Excel-файл сохранён: {output_path}")
        log(f"Файл сохранён: {output_path}")
    except Exception:
        logger.exception(f"Не удалось сохранить {output_path}")
        raise
    return output_path

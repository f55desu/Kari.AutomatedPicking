"""Запись итогового файла Отбор на основе xlsx-шаблона."""
from copy import copy
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

from . import config


def write_otbor_file(df_base: pd.DataFrame, end_date, template_path: Path,
                     output_path: Path, log=print) -> Path:
    """
    Заполняет шаблон данными из df_base и сохраняет под output_path.

    df_base должен содержать столбец 'Количество фото (-2 от скрипта)'.
    Сортировка должна быть применена ДО вызова.
    """
    wb = load_workbook(template_path)
    ws = wb.active

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

        if stock < config.MIN_STOCK or distrib < config.MIN_DISTRIB_PERCENT:
            ws.cell(row=r, column=18).value = 'Мало остатка'
            ws.cell(row=r, column=17).value = 0
        elif shows < config.MIN_SHOWS:
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

    wb.save(output_path)
    log(f"Файл сохранён: {output_path}")
    return output_path

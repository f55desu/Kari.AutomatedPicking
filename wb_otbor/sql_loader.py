"""Загрузка основного датафрейма из SQL."""
from datetime import date, timedelta
import pandas as pd
from sqlalchemy import create_engine

from . import config


def connect_to_sql(server: str, database: str):
    connection_string = (
        f"mssql+pyodbc://{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server"
        "&trusted_connection=yes"
    )
    return create_engine(connection_string)


def get_date_range(offset_days: int = None, period_days: int = None):
    """
    Возвращает (start_date, end_date).
    end_date = today - offset_days
    start_date = end_date - (period_days - 1)
    """
    offset_days = offset_days if offset_days is not None else config.OFFSET_FROM_TODAY
    period_days = period_days if period_days is not None else config.PERIOD_DAYS
    end = date.today() - timedelta(days=offset_days)
    start = end - timedelta(days=period_days - 1)
    return start, end


def build_query(start_date, end_date) -> str:
    start_str = start_date.strftime('%Y-%m-%d')
    end_str = end_date.strftime('%Y-%m-%d')
    end_label = end_date.strftime('%d.%m')
    # bg_list = ", ".join(f"N'{g}'" for g in config.BUSINESS_GROUPS)

    return f"""
WITH ref AS (
    SELECT DISTINCT
        q.itemid                                AS [Артикул],
        b.inventsizeid                          AS [Размер],
        q.businessgroupru                       AS [Бизнес-группа],
        q.DEPARTMENTIDRU                        AS [Розничный отдел],
        CONCAT(q.retailgroup,' ',q.grpnameru)   AS [Группа],
        q.KAR_SEASONCODERU                      AS [Сезон],
        q.trademark                             AS [Бренд],
        q.buyer                                 AS [Ответственный за группу],
        q.KAR_ACTUALCOLLECTION                  AS [Коллекция]
    FROM [DBReport].[dbo].[GuideAssortiment] q
    INNER JOIN [DynamicsAx1].[dbo].[INVENTITEMBARCODE] a
        ON q.itemid = a.itemid AND a.dataareaid = 'vrt'
    INNER JOIN [DynamicsAx1].[dbo].[INVENTDIM] b
        ON a.inventdimid = b.inventdimid AND b.dataareaid = 'vrt'
    INNER JOIN [DynamicsAx1].[dbo].[INVENTTABLE] c
        ON q.itemid = c.itemid AND c.dataareaid = 'vrt' AND c.itemgroupid = 'Goods'
    WHERE q.businessgroupru IN (N'Одежда для детей', N'Одежда и аксессуары')
      AND q.KAR_ACTUALCOLLECTION = '{config.COLLECTION}'
),
total_sizes AS (
    -- Одноразмерные товары (сумки, ремни, аксессуары) хранятся с NULL/пустым
    -- inventsizeid. Если после фильтрации "непустых" размеров COUNT=0, считаем
    -- как 1 (товар существует в одном варианте). Аналог Python-логики
    -- `len(sizes.dropna().unique()) if len(sizes.dropna()) > 0 else 1`.
    SELECT [Артикул],
        ISNULL(
            NULLIF(
                COUNT(DISTINCT CASE WHEN LTRIM(RTRIM([Размер])) <> '' THEN [Размер] END),
                0
            ),
            1
        ) AS [Всего размеров]
    FROM ref GROUP BY [Артикул]
),
attrs AS (
    SELECT [Артикул],
        MAX([Бизнес-группа]) AS [Бизнес-группа],
        MAX([Розничный отдел]) AS [Розничный отдел],
        MAX([Группа]) AS [Группа],
        MAX([Сезон]) AS [Сезон],
        MAX([Бренд]) AS [Бренд],
        MAX([Ответственный за группу]) AS [Ответственный за группу],
        MAX([Коллекция]) AS [Коллекция]
    FROM ref GROUP BY [Артикул]
),
wb AS (
    SELECT ITEMID, MIN(NMID) AS [Артикул WB]
    FROM [DBPartners].[dbo].[WblmRepGetNomenclatureWildberries]
    GROUP BY ITEMID
),
articles_in_period AS (
    -- Артикул попадает в отчёт, если он:
    --   (1) имел qte > 0 на WB в период (реально был в продаже), ИЛИ
    --   (2) имел маркетинговые события (показы/клики/заказы) за период —
    --       даже если на складе ничего нет (распродан, но карточка активна).
    -- Раньше использовался только (1), из-за чего терялись распроданные товары
    -- с активным маркетингом (часто у них показов больше всего).
    SELECT DISTINCT itemid AS [Артикул]
    FROM [DBPartners].[dbo].[WblmRepGetStockWildberries]
    WHERE dt BETWEEN '{start_str}' AND '{end_str}'
      --AND qte > 0
    UNION
    SELECT DISTINCT wb_n.ITEMID AS [Артикул]
    FROM [DBReport].[mp].[wb_sales_funnel_lk] mk
    INNER JOIN [DBPartners].[dbo].[WblmRepGetNomenclatureWildberries] wb_n
        ON wb_n.NMID = mk.[Артикул WB]
    WHERE mk.[Дата] BETWEEN '{start_str}' AND '{end_str}'
      AND wb_n.ITEMID IS NOT NULL
),
stock_end AS (
    SELECT s.itemid AS [Артикул],
           ISNULL(SUM(ISNULL(s.qte, 0)), 0) AS [Остаток]
    FROM [DBPartners].[dbo].[WblmRepGetStockWildberries] s
    WHERE s.dt = '{end_str}'
    GROUP BY s.itemid
),
agg_sizes AS (
    -- Для одноразмерных товаров INVENTSIZEID может быть NULL/'' → COUNT DISTINCT=0.
    -- Но если товар присутствует в stock с qte>0, значит "1 размер" на агрегаторе.
    -- Поэтому 0 → 1 (такая же защита как в total_sizes).
    SELECT s.itemid AS [Артикул],
        ISNULL(
            NULLIF(
                COUNT(DISTINCT CASE WHEN LTRIM(RTRIM(ISNULL(s.INVENTSIZEID, ''))) <> ''
                                    THEN s.INVENTSIZEID END),
                0
            ),
            1
        ) AS [Размеров на агрегаторе]
    FROM [DBPartners].[dbo].[WblmRepGetStockWildberries] s
    WHERE s.dt = '{end_str}' AND s.qte > 0
    GROUP BY s.itemid
),
marketing AS (
    -- ISNULL вокруг агрегатов, чтобы не ловить warning
    -- "Null value is eliminated by an aggregate or other SET operation"
    SELECT
        a.[Артикул WB]                                      AS [Артикул WB],
        ISNULL(SUM(ISNULL(a.[Показы], 0)), 0)               AS [Показы],
        ISNULL(SUM(ISNULL(a.[Кол-во переходов в карточку товара], 0)), 0) AS [Клики],
        ISNULL(SUM(ISNULL(a.[Заказали товаров, шт], 0)), 0) AS [Заказы]
    FROM [DBReport].[mp].[wb_sales_funnel_lk] a
    WHERE a.[Дата] BETWEEN '{start_str}' AND '{end_str}'
    GROUP BY a.[Артикул WB]
)
SELECT
    a.[Артикул],
    wb.[Артикул WB],
    a.[Бизнес-группа],
    a.[Розничный отдел],
    a.[Группа],
    a.[Сезон],
    a.[Бренд],
    a.[Ответственный за группу],
    a.[Коллекция],
    ISNULL(m.[Показы], 0) AS [Показы],
    ISNULL(m.[Клики], 0) AS [Клики],
    ISNULL(m.[Заказы], 0) AS [Заказы],
    -- DECIMAL(9,2) вместо (5,2): (5,2) вмещает максимум 999.99, а Конверсия
    -- при данных wb_sales_funnel_lk (Заказы могут быть много больше Кликов
    -- из-за многоштучных корзин) иногда выходит за эту границу → overflow.
    -- (9,2) поддерживает значения до 9,999,999.99 — с большим запасом.
    CAST(CASE WHEN ISNULL(m.[Показы], 0) = 0 THEN 0
         ELSE ISNULL(m.[Клики], 0) * 100.0 / m.[Показы] END AS DECIMAL(9, 2)) AS [CTR, %],
    CAST(CASE WHEN ISNULL(m.[Клики], 0) = 0 THEN 0
         ELSE ISNULL(m.[Заказы], 0) * 100.0 / m.[Клики] END AS DECIMAL(9, 2)) AS [Конверсия клики в заказы, %],
    ISNULL(st.[Остаток], 0) AS [Остаток на {end_label}],
    CAST(CASE WHEN ISNULL(ts.[Всего размеров], 0) = 0 THEN 0
         ELSE ISNULL(ag.[Размеров на агрегаторе], 0) * 100.0 / ts.[Всего размеров]
         END AS DECIMAL(9, 2)) AS [Дистрибуция, %]
FROM articles_in_period ap
INNER JOIN attrs a       ON ap.[Артикул] = a.[Артикул]
LEFT JOIN wb             ON a.[Артикул] = wb.ITEMID
LEFT JOIN total_sizes ts ON a.[Артикул] = ts.[Артикул]
LEFT JOIN agg_sizes ag   ON a.[Артикул] = ag.[Артикул]
LEFT JOIN stock_end st   ON a.[Артикул] = st.[Артикул]
LEFT JOIN marketing m    ON wb.[Артикул WB] = m.[Артикул WB]
ORDER BY a.[Артикул]
"""


def build_query_ref_based(start_date, end_date) -> str:
    """
    Альтернативный запрос: В ОСНОВЕ — справочник (ref).
    Все артикулы коллекции из справочника попадают в отчёт,
    а сток, маркетинг и дистрибуция подтягиваются к ним через LEFT JOIN.

    Отличие от build_query:
      - Старый: FROM articles_in_period INNER JOIN attrs
                (в отчёт только те, кто был в стоке/маркетинге)
      - Новый:  FROM attrs (справочник = истина, остальное LEFT JOIN)
                (все артикулы коллекции, даже без активности)
    """
    start_str = start_date.strftime('%Y-%m-%d')
    end_str = end_date.strftime('%Y-%m-%d')
    end_label = end_date.strftime('%d.%m')

    return f"""
WITH ref AS (
    SELECT DISTINCT
        q.itemid                                AS [Артикул],
        b.inventsizeid                          AS [Размер],
        q.businessgroupru                       AS [Бизнес-группа],
        q.DEPARTMENTIDRU                        AS [Розничный отдел],
        CONCAT(q.retailgroup,' ',q.grpnameru)   AS [Группа],
        q.KAR_SEASONCODERU                      AS [Сезон],
        q.trademark                             AS [Бренд],
        q.buyer                                 AS [Ответственный за группу],
        q.KAR_ACTUALCOLLECTION                  AS [Коллекция]
    FROM [DBReport].[dbo].[GuideAssortiment] q
    INNER JOIN [DynamicsAx1].[dbo].[INVENTITEMBARCODE] a
        ON q.itemid = a.itemid AND a.dataareaid = 'vrt'
    INNER JOIN [DynamicsAx1].[dbo].[INVENTDIM] b
        ON a.inventdimid = b.inventdimid AND b.dataareaid = 'vrt'
    INNER JOIN [DynamicsAx1].[dbo].[INVENTTABLE] c
        ON q.itemid = c.itemid AND c.dataareaid = 'vrt' AND c.itemgroupid = 'Goods'
    WHERE q.businessgroupru IN (N'Одежда для детей', N'Одежда и аксессуары')
      AND q.KAR_ACTUALCOLLECTION = '{config.COLLECTION}'
),
-- Атрибуты: одна строка на артикул (основа отчёта)
attrs AS (
    SELECT [Артикул],
        MAX([Бизнес-группа])           AS [Бизнес-группа],
        MAX([Розничный отдел])         AS [Розничный отдел],
        MAX([Группа])                  AS [Группа],
        MAX([Сезон])                   AS [Сезон],
        MAX([Бренд])                   AS [Бренд],
        MAX([Ответственный за группу]) AS [Ответственный за группу],
        MAX([Коллекция])               AS [Коллекция]
    FROM ref GROUP BY [Артикул]
),
-- Всего размеров по справочнику (минимум 1 для одноразмерных)
total_sizes AS (
    SELECT [Артикул],
        ISNULL(
            NULLIF(
                COUNT(DISTINCT CASE WHEN LTRIM(RTRIM([Размер])) <> '' THEN [Размер] END),
                0
            ),
            1
        ) AS [Всего размеров]
    FROM ref GROUP BY [Артикул]
),
-- Маппинг Артикул → Артикул WB
wb AS (
    SELECT ITEMID, MIN(NMID) AS [Артикул WB]
    FROM [DBPartners].[dbo].[WblmRepGetNomenclatureWildberries]
    GROUP BY ITEMID
),
-- Остаток на конец периода
stock_end AS (
    SELECT s.itemid AS [Артикул],
           ISNULL(SUM(ISNULL(s.qte, 0)), 0) AS [Остаток]
    FROM [DBPartners].[dbo].[WblmRepGetStockWildberries] s
    WHERE s.dt = '{end_str}'
    GROUP BY s.itemid
),
-- Размеры с остатком > 0 на конец периода
agg_sizes AS (
    SELECT s.itemid AS [Артикул],
        ISNULL(
            NULLIF(
                COUNT(DISTINCT CASE WHEN LTRIM(RTRIM(ISNULL(s.INVENTSIZEID, ''))) <> ''
                                    THEN s.INVENTSIZEID END),
                0
            ),
            1
        ) AS [Размеров на агрегаторе]
    FROM [DBPartners].[dbo].[WblmRepGetStockWildberries] s
    WHERE s.dt = '{end_str}' AND s.qte > 0
    GROUP BY s.itemid
),
-- Маркетинг за период
marketing AS (
    SELECT
        a.[Артикул WB]                                                       AS [Артикул WB],
        ISNULL(SUM(ISNULL(a.[Показы], 0)), 0)                               AS [Показы],
        ISNULL(SUM(ISNULL(a.[Кол-во переходов в карточку товара], 0)), 0)    AS [Клики],
        ISNULL(SUM(ISNULL(a.[Заказали товаров, шт], 0)), 0)                 AS [Заказы]
    FROM [DBReport].[mp].[wb_sales_funnel_lk] a
    WHERE a.[Дата] BETWEEN '{start_str}' AND '{end_str}'
    GROUP BY a.[Артикул WB]
)
-- ОСНОВА: справочник (attrs).
-- INNER JOIN wb — только артикулы, у которых есть карточка на WB.
-- LEFT JOIN остальное — подтягиваем данные.
-- WHERE — отсекаем «мёртвые» артикулы (нет ни стока, ни маркетинга за период).
SELECT
    a.[Артикул],
    wb.[Артикул WB],
    a.[Бизнес-группа],
    a.[Розничный отдел],
    a.[Группа],
    a.[Сезон],
    a.[Бренд],
    a.[Ответственный за группу],
    a.[Коллекция],
    ISNULL(m.[Показы], 0)               AS [Показы],
    ISNULL(m.[Клики], 0)                 AS [Клики],
    ISNULL(m.[Заказы], 0)                AS [Заказы],
    CAST(CASE WHEN ISNULL(m.[Показы], 0) = 0 THEN 0
         ELSE ISNULL(m.[Клики], 0) * 100.0 / m.[Показы]
         END AS DECIMAL(9, 2))           AS [CTR, %],
    CAST(CASE WHEN ISNULL(m.[Клики], 0) = 0 THEN 0
         ELSE ISNULL(m.[Заказы], 0) * 100.0 / m.[Клики]
         END AS DECIMAL(9, 2))           AS [Конверсия клики в заказы, %],
    ISNULL(st.[Остаток], 0)              AS [Остаток на {end_label}],
    CAST(CASE WHEN ISNULL(ts.[Всего размеров], 0) = 0 THEN 0
         ELSE ISNULL(ag.[Размеров на агрегаторе], 0) * 100.0 / ts.[Всего размеров]
         END AS DECIMAL(9, 2))           AS [Дистрибуция, %]
FROM attrs a
INNER JOIN wb            ON a.[Артикул] = wb.ITEMID
LEFT JOIN total_sizes ts ON a.[Артикул] = ts.[Артикул]
LEFT JOIN agg_sizes ag   ON a.[Артикул] = ag.[Артикул]
LEFT JOIN stock_end st   ON a.[Артикул] = st.[Артикул]
LEFT JOIN marketing m    ON wb.[Артикул WB] = m.[Артикул WB]
WHERE ISNULL(st.[Остаток], 0) > 0
   OR ISNULL(m.[Показы], 0) > 0
   OR ISNULL(m.[Клики], 0) > 0
   OR ISNULL(m.[Заказы], 0) > 0
ORDER BY a.[Артикул]
"""


def load_base_dataframe(start_date, end_date) -> pd.DataFrame:
    """Выполняет SQL-запрос и возвращает df_base."""
    engine = connect_to_sql(config.SQL_SERVER, config.SQL_DB_PARTNERS)
    # query = build_query(start_date, end_date)           # старый: от стока/маркетинга
    query = build_query_ref_based(start_date, end_date)   # новый: от справочника
    df = pd.read_sql(query, engine)
    return df

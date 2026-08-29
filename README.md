# AutomatedPicking — WB Отбор

<div align="center">

**[English](#-english)  ·  [Русский](#-русский)**

</div>

---

## English

<details open>
<summary><b>Click to expand / collapse the English version</b></summary>

### Overview

Automation of the weekly report **«Отчёт WB о выбросах CTR, CR»** (WB outliers in CTR / CR).
The tool pulls marketing and stock data for Wildberries items from the corporate SQL Server,
enriches them with the photo count and the card creation date from the **WB Content API**,
fills an Excel template with formulas, and sends the result to Telegram.

Everything runs either from a Tkinter GUI or headless from the Windows Task Scheduler.
End users get two `.exe` files and never need Python installed.

### Pipeline

```
SQL (cl01sql) → WB Content API → Excel (template) → Telegram
     [1/4]           [2/4]            [3/4]            [4/4]
```

1. **SQL** — builds the base dataframe: article, WB nmID, business group, department,
   group, season, brand, responsible buyer, collection, impressions / clicks / orders,
   CTR, CR, stock at period end, distribution %. Filters and the period come from
   `settings.json`. Manager full names are joined from `Распределение категорий.xlsx`.
2. **WB Content API** — number of photos per card (`-2` from the actual count, as the
   report expects) and the card creation date. Results are cached in `photo_cache.xlsx`,
   so a re-run with `--use-cache` takes seconds instead of minutes.
3. **Excel** — fills `Отбор Шаблон.xlsx` keeping the template styles, writes formulas for
   CTR / CR and the selection columns («Техничка», «Отбор по CTR / CR», the same
   within-group), and appends an audit sheet **«Применённые фильтры»**.
4. **Telegram** — sends the file to the chat ids from `.env`. Tries a direct connection
   first, then falls back to a public HTTP proxy. This stage is non-fatal: if it fails,
   the file is still on disk.

### Project structure

```
AutomatedPicking/
├── run_gui.py                    # Tkinter GUI (manual run, settings, scheduler)
├── run_otbor.py                  # CLI entry point (used by the scheduler)
├── build_exe.py                  # PyInstaller build for both exe files
├── wb_otbor/
│   ├── config.py                 # Paths, SQL server, defaults, exe lookup, UNC remap
│   ├── settings.py               # settings.json: filters, thresholds, value cache
│   ├── sql_loader.py             # SQL queries + unique filter values
│   ├── wb_api.py                 # WB Content API + photo_cache.xlsx
│   ├── _wb_content.py            # Low-level API client (.env, retries, backoff)
│   ├── excel_writer.py           # Filling the template + «Применённые фильтры» sheet
│   ├── manager_mapping.py        # «Группа → ФИО менеджера» from xlsx
│   ├── telegram_bot_sender.py    # Sending the file (direct or via proxy)
│   ├── scheduler.py              # Windows Task Scheduler wrapper (schtasks.exe)
│   ├── gui_settings.py           # Settings window
│   ├── logging_setup.py          # Rotating log logs/wb_otbor.log
│   └── pipeline.py               # Orchestrator of the 4 stages
├── wb-photo-report/              # Standalone script: photo count by an nmID list
├── Отбор Шаблон.xlsx             # Report template (styles + formulas)
├── Распределение категорий.xlsx  # Group → manager mapping
├── .env                          # WB token + Telegram ids (NOT in git)
├── settings.json                 # User settings (created automatically)
├── photo_cache.xlsx              # Photo cache (created automatically)
├── logs/wb_otbor.log             # Execution log
├── USAGE.md                      # End-user manual (in Russian)
├── DEPLOYMENT.md                 # Deployment on a machine without Python
└── SQL_CATALOG.md                # Reference of the SQL tables in use
```

### Requirements

- Windows 10 / 11
- **ODBC Driver 17 for SQL Server** —
  [download](https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server)
- Access to the Kari corporate network (`cl01sql`), VPN when remote
- Internet access for `content-api.wildberries.ru` and Telegram
- Python 3.10+ **only for development** — end users run the `.exe` files

Python packages:

```bash
pip install pandas sqlalchemy pyodbc openpyxl python-telegram-bot pyinstaller
```

### Configuration

Create a `.env` file **next to the exe files** (the project root):

```
WB_API_TOKEN = <WB Seller token, Content category>
TALDYKIN_ID  = <telegram chat id>
ANALYTICS_AUTO2 = <telegram chat id>
```

`config.py` also looks in `wb-photo-report/.env` and one level up, for backward
compatibility. `.env` is in `.gitignore` and must never be committed.

### Usage

GUI:

```bash
python run_gui.py
```

CLI (this is what the scheduler runs):

```bash
python run_otbor.py
python run_otbor.py --use-cache
python run_otbor.py --output-dir D:\reports --log-file run.log
```

Building the executables:

```bash
python build_exe.py --clean
```

Produces `wb_otbor_runner.exe` (console pipeline) and `wb_otbor_gui.exe` (GUI); both are
copied to the project root. `config.get_runner_command()` automatically switches the
scheduler from `python run_otbor.py` to the exe as soon as it is built.

As a library:

```python
from wb_otbor.pipeline import run_full_pipeline

path = run_full_pipeline(use_photo_cache=True)
```

### Settings (`settings.json`)

Seven multi-select reference filters — Бизнес-группа, Розничный отдел, Группа, Сезон,
Бренд, Ответственный за группу, Коллекция. **An empty list means the filter is not
applied.** Values are pulled from the reference with the *«Получить уникальные значения
из БД»* button in the settings window and cached in the same file.

| Parameter | Default | Meaning |
| --- | ---: | --- |
| `period_days` | 7 | Report period length in days |
| `offset_from_today` | 1 | Period end = today − N (WB data arrives with a lag) |
| `min_stock` | 10 | Below → the item is marked «Мало остатка» |
| `min_distrib_percent` | 20 | Below → also «Мало остатка» |
| `min_shows` | 1200 | Below → the item is marked «Мало показов» |

### Scheduling

The GUI wraps `schtasks.exe` and manages a task named **`WB_Otbor_AutoRun`**: pick weekdays
and a time, then create / enable / disable / delete it. Creating it manually:

```bash
schtasks /Create /TN WB_Otbor_AutoRun /TR "C:\path\to\wb_otbor_runner.exe" /SC WEEKLY /D MON /ST 09:00 /F
```

> **UNC note.** Windows canonicalises the DFS path `\\kari.local\public\...` into
> `\\fs05\all\...`, which the scheduler account may not have access to.
> `config.SCHEDULER_UNC_REMAP` maps it back before calling `schtasks`, and paths are
> deliberately built with `.absolute()` instead of `.resolve()`.

### Output

`Отчёт WB о выбросах CTR, CR на DD.MM.xlsx`, two sheets:

- **Sheet 1** — 26 columns per item, sorted by impressions descending: reference attributes,
  marketing metrics, stock and distribution, «Техничка» (1 = passes the thresholds, 0 = filtered
  out), the selection verdicts overall and within the group, photo count and the WB card
  creation date.
- **Sheet 2 «Применённые фильтры»** — the settings the report was built with, for audit and
  reproducibility.

### `wb-photo-report/`

A standalone script, independent from the pipeline: it takes a list of `nmID` (from stdin or
`--input-file`) and writes an Excel file with two sheets — `Найдено` and `Не найдено`.
It queries each nmID individually via `textSearch` in `cards/list` with `withPhoto: -1`,
which is slower but far more reliable on large catalogues than paging the whole account.

### Troubleshooting

Everything is written to `logs/wb_otbor.log` (rotating). The most common cases:

| Symptom | Cause |
| --- | --- |
| «Ошибка SQL» / no connection to `cl01sql` | Outside the Kari network or VPN is down |
| «Не найден токен WB API» | `.env` is missing next to the exe or has no `WB_API_TOKEN` |
| Filter lists are empty | The reference cache has not been fetched yet |
| The file was created but never arrived in Telegram | Proxy could not get through — send it manually |
| The scheduled task did not fire | Check `taskschd.msc` → `WB_Otbor_AutoRun` → History; re-save the task if the path starts with `\\fs05\` |

See `USAGE.md` for the full end-user manual and `DEPLOYMENT.md` for deploying to a machine
without Python.

</details>

---

## Русский

<details open>
<summary><b>Нажмите, чтобы развернуть / свернуть русскую версию</b></summary>

### Описание

Автоматизация еженедельного отчёта **«Отчёт WB о выбросах CTR, CR»**.
Программа подтягивает маркетинг и остатки по товарам Wildberries из корпоративного
SQL Server, дополняет их количеством фото и датой создания карточки из **WB Content API**,
заполняет Excel-шаблон с формулами и отправляет результат в Telegram.

Работает как из GUI на Tkinter, так и без окон — из Планировщика задач Windows.
Конечному пользователю выдаются два `.exe`, Python устанавливать не нужно.

### Пайплайн

```
SQL (cl01sql) → WB Content API → Excel (шаблон) → Telegram
     [1/4]           [2/4]            [3/4]           [4/4]
```

1. **SQL** — сборка базового датафрейма: артикул, артикул WB (nmID), бизнес-группа,
   отдел, группа, сезон, бренд, ответственный, коллекция, показы / клики / заказы,
   CTR, конверсия, остаток на конец периода, дистрибуция %. Фильтры и период берутся
   из `settings.json`. ФИО менеджера подтягивается из `Распределение категорий.xlsx`.
2. **WB Content API** — количество фото на карточке (в отчёт пишется `-2` от фактического)
   и дата создания карточки. Результат кэшируется в `photo_cache.xlsx`, поэтому повторный
   прогон с `--use-cache` занимает секунды вместо минут.
3. **Excel** — заполнение `Отбор Шаблон.xlsx` с сохранением стилей шаблона, формулы CTR / CR
   и колонок отбора («Техничка», «Отбор по CTR / CR», то же внутри группы), плюс
   аудит-лист **«Применённые фильтры»**.
4. **Telegram** — отправка файла в чаты из `.env`. Сначала пробуется прямое соединение,
   затем — публичный HTTP-прокси. Стадия некритичная: если не удалось, файл всё равно
   лежит на диске.

### Структура проекта

```
AutomatedPicking/
├── run_gui.py                    # GUI на Tkinter (ручной запуск, настройки, планировщик)
├── run_otbor.py                  # CLI-точка входа (её запускает планировщик)
├── build_exe.py                  # Сборка обоих exe через PyInstaller
├── wb_otbor/
│   ├── config.py                 # Пути, SQL-сервер, дефолты, поиск exe, UNC-ремап
│   ├── settings.py               # settings.json: фильтры, пороги, кэш значений
│   ├── sql_loader.py             # SQL-запросы + уникальные значения фильтров
│   ├── wb_api.py                 # WB Content API + photo_cache.xlsx
│   ├── _wb_content.py            # Низкоуровневый клиент API (.env, ретраи, backoff)
│   ├── excel_writer.py           # Заполнение шаблона + лист «Применённые фильтры»
│   ├── manager_mapping.py        # «Группа → ФИО менеджера» из xlsx
│   ├── telegram_bot_sender.py    # Отправка файла (напрямую или через прокси)
│   ├── scheduler.py              # Обёртка над Планировщиком задач (schtasks.exe)
│   ├── gui_settings.py           # Окно настроек
│   ├── logging_setup.py          # Ротируемый лог logs/wb_otbor.log
│   └── pipeline.py               # Оркестратор четырёх стадий
├── wb-photo-report/              # Отдельный скрипт: количество фото по списку nmID
├── Отбор Шаблон.xlsx             # Шаблон отчёта (стили + формулы)
├── Распределение категорий.xlsx  # Маппинг группа → менеджер
├── .env                          # Токен WB + Telegram-чаты (НЕ в git)
├── settings.json                 # Пользовательские настройки (создаётся сам)
├── photo_cache.xlsx              # Кэш фото (создаётся сам)
├── logs/wb_otbor.log             # Журнал выполнения
├── USAGE.md                      # Инструкция для пользователя
├── DEPLOYMENT.md                 # Развёртывание на ПК без Python
└── SQL_CATALOG.md                # Справочник используемых SQL-таблиц
```

### Требования

- Windows 10 / 11
- **ODBC Driver 17 for SQL Server** —
  [скачать](https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server)
- Доступ в корпоративную сеть Kari (`cl01sql`), удалённо — через VPN
- Интернет для `content-api.wildberries.ru` и Telegram
- Python 3.10+ **только для разработки** — пользователю достаточно `.exe`

Python-пакеты:

```bash
pip install pandas sqlalchemy pyodbc openpyxl python-telegram-bot pyinstaller
```

### Настройка

Создайте файл `.env` **рядом с exe-файлами** (в корне проекта):

```
WB_API_TOKEN = <токен WB Seller, категория Content>
TALDYKIN_ID  = <telegram chat id>
ANALYTICS_AUTO2 = <telegram chat id>
```

Для обратной совместимости `config.py` также ищет `.env` в `wb-photo-report/` и на уровень
выше. Файл перечислен в `.gitignore` и не должен попадать в репозиторий.

### Запуск

GUI:

```bash
python run_gui.py
```

CLI (именно его запускает планировщик):

```bash
python run_otbor.py
python run_otbor.py --use-cache
python run_otbor.py --output-dir D:\reports --log-file run.log
```

Сборка исполняемых файлов:

```bash
python build_exe.py --clean
```

Получаются `wb_otbor_runner.exe` (консольный pipeline) и `wb_otbor_gui.exe` (GUI); оба
копируются в корень проекта. `config.get_runner_command()` автоматически переключает
планировщик с `python run_otbor.py` на exe, как только тот собран.

Как библиотека:

```python
from wb_otbor.pipeline import run_full_pipeline

path = run_full_pipeline(use_photo_cache=True)
```

### Настройки (`settings.json`)

Семь multi-select фильтров справочника — Бизнес-группа, Розничный отдел, Группа, Сезон,
Бренд, Ответственный за группу, Коллекция. **Пустой список = фильтр не применяется.**
Значения тянутся из справочника кнопкой *«Получить уникальные значения из БД»* в окне
настроек и кэшируются в том же файле.

| Параметр | По умолчанию | Что делает |
| --- | ---: | --- |
| `period_days` | 7 | Длина периода отчёта, дней |
| `offset_from_today` | 1 | Конец периода = сегодня − N (данные WB приходят с задержкой) |
| `min_stock` | 10 | Ниже → метка «Мало остатка» |
| `min_distrib_percent` | 20 | Ниже → тоже «Мало остатка» |
| `min_shows` | 1200 | Ниже → метка «Мало показов» |

### Планировщик

GUI — обёртка над `schtasks.exe`, управляет задачей **`WB_Otbor_AutoRun`**: выбор дней
недели и времени, создание / включение / отключение / удаление. Вручную:

```bash
schtasks /Create /TN WB_Otbor_AutoRun /TR "C:\path\to\wb_otbor_runner.exe" /SC WEEKLY /D MON /ST 09:00 /F
```

> **Про UNC.** Windows канонизирует DFS-путь `\\kari.local\public\...` в `\\fs05\all\...`,
> к которому у учётной записи планировщика может не быть доступа.
> `config.SCHEDULER_UNC_REMAP` подменяет префикс обратно перед вызовом `schtasks`,
> а пути специально строятся через `.absolute()`, а не `.resolve()`.

### Результат

`Отчёт WB о выбросах CTR, CR на DD.MM.xlsx`, два листа:

- **Лист 1** — 26 колонок по артикулам, отсортировано по показам по убыванию: атрибуты
  справочника, маркетинговые метрики, остаток и дистрибуция, «Техничка» (1 = проходит по
  порогам, 0 = отсеян), вердикты отбора в целом и внутри группы, количество фото и дата
  создания карточки на WB.
- **Лист 2 «Применённые фильтры»** — на каких настройках собран отчёт, для аудита и
  воспроизводимости.

### `wb-photo-report/`

Отдельный скрипт, не связанный с пайплайном: принимает список `nmID` (со stdin или через
`--input-file`) и пишет Excel с двумя листами — `Найдено` и `Не найдено`.
Запрашивает каждый nmID адресно через `textSearch` в `cards/list` с `withPhoto: -1` —
медленнее, но заметно надёжнее полного обхода кабинета на больших каталогах.

### Диагностика

Всё пишется в `logs/wb_otbor.log` (с ротацией). Частые случаи:

| Симптом | Причина |
| --- | --- |
| «Ошибка SQL» / нет подключения к `cl01sql` | Вне сети Kari или не поднят VPN |
| «Не найден токен WB API» | Нет `.env` рядом с exe или в нём нет `WB_API_TOKEN` |
| Списки фильтров пустые | Справочник ещё не выгружен из БД |
| Файл создан, но не пришёл в Telegram | Прокси не пробился — отправьте вручную |
| Задача в планировщике не сработала | `taskschd.msc` → `WB_Otbor_AutoRun` → «Журнал»; если путь начинается с `\\fs05\` — пересохраните задачу из GUI |

Полная инструкция для пользователя — в `USAGE.md`, развёртывание на ПК без Python —
в `DEPLOYMENT.md`.

</details>

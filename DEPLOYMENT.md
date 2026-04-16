# Развёртывание на другом ПК

## Разработка (свой ПК с Python)

```
AutomatedPicking/
├── wb_otbor/                  ← пакет с логикой
├── run_gui.py                 ← GUI
├── run_otbor.py               ← CLI
├── build_exe.py               ← сборка exe
├── Отбор 13.04.xlsx           ← шаблон
├── wb-photo-report/.env       ← токен WB API
└── photo_cache.xlsx           ← создаётся автоматически
```

Запуск:
```
python run_gui.py       # GUI
python run_otbor.py     # CLI (без кэша, как для планировщика)
```

## Сборка EXE

Через GUI: кнопка **"🔨 Собрать EXE"** (в секции «Сборка EXE»).

Или вручную:
```
python build_exe.py --clean
```

Результат: `dist/wb_otbor_runner.exe` (~80–120 МБ).

После сборки GUI и планировщик **автоматически** начинают использовать exe
(функция `config.find_runner_exe()` ищет его рядом с проектом и в `dist/`).

## Развёртывание на целевом ПК (без Python)

Скопируйте на целевой ПК:
```
AutomatedPicking/
├── wb_otbor_runner.exe        ← из dist/
├── Отбор 13.04.xlsx
└── wb-photo-report/
    └── .env
```

### Системные требования

- **Windows 10/11**
- **ODBC Driver 17 for SQL Server** — поставить отдельно
  (https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server)
- Сетевой доступ к `cl01sql` и `content-api.wildberries.ru`

### Проверка запуска

```cmd
cd AutomatedPicking
wb_otbor_runner.exe
```

Должен выполниться полный цикл и создаться файл `Отбор DD.MM.xlsx`.

### Планировщик на целевом ПК

Без GUI (работает только через CLI, потому что GUI на целевом ПК не нужен —
задачу создаём там, где удобно), но **scheduler.py** умеет работать из exe:

Вариант 1 — создать задачу через `schtasks.exe` напрямую:
```cmd
schtasks /Create /TN WB_Otbor_AutoRun /TR "C:\path\to\wb_otbor_runner.exe" /SC WEEKLY /D MON,WED,FRI /ST 09:00 /F
```

Вариант 2 — скопировать также `run_gui.py` + `wb_otbor/` на целевой ПК
(GUI — просто интерфейс к `schtasks`, сам не запускает pipeline через exe).
Python всё равно не нужен для запуска задач — только для запуска GUI.

## Файлы, которые нельзя забыть

| Файл | Для чего |
|------|----------|
| `wb_otbor_runner.exe` | Сам pipeline |
| `Отбор 13.04.xlsx` | Шаблон для заполнения |
| `wb-photo-report/.env` | Токен WB API (`WB_API_TOKEN=...`) |

## Как работает переключение exe/script

`wb_otbor/config.py::get_runner_command()`:

1. Если рядом или в `dist/` есть `wb_otbor_runner.exe` → возвращает `[exe_path]`
2. Иначе → возвращает `[python.exe, run_otbor.py]`

Функция `scheduler.build_tr_command()` собирает из этого строку для `/TR`.
Планировщик автоматически переключается на exe, как только его собрали —
пересоздавать задачу не нужно, но имеет смысл (иначе в задаче останется
путь к python+script с предыдущего сохранения).

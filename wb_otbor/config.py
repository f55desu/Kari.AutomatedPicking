"""Конфигурация путей и параметров."""
import sys
from pathlib import Path


def _resolve_base_dir() -> Path:
    """
    Возвращает корень проекта.
      - В обычном режиме (не frozen): родитель пакета wb_otbor (то есть AutomatedPicking/).
      - В frozen-режиме (PyInstaller onefile/onedir): директория, где лежит exe.
        Это позволяет класть Отбор date.xlsx, wb-photo-report/.env и photo_cache.xlsx
        РЯДОМ с exe, а не внутри временной _MEIPASS.

    Используем .absolute() вместо .resolve(), чтобы СОХРАНИТЬ исходный UNC-путь
    (\\\\kari.local\\public\\...). .resolve() канонизирует UNC к форме \\\\fs05\\all\\...,
    и планировщик задач затем прописывает такой путь, к которому у его учётки
    доступа нет.
    """
    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).absolute().parent
        # Если exe лежит в dist/ — поднимаемся на уровень выше к корню проекта
        if exe_dir.name.lower() == 'dist' and (exe_dir.parent / 'wb-photo-report').exists():
            return exe_dir.parent
        return exe_dir
    # обычный запуск python -m wb_otbor / python run_otbor.py
    return Path(__file__).absolute().parent.parent


BASE_DIR = _resolve_base_dir()


def _find_env_file() -> Path:
    """
    Ищет .env в нескольких местах, чтобы работать и из корня, и из dist/.
    """
    candidates = [
        BASE_DIR / 'wb-photo-report' / '.env',
        BASE_DIR / '.env',
        BASE_DIR.parent / 'wb-photo-report' / '.env',  # если exe в dist/
    ]
    for p in candidates:
        if p.exists():
            return p
    # Возвращаем стандартный (для внятного сообщения об ошибке)
    return candidates[0]


# Файлы
TEMPLATE_FILE = BASE_DIR / 'Отбор 11.04.xlsx'
PHOTO_CACHE_FILE = BASE_DIR / 'photo_cache.xlsx'
ENV_FILE = _find_env_file()

# SQL
SQL_SERVER = 'cl01sql'
SQL_DB_PARTNERS = 'DBPartners'
SQL_DB_REPORT = 'DBReport'

# Параметры отчёта
# BUSINESS_GROUPS = ('Одежда для детей', 'Одежда и аксессуары')
COLLECTION = '2026SS'
PERIOD_DAYS = 7           # длина периода
OFFSET_FROM_TODAY = 1     # конец периода = сегодня - N дней

# Пороги для Технички и Отбора
MIN_STOCK = 10
MIN_DISTRIB_PERCENT = 20
MIN_SHOWS = 1200

# Планировщик задач
TASK_NAME = "WB_Otbor_AutoRun"

# UNC-ремап для путей, которые прописываются в /TR у schtasks.
# Windows при открытии DFS-путей \\kari.local\public\... выдаёт Python
# каноническое имя \\fs05\all\..., и планировщик сохраняет путь в такой форме.
# К серверу \\fs05 у учётной записи планировщика может не быть доступа,
# поэтому перед вызовом schtasks подменяем префикс обратно на DFS.
#   ключ = канонический префикс (как даёт Path)
#   значение = пользовательский префикс (DFS / то, что реально доступно)
# Сравнение case-insensitive.
SCHEDULER_UNC_REMAP = {
    r"\\fs05\all": r"\\kari.local\public\all",
}

# Имена exe-файлов после сборки через build_exe.py
RUNNER_EXE_NAME = "wb_otbor_runner.exe"
GUI_EXE_NAME = "wb_otbor_gui.exe"

# Относительный путь к CLI-скрипту (для запуска через python)
RUNNER_SCRIPT_NAME = "run_otbor.py"

F55_ID = "421762273"
TALDYKIN_ID = "6917999439"
ANALYTICS_AUTO = "-5183358607"

def find_runner_exe() -> Path | None:
    """
    Ищет скомпилированный exe в типовых местах относительно BASE_DIR.
    Возвращает Path как есть, БЕЗ .resolve() — чтобы сохранить исходный
    UNC-путь (например, \\\\kari.local\\public\\...). Иначе .resolve()
    канонизирует UNC к форме \\\\fs05\\all\\..., которая может быть
    недоступна учётной записи планировщика задач.
    """
    candidates = [
        BASE_DIR / RUNNER_EXE_NAME,                # рядом с проектом (после копирования из dist)
        BASE_DIR / 'dist' / RUNNER_EXE_NAME,       # стандартный выход PyInstaller
    ]
    for p in candidates:
        if p.exists():
            return p
    return None


def get_runner_command(use_cache: bool = False) -> list[str]:
    """
    Возвращает команду для запуска сборки — либо exe, либо python+скрипт.
    Используется scheduler'ом и GUI при формировании задачи планировщика.
    """
    exe = find_runner_exe()
    if exe:
        cmd = [str(exe)]
    else:
        script = BASE_DIR / RUNNER_SCRIPT_NAME
        cmd = [sys.executable, str(script)]
    if use_cache:
        cmd.append('--use-cache')
    return cmd

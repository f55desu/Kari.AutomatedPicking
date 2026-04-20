"""
Сборка exe-файлов через PyInstaller.

Использование:
    python build_exe.py                  # собрать ОБА exe
    python build_exe.py --target runner  # только консольный pipeline
    python build_exe.py --target gui     # только GUI
    python build_exe.py --clean          # с очисткой кэша pyinstaller

Что получается:
    wb_otbor_runner.exe  — консольный pipeline (для планировщика задач)
    wb_otbor_gui.exe     — GUI с кнопками и планировщиком (для пользователя)

Оба exe копируются в корень проекта автоматически.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent

# Общие скрытые импорты для обоих exe
COMMON_HIDDEN_IMPORTS = [
    # SQLAlchemy + pyodbc
    'pyodbc',
    'sqlalchemy.dialects.mssql',
    'sqlalchemy.dialects.mssql.pyodbc',
    'sqlalchemy.sql.default_comparator',
    # pandas / openpyxl
    'pandas',
    'openpyxl',
    'openpyxl.cell._writer',
    # наш пакет
    'wb_otbor',
    'wb_otbor.config',
    'wb_otbor.sql_loader',
    'wb_otbor.wb_api',
    'wb_otbor._wb_content',
    'wb_otbor.excel_writer',
    'wb_otbor.pipeline',
    'wb_otbor.scheduler',
    'wb_otbor.settings',
    'wb_otbor.logging_setup',
    'wb_otbor.telegram_bot_sender',
]

# Конфигурация каждого exe
TARGETS = {
    'runner': {
        'entry': PROJECT_ROOT / 'run_otbor.py',
        'name': 'wb_otbor_runner',
        'console': True,
        'extra_hidden': [],
    },
    'gui': {
        'entry': PROJECT_ROOT / 'run_gui.py',
        'name': 'wb_otbor_gui',
        'console': False,  # --windowed: без чёрного окна консоли
        'extra_hidden': [
            'telegram',
            'telegram.request',
            'httpx',
            'httpcore',
            'wb_otbor.gui_settings',  # окно настроек импортируется лениво
        ],
    },
}


def check_pyinstaller():
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller не установлен. Ставим его:")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])


def clean_build():
    for folder in ('build', 'dist', '__pycache__'):
        p = PROJECT_ROOT / folder
        if p.exists():
            print(f"Удаляю {p}...")
            shutil.rmtree(p, ignore_errors=True)
    for spec in PROJECT_ROOT.glob('*.spec'):
        print(f"Удаляю {spec}...")
        spec.unlink()


def build_one(target_key: str):
    cfg = TARGETS[target_key]
    entry = cfg['entry']
    exe_name = cfg['name']
    is_console = cfg['console']

    print()
    print("=" * 60)
    print(f"Сборка: {exe_name}.exe  ({'console' if is_console else 'windowed'})")
    print(f"Точка входа: {entry}")
    print("=" * 60)

    all_hidden = COMMON_HIDDEN_IMPORTS + cfg['extra_hidden']

    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',
        '--console' if is_console else '--windowed',
        '--name', exe_name,
        '--distpath', str(PROJECT_ROOT / 'dist'),
        '--workpath', str(PROJECT_ROOT / 'build' / exe_name),
        '--specpath', str(PROJECT_ROOT),
    ]
    for hi in all_hidden:
        cmd += ['--hidden-import', hi]
    cmd.append(str(entry))

    subprocess.check_call(cmd, cwd=str(PROJECT_ROOT))

    dist_exe = PROJECT_ROOT / 'dist' / f'{exe_name}.exe'
    root_exe = PROJECT_ROOT / f'{exe_name}.exe'

    if not dist_exe.exists():
        print(f"ОШИБКА: {dist_exe} не найден после сборки.")
        sys.exit(1)

    shutil.copy2(dist_exe, root_exe)
    size_mb = root_exe.stat().st_size / (1024 * 1024)
    print(f"OK: {root_exe}  ({size_mb:.1f} МБ)")
    return root_exe


def main():
    parser = argparse.ArgumentParser(description="Сборка exe-файлов проекта WB Отбор")
    parser.add_argument('--target', choices=['runner', 'gui', 'all'], default='all',
                        help='Что собирать: runner (консоль), gui (окно), all (оба). По умолчанию: all')
    parser.add_argument('--clean', action='store_true',
                        help='Удалить build/dist/spec перед сборкой')
    args = parser.parse_args()

    check_pyinstaller()

    if args.clean:
        clean_build()

    targets = list(TARGETS.keys()) if args.target == 'all' else [args.target]
    results = []
    for t in targets:
        exe_path = build_one(t)
        results.append(exe_path)

    print()
    print("=" * 60)
    print("ГОТОВО. Собранные exe (скопированы в корень проекта):")
    for p in results:
        print(f"  {p.name}  ({p.stat().st_size / 1024 / 1024:.1f} МБ)")
    print("=" * 60)
    print()
    print("Для запуска на другом ПК скопируйте в одну папку:")
    print("  1. wb_otbor_runner.exe  (для планировщика / CLI)")
    print("  2. wb_otbor_gui.exe     (для пользователя)")
    print("  3. Отбор *.xlsx         (шаблон)")
    print("  4. wb-photo-report/.env (токен WB API)")
    print()
    print("На целевом ПК требуется: ODBC Driver 17 for SQL Server.")


if __name__ == '__main__':
    main()

"""
Сборка wb_otbor_runner.exe через PyInstaller.

Использование:
    python build_exe.py              # обычная сборка --onefile
    python build_exe.py --clean      # с очисткой кэша pyinstaller

Требования:
    pip install pyinstaller

Что получается:
    dist/wb_otbor_runner.exe         — автономный консольный exe

Что нужно на целевом ПК (рядом с exe):
    Отбор 13.04.xlsx                 — шаблон
    wb-photo-report/.env             — токен WB API
    photo_cache.xlsx                 — создаётся автоматически при первом запуске

Системные требования на целевом ПК:
    ODBC Driver 17 for SQL Server    — установить отдельно
    Сетевой доступ к cl01sql и content-api.wildberries.ru
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
ENTRY_SCRIPT = PROJECT_ROOT / 'run_otbor.py'
EXE_NAME = 'wb_otbor_runner'

# Скрытые импорты, которые PyInstaller сам может не найти
HIDDEN_IMPORTS = [
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
]


def check_pyinstaller():
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller не установлен. Ставим его:")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])


def build(clean: bool = False):
    check_pyinstaller()

    if clean:
        for folder in ('build', 'dist', '__pycache__'):
            p = PROJECT_ROOT / folder
            if p.exists():
                print(f"Удаляю {p}...")
                shutil.rmtree(p, ignore_errors=True)
        spec = PROJECT_ROOT / f'{EXE_NAME}.spec'
        if spec.exists():
            spec.unlink()

    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',
        '--console',
        '--name', EXE_NAME,
        '--distpath', str(PROJECT_ROOT / 'dist'),
        '--workpath', str(PROJECT_ROOT / 'build'),
        '--specpath', str(PROJECT_ROOT),
    ]
    for hi in HIDDEN_IMPORTS:
        cmd += ['--hidden-import', hi]
    cmd.append(str(ENTRY_SCRIPT))

    print("=" * 60)
    print("Команда PyInstaller:")
    print(" ".join(cmd))
    print("=" * 60)

    subprocess.check_call(cmd, cwd=str(PROJECT_ROOT))

    dist_exe = PROJECT_ROOT / 'dist' / f'{EXE_NAME}.exe'
    root_exe = PROJECT_ROOT / f'{EXE_NAME}.exe'

    if not dist_exe.exists():
        print(f"ОШИБКА: {dist_exe} не найден после сборки.")
        sys.exit(1)

    # Копируем exe в корень проекта — рядом с .env и шаблоном,
    # чтобы config.BASE_DIR (= директория exe) сразу видел все файлы.
    shutil.copy2(dist_exe, root_exe)

    size_mb = root_exe.stat().st_size / (1024 * 1024)
    print()
    print("=" * 60)
    print(f"ГОТОВО. {root_exe}  ({size_mb:.1f} МБ)")
    print("=" * 60)
    print()
    print("EXE автоматически скопирован в корень проекта:")
    print(f"  {root_exe}")
    print()
    print("Для запуска на другом ПК скопируйте в одну папку:")
    print(f"  1. {root_exe.name}")
    print(f"  2. Отбор *.xlsx  (шаблон)")
    print(f"  3. wb-photo-report/.env  (папку целиком)")
    print()
    print("На целевом ПК требуется: ODBC Driver 17 for SQL Server.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--clean', action='store_true',
                        help='Удалить build/dist/spec перед сборкой')
    args = parser.parse_args()
    build(clean=args.clean)


if __name__ == '__main__':
    main()

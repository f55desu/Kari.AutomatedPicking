"""
CLI-точка входа. Вызывается планировщиком задач и из GUI.

Примеры:
    python run_otbor.py                 # полный цикл с запросом к API
    python run_otbor.py --use-cache     # использовать кэш photo_cache.xlsx
    python run_otbor.py --log-file out.log
"""
from __future__ import annotations

import argparse
import sys
import traceback
from datetime import datetime
from pathlib import Path

# Добавляем путь к пакету wb_otbor, если запускаем не из своей директории
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from wb_otbor.pipeline import run_full_pipeline


def _make_logger(log_file: Path | None):
    """Возвращает функцию log, которая пишет в stdout + опционально в файл."""
    fh = None
    if log_file:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        fh = open(log_file, 'a', encoding='utf-8')

    def log(msg: str):
        stamp = datetime.now().strftime('%H:%M:%S')
        line = f"[{stamp}] {msg}"
        print(line, flush=True)
        if fh:
            fh.write(line + '\n')
            fh.flush()

    return log, fh


def main() -> int:
    parser = argparse.ArgumentParser(description="Сборка файла Отбор WB")
    parser.add_argument('--use-cache', action='store_true',
                        help='Читать количество фото из photo_cache.xlsx вместо WB API.')
    parser.add_argument('--output-dir', type=Path, default=None,
                        help='Директория для итогового файла (по умолчанию — корень проекта).')
    parser.add_argument('--log-file', type=Path, default=None,
                        help='Путь к файлу лога. По умолчанию только stdout.')
    args = parser.parse_args()

    log, fh = _make_logger(args.log_file)
    try:
        log(f"Старт. use_cache={args.use_cache}, output_dir={args.output_dir}")
        output_path = run_full_pipeline(
            use_photo_cache=args.use_cache,
            output_dir=args.output_dir,
            log=log,
        )
        log(f"УСПЕХ: {output_path}")
        return 0
    except Exception as exc:
        log(f"ОШИБКА: {exc}")
        log(traceback.format_exc())
        return 1
    finally:
        if fh:
            fh.close()


if __name__ == '__main__':
    raise SystemExit(main())

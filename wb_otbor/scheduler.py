"""
Windows Task Scheduler wrapper.

Использует встроенный schtasks.exe (есть на любой Windows), поэтому не требует
pywin32. Операция всегда идемпотентна: если задача существует — обновляем её,
если нет — создаём. Удаление задачи — отдельный вызов.
"""
from __future__ import annotations

import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from . import config


# schtasks принимает сокращения дней недели через запятую
DAYS_FULL_TO_SHORT = {
    'Monday': 'MON', 'Tuesday': 'TUE', 'Wednesday': 'WED',
    'Thursday': 'THU', 'Friday': 'FRI', 'Saturday': 'SAT', 'Sunday': 'SUN',
}
DAYS_RU = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']
DAYS_SHORT_ORDER = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN']


@dataclass
class TaskInfo:
    exists: bool
    enabled: bool = False
    schedule: str = ''
    last_run: str = ''
    next_run: str = ''
    raw: str = ''


def _run_schtasks(args: list, check: bool = False) -> subprocess.CompletedProcess:
    """Запускает schtasks, возвращает CompletedProcess."""
    # Для русского вывода schtasks
    encodings = ['cp866', 'cp1251', 'utf-8']
    result = subprocess.run(
        ['schtasks'] + args,
        capture_output=True,
        check=False,
    )
    # Пытаемся декодировать stdout и stderr
    for enc in encodings:
        try:
            stdout = result.stdout.decode(enc)
            stderr = result.stderr.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        stdout = result.stdout.decode('utf-8', errors='replace')
        stderr = result.stderr.decode('utf-8', errors='replace')

    # Собираем обратно объект с декодированными строками
    class _Result:
        pass
    r = _Result()
    r.returncode = result.returncode
    r.stdout = stdout
    r.stderr = stderr
    r.args = result.args

    if check and result.returncode != 0:
        raise RuntimeError(f"schtasks failed: {stderr or stdout}")
    return r


def task_exists(task_name: str = None) -> bool:
    name = task_name or config.TASK_NAME
    r = _run_schtasks(['/Query', '/TN', name])
    return r.returncode == 0


def get_task_info(task_name: str = None) -> TaskInfo:
    name = task_name or config.TASK_NAME
    r = _run_schtasks(['/Query', '/TN', name, '/FO', 'LIST', '/V'])
    if r.returncode != 0:
        return TaskInfo(exists=False)

    info = TaskInfo(exists=True, raw=r.stdout)
    for raw_line in r.stdout.splitlines():
        line = raw_line.strip()
        low = line.lower()
        # Статус задачи
        if low.startswith('scheduled task state') or low.startswith('состояние запланированной'):
            val = line.split(':', 1)[-1].strip().lower()
            info.enabled = 'enabled' in val or 'включ' in val
        elif low.startswith('schedule type') or low.startswith('тип расписания'):
            info.schedule = line.split(':', 1)[-1].strip()
        elif low.startswith('start time') or low.startswith('время начала'):
            if not info.schedule.endswith(line.split(':', 1)[-1].strip()):
                info.schedule += ' @ ' + line.split(':', 1)[-1].strip()
        elif low.startswith('days') or low.startswith('дни'):
            info.schedule += ' (' + line.split(':', 1)[-1].strip() + ')'
        elif low.startswith('last run time') or low.startswith('время последнего'):
            info.last_run = line.split(':', 1)[-1].strip()
        elif low.startswith('next run time') or low.startswith('следующее время'):
            info.next_run = line.split(':', 1)[-1].strip()
    return info


def delete_task(task_name: str = None) -> None:
    name = task_name or config.TASK_NAME
    if task_exists(name):
        _run_schtasks(['/Delete', '/TN', name, '/F'], check=True)


def enable_task(enabled: bool, task_name: str = None) -> None:
    name = task_name or config.TASK_NAME
    flag = '/ENABLE' if enabled else '/DISABLE'
    _run_schtasks(['/Change', '/TN', name, flag], check=True)


def days_from_weekday_flags(flags: Iterable[bool]) -> str:
    """
    Из списка из 7 булевых (Пн..Вс) возвращает строку 'MON,WED,FRI'.
    """
    flags = list(flags)
    if len(flags) != 7:
        raise ValueError("flags должен содержать ровно 7 элементов (Пн–Вс).")
    picked = [DAYS_SHORT_ORDER[i] for i, on in enumerate(flags) if on]
    if not picked:
        raise ValueError("Выберите хотя бы один день недели.")
    return ','.join(picked)


def _quote_for_tr(arg: str) -> str:
    """
    Квотим ВСЁ, что не выглядит как флаг (--foo / -x).
    Schtasks /TR парсит строку, пути с пробелами обязательно должны быть в кавычках.
    """
    if arg.startswith('-'):
        return arg
    if arg.startswith('"') and arg.endswith('"'):
        return arg
    return f'"{arg}"'


def build_tr_command(use_cache: bool = False) -> str:
    """
    Собирает строку для /TR. Использует config.get_runner_command, который
    автоматически подбирает exe или python+скрипт в зависимости от того,
    что доступно.
    """
    cmd = config.get_runner_command(use_cache=use_cache)
    return ' '.join(_quote_for_tr(p) for p in cmd)


def create_or_update_task(days_flags: Iterable[bool],
                          time_str: str,
                          task_name: str = None,
                          use_cache: bool = False) -> None:
    """
    Создаёт или обновляет задачу. Идемпотентно: если существует — удаляем и создаём заново.

    days_flags : 7 булевых — дни недели Пн..Вс
    time_str   : 'HH:MM'
    use_cache  : флаг --use-cache для целевого runner'а
    """
    name = task_name or config.TASK_NAME

    # Проверяем время
    if len(time_str) != 5 or time_str[2] != ':':
        raise ValueError(f"time_str должен быть 'HH:MM', получено: {time_str!r}")
    hh, mm = time_str.split(':')
    if not (hh.isdigit() and mm.isdigit()):
        raise ValueError("time_str должен быть 'HH:MM'")

    days_str = days_from_weekday_flags(days_flags)

    # Автовыбор: exe, если есть; иначе python + run_otbor.py
    tr_cmd = build_tr_command(use_cache=use_cache)

    # Удаляем старую, если есть
    if task_exists(name):
        _run_schtasks(['/Delete', '/TN', name, '/F'], check=True)

    # Создаём новую
    args = [
        '/Create',
        '/TN', name,
        '/TR', tr_cmd,
        '/SC', 'WEEKLY',
        '/D', days_str,
        '/ST', time_str,
        '/RL', 'LIMITED',
        '/F',
    ]
    _run_schtasks(args, check=True)


def run_task_now(task_name: str = None) -> None:
    """Запускает задачу немедленно через планировщик."""
    name = task_name or config.TASK_NAME
    _run_schtasks(['/Run', '/TN', name], check=True)

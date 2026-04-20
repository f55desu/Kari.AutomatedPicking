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
from .logging_setup import get_logger


logger = get_logger('scheduler')


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


def _classify_enabled(value: str) -> bool | None:
    """
    Пытается определить, включена ли задача, по значению поля «Состояние».
    Возвращает True/False/None (неизвестно).

    Учитываем и EN-, и RU-локаль, а также runtime-статусы вроде Ready/Running.
    """
    v = (value or '').strip().lower()
    if not v:
        return None
    # Disabled всегда выигрывает — проверяем ПЕРВЫМ
    disabled_markers = ('disabled', 'отключ', 'выключ', 'запрещ')
    for m in disabled_markers:
        if m in v:
            return False
    enabled_markers = ('enabled', 'ready', 'running', 'готов', 'выполн', 'работ', 'включ')
    for m in enabled_markers:
        if m in v:
            return True
    return None


def get_task_info(task_name: str = None) -> TaskInfo:
    name = task_name or config.TASK_NAME
    r = _run_schtasks(['/Query', '/TN', name, '/FO', 'LIST', '/V'])
    if r.returncode != 0:
        logger.debug(f"task_exists FAIL (rc={r.returncode}): {r.stderr[:200]}")
        return TaskInfo(exists=False)

    # Полный дамп в DEBUG — чтобы при проблемах можно было увидеть сырой вывод
    logger.debug(f"schtasks /Query /V /FO LIST для '{name}':\n{r.stdout}")

    info = TaskInfo(exists=True, raw=r.stdout)
    state_from_task_state = None   # из "Scheduled Task State" / "Состояние запланированной задачи"
    state_from_status = None       # из "Status" / "Состояние"

    for raw_line in r.stdout.splitlines():
        line = raw_line.strip()
        if ':' not in line:
            continue
        key, _, value = line.partition(':')
        key_low = key.strip().lower()
        val = value.strip()
        low = line.lower()

        # --- Состояние (enabled/disabled) ---
        # "Scheduled Task State" или "Состояние запланированной задачи"
        if 'scheduled task state' in key_low or (
                'состоян' in key_low and ('запланированн' in key_low or 'задач' in key_low)):
            state_from_task_state = val
        # "Status" / "Состояние" (runtime state: Ready/Running/Disabled)
        elif key_low in ('status', 'состояние'):
            state_from_status = val
        elif low.startswith('schedule type') or low.startswith('тип расписания'):
            info.schedule = val
        elif low.startswith('start time') or low.startswith('время начала'):
            if not info.schedule.endswith(line.split(':', 1)[-1].strip()):
                info.schedule += ' @ ' + line.split(':', 1)[-1].strip()
        elif low.startswith('days') or low.startswith('дни'):
            info.schedule += ' (' + line.split(':', 1)[-1].strip() + ')'
        elif low.startswith('last run time') or low.startswith('время последнего'):
            info.last_run = val
        elif low.startswith('next run time') or low.startswith('следующее время'):
            info.next_run = val

    # --- Определяем enabled: сначала по Scheduled Task State, потом по Status ---
    for candidate in (state_from_task_state, state_from_status):
        if candidate is None:
            continue
        result = _classify_enabled(candidate)
        if result is not None:
            info.enabled = result
            logger.debug(f"Определил enabled={info.enabled} по значению {candidate!r}")
            break
    else:
        # Ни одно поле не распарсили — ругаемся в лог и считаем включённой
        # (если задача физически существует и нашлась по /Query).
        logger.warning(
            f"Не удалось определить enabled из полей "
            f"state_from_task_state={state_from_task_state!r}, "
            f"state_from_status={state_from_status!r}. "
            f"По умолчанию считаем включённой."
        )
        info.enabled = True

    return info


def delete_task(task_name: str = None) -> None:
    name = task_name or config.TASK_NAME
    try:
        if task_exists(name):
            _run_schtasks(['/Delete', '/TN', name, '/F'], check=True)
            logger.info(f"Задача '{name}' удалена.")
        else:
            logger.info(f"Задача '{name}' не существовала — ничего не удалено.")
    except Exception:
        logger.exception(f"Ошибка при удалении задачи '{name}'")
        raise


def enable_task(enabled: bool, task_name: str = None) -> None:
    name = task_name or config.TASK_NAME
    flag = '/ENABLE' if enabled else '/DISABLE'
    try:
        _run_schtasks(['/Change', '/TN', name, flag], check=True)
        logger.info(f"Задача '{name}' {'включена' if enabled else 'отключена'}.")
    except Exception:
        logger.exception(f"Ошибка при изменении состояния задачи '{name}'")
        raise


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


def _remap_unc_for_scheduler(path: str) -> str:
    """
    Подменяет канонические UNC-префиксы (\\\\fs05\\all\\...) на DFS-префиксы
    (\\\\kari.local\\public\\all\\...) согласно config.SCHEDULER_UNC_REMAP.

    Нужно потому, что Windows DFS при чтении sys.executable/Path(__file__)
    даёт каноническое имя сервера, а у учётной записи планировщика прав
    на него может не быть.
    """
    remap = getattr(config, 'SCHEDULER_UNC_REMAP', None) or {}
    low = path.lower()
    for canonical, friendly in remap.items():
        if low.startswith(canonical.lower()):
            replaced = friendly + path[len(canonical):]
            logger.info(f"UNC-ремап: {path}  →  {replaced}")
            return replaced
    return path


def build_tr_command(use_cache: bool = False) -> str:
    """
    Собирает строку для /TR. Использует config.get_runner_command, который
    автоматически подбирает exe или python+скрипт в зависимости от того,
    что доступно. Применяет UNC-ремап (config.SCHEDULER_UNC_REMAP).
    """
    cmd = config.get_runner_command(use_cache=use_cache)
    # Применяем UNC-ремап к каждому аргументу, похожему на путь
    remapped = []
    for part in cmd:
        if part.startswith('\\\\') or part.startswith('//'):
            remapped.append(_remap_unc_for_scheduler(part))
        else:
            remapped.append(part)
    return ' '.join(_quote_for_tr(p) for p in remapped)


def create_or_update_task(days_flags: Iterable[bool],
                          time_str: str,
                          task_name: str = None,
                          use_cache: bool = False) -> None:
    """
    Создаёт или обновляет задачу. Идемпотентно: если существует — удаляем и создаём заново.
    """
    name = task_name or config.TASK_NAME
    logger.info(f"Создание/обновление задачи '{name}': "
                f"days={list(days_flags)}, time={time_str}, use_cache={use_cache}")
    try:
        # Проверяем время
        if len(time_str) != 5 or time_str[2] != ':':
            raise ValueError(f"time_str должен быть 'HH:MM', получено: {time_str!r}")
        hh, mm = time_str.split(':')
        if not (hh.isdigit() and mm.isdigit()):
            raise ValueError("time_str должен быть 'HH:MM'")

        days_str = days_from_weekday_flags(days_flags)

        # Автовыбор: exe, если есть; иначе python + run_otbor.py
        tr_cmd = build_tr_command(use_cache=use_cache)
        logger.info(f"Задача будет запускать: {tr_cmd}")

        # Удаляем старую, если есть
        if task_exists(name):
            _run_schtasks(['/Delete', '/TN', name, '/F'], check=True)
            logger.info(f"Старая задача '{name}' удалена перед пересозданием.")

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
        logger.info(f"Задача '{name}' создана (WEEKLY, {days_str}, {time_str}).")
    except Exception:
        logger.exception(f"Ошибка при создании/обновлении задачи '{name}'")
        raise


def run_task_now(task_name: str = None) -> None:
    """Запускает задачу немедленно через планировщик."""
    name = task_name or config.TASK_NAME
    _run_schtasks(['/Run', '/TN', name], check=True)

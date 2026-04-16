"""
GUI для запуска сборки файла Отбор WB.

Возможности:
  * Кнопка «Запустить сейчас» — синхронный прогон pipeline в фоновом потоке.
  * Флажок «Использовать кэш фото» — пропустить обращение к WB API.
  * Планировщик Windows Task Scheduler:
      - выбор дней недели и времени
      - кнопка «Сохранить задачу» (идемпотентно: создаёт или обновляет)
      - «Включить / Отключить» задачу
      - «Удалить задачу»
      - статус текущей задачи (включена, расписание, последний/следующий запуск)
  * Лог-окно со всем выводом pipeline.
"""
from __future__ import annotations

import queue
import subprocess
import sys
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import ttk, messagebox

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from wb_otbor import config, scheduler
from wb_otbor.pipeline import run_full_pipeline


# Используем config.BASE_DIR — он корректно резолвится и из python, и из exe
# (SCRIPT_DIR через __file__ в frozen-режиме указывает на _MEIPASS, а не на проект)
CLI_SCRIPT = config.BASE_DIR / 'run_otbor.py'
BUILD_SCRIPT = config.BASE_DIR / 'build_exe.py'


class OtborGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("WB Отбор — автоматизация")
        self.root.geometry("900x700")
        self.log_queue: queue.Queue[str] = queue.Queue()
        self._worker: threading.Thread | None = None
        self._build_ui()
        self._poll_log_queue()
        self.refresh_runner_status()
        self.refresh_task_status()

    # ----------------------------- UI ------------------------------------

    def _build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass

        container = ttk.Frame(self.root, padding=10)
        container.pack(fill='both', expand=True)

        # --- Секция "Запуск сейчас" ----------------------------------------
        run_frame = ttk.LabelFrame(container, text="Ручной запуск", padding=10)
        run_frame.pack(fill='x', pady=(0, 10))

        self.use_cache_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            run_frame,
            text="Использовать кэш фото (photo_cache.xlsx) — без обращения к WB API",
            variable=self.use_cache_var,
        ).pack(anchor='w')

        buttons_row = ttk.Frame(run_frame)
        buttons_row.pack(fill='x', pady=(8, 0))

        self.run_btn = ttk.Button(
            buttons_row, text="▶ Запустить сейчас",
            command=self.on_run_now,
        )
        self.run_btn.pack(side='left')

        ttk.Button(
            buttons_row, text="Открыть папку проекта",
            command=self.on_open_folder,
        ).pack(side='left', padx=(8, 0))

        # --- Секция "Сборка EXE" -------------------------------------------
        exe_frame = ttk.LabelFrame(container, text="Сборка EXE (для запуска на других ПК)",
                                   padding=10)
        exe_frame.pack(fill='x', pady=(0, 10))

        self.runner_status_var = tk.StringVar()
        ttk.Label(exe_frame, textvariable=self.runner_status_var,
                  font=('Segoe UI', 9)).pack(anchor='w')

        exe_btn_row = ttk.Frame(exe_frame)
        exe_btn_row.pack(fill='x', pady=(6, 0))

        self.build_btn = ttk.Button(
            exe_btn_row, text="🔨 Собрать EXE (PyInstaller)",
            command=self.on_build_exe,
        )
        self.build_btn.pack(side='left')
        ttk.Button(exe_btn_row, text="⟳ Обновить статус",
                   command=self.refresh_runner_status).pack(side='left', padx=(8, 0))
        ttk.Label(exe_btn_row,
                  text="  Требует Python + PyInstaller в PATH",
                  foreground='#666', font=('Segoe UI', 8)).pack(side='left')

        # --- Секция "Планировщик" ------------------------------------------
        sched_frame = ttk.LabelFrame(
            container,
            text=f"Планировщик задач Windows (имя задачи: {config.TASK_NAME})",
            padding=10,
        )
        sched_frame.pack(fill='x', pady=(0, 10))

        # Дни недели
        days_row = ttk.Frame(sched_frame)
        days_row.pack(fill='x')
        ttk.Label(days_row, text="Дни недели:").pack(side='left')
        self.day_vars = []
        for i, name in enumerate(scheduler.DAYS_RU):
            v = tk.BooleanVar(value=(i < 5))  # по умолчанию Пн–Пт
            cb = ttk.Checkbutton(days_row, text=name, variable=v)
            cb.pack(side='left', padx=4)
            self.day_vars.append(v)

        # Время
        time_row = ttk.Frame(sched_frame)
        time_row.pack(fill='x', pady=(8, 0))
        ttk.Label(time_row, text="Время запуска (HH:MM):").pack(side='left')
        self.time_var = tk.StringVar(value="09:00")
        ttk.Entry(time_row, textvariable=self.time_var, width=8).pack(side='left', padx=(4, 10))

        self.task_use_cache_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            time_row,
            text="Задача запускает с --use-cache",
            variable=self.task_use_cache_var,
        ).pack(side='left')

        # Кнопки управления задачей
        ctrl_row = ttk.Frame(sched_frame)
        ctrl_row.pack(fill='x', pady=(8, 0))

        ttk.Button(ctrl_row, text="Сохранить задачу",
                   command=self.on_save_task).pack(side='left')
        ttk.Button(ctrl_row, text="Включить",
                   command=lambda: self.on_toggle_task(True)).pack(side='left', padx=4)
        ttk.Button(ctrl_row, text="Отключить",
                   command=lambda: self.on_toggle_task(False)).pack(side='left', padx=4)
        ttk.Button(ctrl_row, text="Удалить задачу",
                   command=self.on_delete_task).pack(side='left', padx=4)
        ttk.Button(ctrl_row, text="⟳ Обновить статус",
                   command=self.refresh_task_status).pack(side='right')

        # Статус задачи
        self.status_var = tk.StringVar(value="Статус: неизвестно")
        ttk.Label(sched_frame, textvariable=self.status_var,
                  foreground='#444', font=('Segoe UI', 9)).pack(anchor='w', pady=(8, 0))

        # --- Лог-окно ------------------------------------------------------
        log_frame = ttk.LabelFrame(container, text="Журнал выполнения", padding=6)
        log_frame.pack(fill='both', expand=True)

        self.log_text = tk.Text(log_frame, wrap='word', height=20,
                                font=('Consolas', 9), bg='#111', fg='#ddd',
                                insertbackground='#ddd')
        self.log_text.pack(side='left', fill='both', expand=True)
        log_scroll = ttk.Scrollbar(log_frame, orient='vertical', command=self.log_text.yview)
        log_scroll.pack(side='right', fill='y')
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.configure(state='disabled')

    # ------------------------- Логирование -------------------------------

    def log(self, msg: str):
        """Может вызываться из любого потока."""
        self.log_queue.put(msg)

    def _poll_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self._append_log(msg)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._poll_log_queue)

    def _append_log(self, msg: str):
        self.log_text.configure(state='normal')
        self.log_text.insert('end', msg + '\n')
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    # --------------------- Кнопка "Запустить сейчас" ---------------------

    def on_run_now(self):
        if self._worker and self._worker.is_alive():
            messagebox.showinfo("Выполняется", "Сборка уже запущена, дождитесь завершения.")
            return

        use_cache = self.use_cache_var.get()
        self.run_btn.configure(state='disabled', text='Выполняется...')
        self.log("=" * 60)
        self.log(f"Старт ручного запуска. use_cache={use_cache}")

        def worker():
            try:
                run_full_pipeline(use_photo_cache=use_cache, log=self.log)
            except Exception as exc:
                self.log(f"ОШИБКА: {exc}")
                self.log(traceback.format_exc())
            finally:
                self.root.after(0, lambda: self.run_btn.configure(
                    state='normal', text='▶ Запустить сейчас'))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def on_open_folder(self):
        import os
        os.startfile(str(config.BASE_DIR))

    # ------------------------- Планировщик -------------------------------

    def on_save_task(self):
        try:
            flags = [v.get() for v in self.day_vars]
            time_str = self.time_var.get().strip()
            scheduler.create_or_update_task(
                days_flags=flags,
                time_str=time_str,
                use_cache=self.task_use_cache_var.get(),
            )
            runner_cmd = ' '.join(config.get_runner_command(self.task_use_cache_var.get()))
            messagebox.showinfo("Готово",
                f"Задача сохранена в планировщике.\n\nБудет запускать:\n{runner_cmd}")
            self.log(f"Задача создана/обновлена: {runner_cmd}")
            self.refresh_task_status()
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))
            self.log(f"Ошибка сохранения задачи: {exc}")

    def refresh_runner_status(self):
        exe = config.find_runner_exe()
        if exe:
            self.runner_status_var.set(
                f"✓ EXE найден: {exe}\n"
                f"   Задача планировщика будет запускать ИМЕННО ЕГО."
            )
        else:
            self.runner_status_var.set(
                f"✗ EXE не собран. Задача будет запускать python + {CLI_SCRIPT.name}\n"
                f"   Для запуска на ПК без Python — нажмите 'Собрать EXE' ниже."
            )

    @staticmethod
    def _find_python() -> str | None:
        """
        Ищет python.exe на машине. Нужен для сборки EXE из frozen-GUI,
        где sys.executable указывает на сам exe, а не на python.
        """
        import shutil
        # 1. Если не frozen — sys.executable и есть python
        if not getattr(sys, 'frozen', False):
            return sys.executable
        # 2. py.exe — Windows Python Launcher (ставится вместе с Python)
        py = shutil.which('py')
        if py:
            return py
        # 3. python.exe в PATH
        python = shutil.which('python')
        if python:
            return python
        # 4. python3.exe
        python3 = shutil.which('python3')
        if python3:
            return python3
        return None

    def on_build_exe(self):
        python = self._find_python()
        if not python:
            messagebox.showwarning(
                "Python не найден",
                "Для сборки EXE нужен Python с PyInstaller.\n\n"
                "Установите Python и убедитесь, что он добавлен в PATH,\n"
                "либо запустите:  python run_gui.py")
            return
        if self._worker and self._worker.is_alive():
            messagebox.showinfo("Занято", "Другая операция уже выполняется.")
            return
        if not messagebox.askyesno(
                "Сборка EXE",
                f"Будет использован Python:\n{python}\n\n"
                "Это займёт 2-5 минут. Продолжить?\n"
                "(PyInstaller будет поставлен автоматически, если отсутствует)"):
            return

        self.log("=" * 60)
        self.log(f"Запуск сборки EXE через build_exe.py (python: {python})...")

        def worker():
            try:
                proc = subprocess.Popen(
                    [python, str(BUILD_SCRIPT), '--target', 'runner', '--clean'],
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    cwd=str(config.BASE_DIR), text=True, encoding='utf-8',
                    errors='replace', creationflags=0x08000000,  # CREATE_NO_WINDOW
                )
                for line in proc.stdout:
                    self.log(line.rstrip())
                rc = proc.wait()
                if rc == 0:
                    self.log("Сборка EXE завершена успешно.")
                    self.root.after(0, self.refresh_runner_status)
                else:
                    self.log(f"Сборка EXE завершилась с ошибкой (код {rc}).")
            except Exception as exc:
                self.log(f"ОШИБКА сборки: {exc}")
                self.log(traceback.format_exc())

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def on_toggle_task(self, enabled: bool):
        try:
            if not scheduler.task_exists():
                messagebox.showwarning("Нет задачи",
                                       "Задача не найдена. Сначала сохраните её.")
                return
            scheduler.enable_task(enabled)
            action = "включена" if enabled else "отключена"
            messagebox.showinfo("Готово", f"Задача {action}.")
            self.log(f"Задача {action}.")
            self.refresh_task_status()
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))

    def on_delete_task(self):
        if not scheduler.task_exists():
            messagebox.showinfo("Нет задачи", "Задача отсутствует в планировщике.")
            self.refresh_task_status()
            return
        if not messagebox.askyesno("Подтверждение",
                                    "Удалить задачу из планировщика?"):
            return
        try:
            scheduler.delete_task()
            self.log("Задача удалена из Task Scheduler.")
            self.refresh_task_status()
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))

    def refresh_task_status(self):
        try:
            info = scheduler.get_task_info()
            if not info.exists:
                self.status_var.set("Статус: задача не создана")
                return
            state = "✓ включена" if info.enabled else "✗ отключена"
            parts = [f"Статус: {state}"]
            if info.schedule:
                parts.append(f"расписание: {info.schedule}")
            if info.next_run:
                parts.append(f"следующий запуск: {info.next_run}")
            if info.last_run:
                parts.append(f"последний запуск: {info.last_run}")
            self.status_var.set(" | ".join(parts))
        except Exception as exc:
            self.status_var.set(f"Статус: ошибка чтения ({exc})")


def main():
    root = tk.Tk()
    try:
        # иконка если есть
        pass
    except Exception:
        pass
    OtborGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()

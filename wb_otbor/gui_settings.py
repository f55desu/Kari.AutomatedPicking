"""
Окно настроек — фильтры справочника + числовые параметры отчёта.

Открывается из run_gui.py по кнопке "⚙ Настройки".

Функциональность:
  - 7 фильтров (Бизнес-группа / Розничный отдел / Группа / Сезон / Бренд /
    Ответственный за группу / Коллекция) — multi-select Listbox'ы.
  - Кнопка "🔄 Получить уникальные значения из БД" — загружает все уникальные
    значения справочника и кеширует их в settings.json.
  - Числовые поля: период, оффсет, пороги Технички (min_stock / min_distrib / min_shows).
  - "Применить параметры" — сохраняет в settings.json.
"""
from __future__ import annotations

import threading
import tkinter as tk
import traceback
from tkinter import ttk, messagebox

from . import settings as app_settings
from .logging_setup import get_logger


logger = get_logger('gui_settings')


class SettingsWindow(tk.Toplevel):
    def __init__(self, parent, log=print, on_applied=None):
        super().__init__(parent)
        self.title("Настройки фильтров и параметров")
        self.geometry("1280x780")
        self.minsize(1100, 700)
        self.transient(parent)
        self.grab_set()

        self.log = log
        self.on_applied = on_applied
        self._settings = app_settings.load()
        self._filter_listboxes: dict[str, tk.Listbox] = {}
        self._numeric_vars: dict[str, tk.StringVar] = {}
        self._build_ui()
        self._populate_from_settings()
        self._refresh_summary()
        # Отвязываем глобальные mousewheel-биндинги при закрытии,
        # чтобы они не мешали главному окну.
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Destroy>", self._on_destroy)

    def _on_close(self):
        self.destroy()

    def _on_destroy(self, event):
        if event.widget is self:
            try:
                self.unbind_all('<MouseWheel>')
                self.unbind_all('<Button-4>')
                self.unbind_all('<Button-5>')
            except tk.TclError:
                pass

    # ----------------------------- UI ------------------------------------

    def _build_ui(self):
        # --- ВЕРХ (fixed): сводка + кнопка fetch ---
        top_wrap = ttk.Frame(self, padding=(10, 10, 10, 0))
        top_wrap.pack(fill='x')

        summary_frame = ttk.LabelFrame(top_wrap, text="Текущий выбор", padding=6)
        summary_frame.pack(fill='x')
        self.summary_var = tk.StringVar(value="(загрузка...)")
        ttk.Label(summary_frame, textvariable=self.summary_var,
                  font=('Segoe UI', 9), wraplength=1240, justify='left',
                  foreground='#1a4480').pack(anchor='w', fill='x')

        fetch_row = ttk.Frame(self, padding=(10, 6, 10, 6))
        fetch_row.pack(fill='x')
        self.fetch_btn = ttk.Button(
            fetch_row,
            text="🔄 Получить уникальные значения параметров из БД",
            command=self.on_fetch_unique,
        )
        self.fetch_btn.pack(side='left')
        updated = self._settings.get('unique_values_updated')
        self.updated_var = tk.StringVar(
            value=f"Кэш обновлён: {updated}" if updated else "Кэш пуст — нажмите кнопку"
        )
        ttk.Label(fetch_row, textvariable=self.updated_var,
                  foreground='#666', font=('Segoe UI', 9)).pack(side='left', padx=(10, 0))

        # --- НИЗ (fixed): кнопки — packим ДО центра, чтобы они остались видны ---
        btn_frame = ttk.Frame(self, padding=(10, 6, 10, 10))
        btn_frame.pack(side='bottom', fill='x')
        ttk.Button(btn_frame, text="Применить параметры",
                   command=self.on_apply).pack(side='right', padx=(4, 0))
        ttk.Button(btn_frame, text="Отмена",
                   command=self.destroy).pack(side='right')
        ttk.Button(btn_frame, text="Сбросить к дефолтам",
                   command=self.on_reset_defaults).pack(side='left')

        # --- ЦЕНТР (scrollable): Canvas + Scrollbar ---
        scroll_wrap = ttk.Frame(self)
        scroll_wrap.pack(fill='both', expand=True, padx=10, pady=(0, 6))

        canvas = tk.Canvas(scroll_wrap, highlightthickness=0)
        vsb = ttk.Scrollbar(scroll_wrap, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        # Внутренний фрейм, в котором всё содержимое
        inner = ttk.Frame(canvas)
        inner_id = canvas.create_window((0, 0), window=inner, anchor='nw')

        # Сделаем inner растягивающимся под ширину canvas
        def _on_canvas_resize(event):
            canvas.itemconfigure(inner_id, width=event.width)
        canvas.bind('<Configure>', _on_canvas_resize)
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox('all')))

        # Mouse-wheel прокрутка (Windows / Mac / Linux)
        def _on_mousewheel(event):
            if event.num == 4:      # Linux up
                canvas.yview_scroll(-3, 'units')
            elif event.num == 5:    # Linux down
                canvas.yview_scroll(3, 'units')
            else:                   # Windows / Mac
                canvas.yview_scroll(int(-event.delta / 120) * 3, 'units')
        # Биндим на все виджеты окна — чтобы колесо работало везде
        self.bind_all('<MouseWheel>', _on_mousewheel)
        self.bind_all('<Button-4>', _on_mousewheel)
        self.bind_all('<Button-5>', _on_mousewheel)

        # --- Разметка внутреннего фрейма: слева фильтры, справа числовые ---
        inner.columnconfigure(0, weight=3, uniform='main')   # фильтры
        inner.columnconfigure(1, weight=1, uniform='main')   # числовые
        inner.rowconfigure(0, weight=1)

        # ЛЕВО: фильтры (2 колонки × 4 ряда)
        filters_frame = ttk.LabelFrame(inner,
                                        text="Фильтры справочника (multi-select)",
                                        padding=8)
        filters_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 6))
        for col in range(2):
            filters_frame.columnconfigure(col, weight=1, uniform='filter_col')

        positions = [
            ('Бизнес-группа',           0, 0),
            ('Розничный отдел',         0, 1),
            ('Группа',                  1, 0),
            ('Сезон',                   1, 1),
            ('Бренд',                   2, 0),
            ('Ответственный за группу', 2, 1),
            ('Коллекция',               3, 0),
        ]
        for key, row, col in positions:
            self._build_filter_box(filters_frame, key, row, col)

        # ПРАВО: числовые параметры (вертикальный стек)
        numeric_frame = ttk.LabelFrame(inner, text="Параметры отчёта", padding=8)
        numeric_frame.grid(row=0, column=1, sticky='nsew', padx=(6, 0))
        numeric_frame.columnconfigure(0, weight=1)

        self._build_numeric_field(numeric_frame, 'period_days',
                                   "Длина периода (дней):", 0)
        self._build_numeric_field(numeric_frame, 'offset_from_today',
                                   "Сдвиг от сегодня (дней):", 1)
        ttk.Separator(numeric_frame, orient='horizontal').grid(
            row=2, column=0, sticky='ew', pady=8)
        ttk.Label(numeric_frame, text="Пороги «Технички»:",
                  font=('Segoe UI', 9, 'bold')).grid(row=3, column=0, sticky='w', pady=(0, 4))
        self._build_numeric_field(numeric_frame, 'min_stock',
                                   "Мин. остаток (шт.):", 4)
        self._build_numeric_field(numeric_frame, 'min_distrib_percent',
                                   "Мин. дистрибуция (%):", 5)
        self._build_numeric_field(numeric_frame, 'min_shows',
                                   "Мин. показы:", 6)

    def _build_filter_box(self, parent, key: str, row: int, col: int):
        box = ttk.LabelFrame(parent, text=key, padding=4)
        box.grid(row=row, column=col, sticky='nsew', padx=4, pady=4)
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)

        lb = tk.Listbox(box, selectmode='extended', exportselection=False,
                         height=6, font=('Segoe UI', 9))
        lb.grid(row=0, column=0, sticky='nsew')
        sb = ttk.Scrollbar(box, orient='vertical', command=lb.yview)
        sb.grid(row=0, column=1, sticky='ns')
        lb.configure(yscrollcommand=sb.set)
        # Обновление сводки при изменении выделения
        lb.bind('<<ListboxSelect>>', lambda e: self._refresh_summary())

        # Кнопки "Всё"/"Ничего"
        btn_row = ttk.Frame(box)
        btn_row.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(2, 0))
        ttk.Button(btn_row, text="Выделить всё", width=14,
                   command=lambda lb=lb: (lb.selection_set(0, 'end'), self._refresh_summary())
                   ).pack(side='left')
        ttk.Button(btn_row, text="Снять выделение", width=14,
                   command=lambda lb=lb: (lb.selection_clear(0, 'end'), self._refresh_summary())
                   ).pack(side='left', padx=(4, 0))

        self._filter_listboxes[key] = lb

    def _build_numeric_entry(self, parent, key: str, label: str,
                              row: int, col: int):
        """Старый горизонтальный вариант (label + entry в соседних колонках)."""
        ttk.Label(parent, text=label).grid(row=row, column=col,
                                           sticky='w', padx=4, pady=2)
        var = tk.StringVar()
        var.trace_add('write', lambda *args: self._refresh_summary())
        entry = ttk.Entry(parent, textvariable=var, width=12)
        entry.grid(row=row, column=col + 1, sticky='w', padx=4, pady=2)
        self._numeric_vars[key] = var

    def _build_numeric_field(self, parent, key: str, label: str, row: int):
        """Вертикальный вариант: label сверху, entry снизу на всю ширину."""
        wrap = ttk.Frame(parent)
        wrap.grid(row=row, column=0, sticky='ew', pady=3)
        wrap.columnconfigure(1, weight=1)
        ttk.Label(wrap, text=label, font=('Segoe UI', 9)).grid(
            row=0, column=0, sticky='w')
        var = tk.StringVar()
        var.trace_add('write', lambda *args: self._refresh_summary())
        ttk.Entry(wrap, textvariable=var, width=10, font=('Segoe UI', 10)).grid(
            row=0, column=1, sticky='e', padx=(6, 0))
        self._numeric_vars[key] = var

    # ---------------------- Обновление сводки ----------------------------

    def _refresh_summary(self):
        """Собирает одну строку из текущего выбора и обновляет лейбл сверху."""
        parts: list[str] = []

        # Фильтры
        for key, lb in self._filter_listboxes.items():
            sel = lb.curselection()
            if not sel:
                parts.append(f"{key}: все")
                continue
            values = [lb.get(i) for i in sel]
            if len(values) <= 3:
                parts.append(f"{key}: {', '.join(values)}")
            else:
                shown = ', '.join(values[:2])
                parts.append(f"{key}: {shown}… ({len(values)} шт.)")

        # Числовые
        def _g(k):
            try:
                return self._numeric_vars[k].get().strip()
            except KeyError:
                return '?'

        parts.append(f"Период: {_g('period_days')}д, сдвиг: {_g('offset_from_today')}д")
        parts.append(f"Пороги: остаток≥{_g('min_stock')}, "
                     f"дистрибуция≥{_g('min_distrib_percent')}%, "
                     f"показы≥{_g('min_shows')}")

        self.summary_var.set("  |  ".join(parts))

    # ---------------------- Заполнение значениями -----------------------

    def _populate_from_settings(self):
        """Заполняет все UI-элементы значениями из текущих настроек."""
        cache = self._settings.get('unique_values_cache', {})
        selected = self._settings.get('filters', {})

        for key, lb in self._filter_listboxes.items():
            lb.delete(0, 'end')
            values = cache.get(key, [])
            # Если кэша нет, но в настройках есть выбранные значения —
            # покажем их (чтобы пользователь видел что уже задано).
            if not values:
                values = list(selected.get(key, []))
            for v in values:
                lb.insert('end', v)
            # Выделяем те, что уже выбраны
            selected_vals = set(selected.get(key, []))
            for i, v in enumerate(values):
                if v in selected_vals:
                    lb.selection_set(i)

        for key, var in self._numeric_vars.items():
            var.set(str(self._settings.get(key, app_settings.NUMERIC_DEFAULTS[key])))

    # -------------------- Загрузка уникальных значений -------------------

    def on_fetch_unique(self):
        self.fetch_btn.configure(state='disabled', text='Загрузка... подождите')
        self.updated_var.set("Загрузка уникальных значений из БД...")

        def worker():
            try:
                from . import sql_loader
                values = sql_loader.load_unique_filter_values()
                app_settings.update_unique_values(values)
                self.log(f"Уникальные значения получены: "
                         f"{', '.join(f'{k}={len(v)}' for k, v in values.items())}")
                # Обновить UI в главном потоке
                self.after(0, self._on_fetch_done)
            except Exception as exc:
                self.log(f"Ошибка получения уникальных значений: {exc}")
                self.log(traceback.format_exc())
                self.after(0, lambda: (
                    self.fetch_btn.configure(state='normal',
                        text="🔄 Получить уникальные значения параметров из БД"),
                    messagebox.showerror("Ошибка", f"Не удалось получить значения:\n{exc}"),
                ))

        threading.Thread(target=worker, daemon=True).start()

    def _on_fetch_done(self):
        self._settings = app_settings.load()
        self._populate_from_settings()
        self._refresh_summary()
        self.fetch_btn.configure(state='normal',
            text="🔄 Получить уникальные значения параметров из БД")
        updated = self._settings.get('unique_values_updated')
        self.updated_var.set(f"Кэш обновлён: {updated}")

    # --------------------------- Применение ------------------------------

    def on_apply(self):
        try:
            # Собираем выбранные значения фильтров
            selected_filters: dict[str, list[str]] = {}
            for key, lb in self._filter_listboxes.items():
                sel_indices = lb.curselection()
                selected_filters[key] = [lb.get(i) for i in sel_indices]

            # Собираем числовые параметры
            try:
                numerics = {key: int(var.get().strip())
                            for key, var in self._numeric_vars.items()}
            except ValueError:
                logger.warning("on_apply: числовые параметры не int")
                messagebox.showerror("Ошибка",
                    "Числовые параметры должны быть целыми числами.")
                return
            if numerics['period_days'] < 1 or numerics['offset_from_today'] < 0:
                logger.warning(f"on_apply: некорректные числа {numerics}")
                messagebox.showerror("Ошибка",
                    "Период должен быть ≥ 1, сдвиг ≥ 0.")
                return

            # Сохраняем — merge с текущими (чтобы не потерять unique_values_cache)
            data = app_settings.load()
            data['filters'] = selected_filters
            for key, val in numerics.items():
                data[key] = val
            path = app_settings.save(data)

            active = sum(1 for v in selected_filters.values() if v)
            logger.info(f"Настройки применены: активных фильтров {active}, "
                        f"{numerics}")
            msg = (f"Настройки сохранены в {path}\n\n"
                   f"Активных фильтров: {active}")
            self.log(f"Настройки применены (сохранены в {path.name})")
            messagebox.showinfo("Готово", msg)

            if callable(self.on_applied):
                try:
                    self.on_applied()
                except Exception:
                    logger.exception("Ошибка в callback on_applied")
            self.destroy()
        except Exception as exc:
            logger.exception("Ошибка on_apply")
            messagebox.showerror("Ошибка", f"Не удалось сохранить настройки:\n{exc}")

    def on_reset_defaults(self):
        if not messagebox.askyesno("Подтверждение",
                "Сбросить все настройки к значениям по умолчанию?"):
            return
        # Сбрасываем значения в UI, не трогая кэш уникальных значений
        self._settings = app_settings.load()  # перечитываем cache
        # Применяем DEFAULTS к фильтрам и числам
        from copy import deepcopy
        self._settings['filters'] = deepcopy(app_settings.DEFAULTS['filters'])
        for k, v in app_settings.NUMERIC_DEFAULTS.items():
            self._settings[k] = v
        self._populate_from_settings()

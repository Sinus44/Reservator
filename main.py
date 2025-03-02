import os
import shutil
import zipfile
import json
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, timedelta
import pystray
from PIL import Image
import sys
import logging
import calendar
import threading


def resource_path(relative_path):
    """ Получает абсолютный путь к ресурсу """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


ICON_PATH = resource_path("icon.ico")


def str_or_bytes_to_str(val: str | bytes) -> str:
    if isinstance(val, bytes):
        return val.decode()

    return val


# Классы данных
class BackupTask:
    def __init__(self, name, source, destination, compression, frequency, time_params):
        self.name = name
        self.source = source
        self.destination = destination
        self.compression = compression
        self.frequency = frequency
        self.time_params = time_params
        self.last_run = None
        self.next_run = self.calculate_next_run()

    def calculate_next_run(self, now=None):
        now = now or datetime.now()
        try:
            if self.frequency == "hourly":
                # Пример: время_params = 30 (минута)
                next_run = now.replace(minute=self.time_params, second=0, microsecond=0)
                if next_run <= now:
                    next_run += timedelta(hours=1)
                return next_run

            elif self.frequency == "daily":
                # Пример: time_params = (15, 30) (час, минута)
                hour, minute = self.time_params
                next_run = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
                if next_run <= now:
                    next_run += timedelta(days=1)
                return next_run

            elif self.frequency == "weekly":
                # Пример: time_params = (2, 15, 30) (день_недели, час, минута)
                weekday, hour, minute = self.time_params
                days_ahead = (weekday - now.weekday()) % 7
                next_run = now.replace(hour=hour, minute=minute, second=0, microsecond=0) + timedelta(days=days_ahead)
                if next_run <= now:
                    next_run += timedelta(weeks=1)
                return next_run

            elif self.frequency == "monthly":
                # Пример: time_params = (5, 15, 30) (день, час, минута)
                day, hour, minute = self.time_params
                next_month = now.month + 1
                next_year = now.year
                if next_month > 12:
                    next_month = 1
                    next_year += 1

                try:
                    next_run = datetime(next_year, next_month, day, hour, minute)
                except ValueError:
                    last_day = calendar.monthrange(next_year, next_month)[1]
                    next_run = datetime(next_year, next_month, last_day, hour, minute)

                return next_run if next_run > now else self.calculate_next_run(now + timedelta(days=1))

            return now

        except Exception as e:
            logging.error(f"Ошибка расчета времени для задачи {self.name}: {str(e)}")
            return now + timedelta(minutes=5)  # Fallback

    def to_dict(self):
        return {
            "name": self.name,
            "source": self.source,
            "destination": self.destination,
            "compression": self.compression,
            "frequency": self.frequency,
            "time_params": self.time_params,
            "last_run": self.last_run.isoformat() if self.last_run else None,
        }


# Основное приложение
class BackupApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.tree = None
        self.last_click_time = None
        self.tray_icon_lock = threading.Lock()
        self.tray_icon = None
        self.running_tasks = 0
        self.tasks = []
        self.config = {
            "compression_level": 9,
        }

        self.title("Backup Tools By Deepseek")
        self.geometry("800x600")

        try:
            self.logo_image = Image.open(ICON_PATH)  # Для трея

        except Exception as e:
            logging.error(e)
            self.logo_image = None

        # Установка иконки
        try:
            self.iconbitmap(ICON_PATH)  # Для окна
        except Exception as e:
            print(f"Ошибка загрузки иконки: {str(e)}")

        # Инициализация планировщика
        self.scheduler_running = True
        self.scheduler_thread = threading.Thread(target=self.scheduler_loop, daemon=True)
        self.scheduler_thread.start()

        # Настройка трея
        self.protocol("WM_DELETE_WINDOW", self.hide_to_tray)

        self.load_config()
        self.load_tasks()

        self.create_widgets()
        self.update_status()

        self.scheduler_running = True
        self.scheduler_thread = threading.Thread(
            target=self.scheduler_loop,
            daemon=True,
            name="SchedulerThread"
        )
        self.scheduler_thread.start()

    def scheduler_loop(self):
        """Основной цикл планировщика, работает независимо от GUI"""
        while self.scheduler_running:
            try:
                now = datetime.now()
                tasks_to_run = []

                # Проверка задач
                with threading.Lock():
                    for task in self.tasks.copy():
                        if task.next_run <= now:
                            tasks_to_run.append(task)
                            task.last_run = now
                            task.next_run = task.calculate_next_run()

                # Запуск задач
                for task in tasks_to_run:
                    logging.info(f"Запуск задачи: {task.name}")
                    self.run_task(task)

                time.sleep(5)  # Проверка каждые 5 секунд

            except Exception as e:
                logging.error(f"Ошибка планировщика: {str(e)}")
                time.sleep(10)

    def hide_to_tray(self):
        self.withdraw()
        if not self.tray_icon:
            try:
                image = self.logo_image
                menu = pystray.Menu(
                    pystray.MenuItem("Открыть", self.restore_from_tray),
                    pystray.MenuItem("Выход", self.quit_app)
                )
                self.tray_icon = pystray.Icon(
                    "Backup Tools",
                    image,
                    "Backup Tools By Deepseek",
                    menu=menu
                )
                self.tray_icon.run()
            except Exception as e:
                logging.error(f"Ошибка создания трея: {str(e)}")

    def restore_from_tray(self, icon=None, item=None):
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None

        self.deiconify()
        self.lift()
        self.focus_force()
        self.update_status()
        self.update_task_list()
        self.create_widgets()  # Добавлено пересоздание виджетов

    def quit_app(self):
        """Гарантированное завершение"""
        self.scheduler_running = False
        if self.tray_icon:
            self.tray_icon.stop()

        self.destroy()
        exit(0)

    def update_status(self):
        now = datetime.now()
        next_tasks = [t.next_run for t in self.tasks if t.next_run]
        nearest = min(next_tasks, default=None)

        # Определение статуса и цвета
        if self.running_tasks > 0:
            status_text = "Выполняется резервное копирование..."
            color = "orange"
        elif not self.tasks:
            status_text = "Нет активных задач"
            color = "red"
        else:
            status_text = "Работает в фоне"
            color = "green"

        self.status_var.set(f"Статус: {status_text}")
        self.status_label.config(foreground=color)

        # Расчет времени до следующей задачи
        if nearest is not None:
            delta = nearest - now
            if delta.total_seconds() > 0:
                hours, remainder = divmod(delta.seconds, 3600)
                minutes = remainder // 60
                next_task_text = f"Ближайшая задача через: {hours}ч {minutes}м"
            else:
                next_task_text = "Следующая задача: сейчас"
        else:
            next_task_text = "Ближайшая задача: Нет активных задач"

        self.next_task_var.set(next_task_text)
        self.after(60000, self.update_status)  # Обновляем каждую минуту

    def create_widgets(self):
        for widget in self.winfo_children():
            widget.destroy()

        # Панель инструментов
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(toolbar, text="Добавить", command=self.add_task).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Редактировать", command=self.edit_task).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Удалить", command=self.delete_task).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Настройки", command=self.open_settings).pack(side=tk.RIGHT)

        # Список задач
        self.tree = ttk.Treeview(self, columns=("name", "source", "destination", "frequency", "next_run"),
                                 show="headings")
        self.tree.heading("name", text="Имя")
        self.tree.heading("source", text="Источник")
        self.tree.heading("destination", text="Назначение")
        self.tree.heading("frequency", text="Расписание")
        self.tree.heading("next_run", text="Следующий запуск")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.status_var = tk.StringVar(value="Статус: Работает")
        self.next_task_var = tk.StringVar(value="Ближайшая задача: Нет активных задач")

        status_frame = ttk.Frame(self)
        status_frame.pack(fill=tk.X, padx=5, pady=5)
        self.status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            font=('Arial', 10, 'bold')
        )
        self.status_label.pack(side=tk.LEFT)
        ttk.Label(status_frame, textvariable=self.next_task_var).pack(side=tk.RIGHT)

        self.update_task_list()

    def update_task_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for task in self.tasks:
            self.tree.insert("", tk.END, values=(
                task.name,
                task.source,
                task.destination,
                task.frequency,
                task.next_run.strftime("%Y-%m-%d %H:%M") if task.next_run else ""
            ))

    def add_task(self):
        dialog = TaskDialog(self)
        self.wait_window(dialog)
        if dialog.result:
            self.tasks.append(dialog.result)
            self.save_tasks()
            self.update_task_list()

    def edit_task(self):
        selected = self.tree.selection()
        if not selected:
            return

        index = self.tree.index(selected[0])
        task = self.tasks[index]
        dialog = TaskDialog(self, task)
        self.wait_window(dialog)
        if dialog.result:
            self.tasks[index] = dialog.result
            self.save_tasks()
            self.update_task_list()

    def delete_task(self):
        selected = self.tree.selection()
        if not selected:
            return

        index = self.tree.index(selected[0])
        del self.tasks[index]
        self.save_tasks()
        self.update_task_list()

    def open_settings(self):
        SettingsDialog(self)

    def start_scheduler(self):
        self.check_scheduled_tasks()
        self.after(30000, self.start_scheduler)  # Проверка каждые 30 секунд

    def check_scheduled_tasks(self):
        now = datetime.now()
        for task in self.tasks:
            if task.next_run <= now:
                self.run_task(task)
                task.last_run = now
                task.next_run = task.calculate_next_run()

        self.save_tasks()
        self.after(0, self.update_task_list)
        self.after(0, self.update_status)

        self.save_tasks()
        self.update_task_list()
        self.update_status()

    def run_task(self, task):
        logging.info(f"Запуск задачи: {task.name}")

        def backup():
            try:
                # Формирование путей
                logging.info(f"1")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                dest_dir = os.path.abspath(task.destination)
                os.makedirs(dest_dir, exist_ok=True)
                logging.info(f"2")

                # Создание backup
                if task.compression:
                    logging.info(f"3")
                    zip_path = os.path.join(dest_dir, f"backup_{task.name}_{timestamp}.zip")
                    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        logging.info(f"4")
                        if os.path.isfile(task.source):
                            zipf.write(os.path.abspath(task.source), os.path.basename(task.source))
                        else:
                            for root, _, files in os.walk(os.path.abspath(task.source)):
                                for file in files:
                                    file_path = str(os.path.join(root, file))
                                    zipf.write(file_path, os.path.relpath(file_path, task.source))
                else:
                    backup_dir = os.path.join(dest_dir, f"backup_{task.name}_{timestamp}")
                    logging.info(f"5")
                    shutil.copytree(
                        os.path.abspath(task.source),
                        backup_dir,
                        copy_function=self.copy_with_errors,
                        dirs_exist_ok=True
                    )
                logging.info(f"6")
                # Принудительная запись на диск
                if hasattr(os, 'sync'):
                    os.sync()

            except Exception as e:
                logging.error(f"Ошибка: {str(e)}")

            finally:
                with threading.Lock():
                    self.running_tasks -= 1
                    self.update_status()

        # Запуск с высоким приоритетом
        threading.Thread(target=backup, daemon=True).start()

    def create_zip(self, source: str, zip_path):
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED,
                             compresslevel=self.config['compression_level']) as zipf:
            basename = str_or_bytes_to_str(os.path.basename(source))
            if os.path.isfile(source):
                self.add_to_zip(source, basename, zipf)
            else:
                for root, dirs, files in os.walk(source):
                    for file in files:
                        file_path: str = os.path.join(root, file)
                        rel_path = os.path.relpath(file_path, source)
                        arc_name = os.path.join(basename, rel_path)
                        self.add_to_zip(file_path, arc_name, zipf)

    def add_to_zip(self, file_path, arcname, zipf):
        try:
            zipf.write(file_path, arcname)

        except PermissionError:
            print(f"Пропущен заблокированный файл: {file_path}")

        except Exception as e:
            print(f"Ошибка при добавлении файла: {file_path} - {str(e)}")

    def copy_with_errors(self, src, dst):
        try:
            shutil.copy2(src, dst)
        except PermissionError:
            print(f"Пропущен заблокированный файл: {src}")
        except Exception as e:
            print(f"Ошибка копирования: {src} - {str(e)}")

    def load_tasks(self):
        try:
            with open("tasks.json") as f:
                data = json.load(f)
                for item in data:
                    task = BackupTask(
                        name=item['name'],
                        source=item['source'],
                        destination=item['destination'],
                        compression=item['compression'],
                        frequency=item['frequency'],
                        time_params=item['time_params']
                    )
                    task.last_run = datetime.fromisoformat(item['last_run']) if item['last_run'] else None
                    task.next_run = task.calculate_next_run()
                    self.tasks.append(task)
        except FileNotFoundError:
            pass

    def save_tasks(self):
        data = [task.to_dict() for task in self.tasks]
        with open("tasks.json", "w") as f:
            json.dump(data, f, indent=2)

    def load_config(self):
        try:
            with open("config.json") as f:
                self.config.update(json.load(f))
        except FileNotFoundError:
            pass

    def save_config(self):
        with open("config.json", "w") as f:
            json.dump(self.config, f, indent=2)


# Диалог редактирования задачи
class TaskDialog(tk.Toplevel):
    def __init__(self, parent, task=None):
        super().__init__(parent)
        self.parent = parent
        self.task = task
        self.result = None

        try:
            self.iconbitmap(ICON_PATH)  # Для окна
        except Exception as e:
            print(f"Ошибка загрузки иконки: {str(e)}")

        self.title("Редактирование задачи" if task else "Новая задача")
        self.create_widgets()
        self.grab_set()

    def create_widgets(self):
        for widget in self.winfo_children():
            widget.destroy()

        ttk.Label(self, text="Имя:").grid(row=0, column=0, sticky=tk.W)
        self.name_entry = ttk.Entry(self)
        self.name_entry.grid(row=0, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(self, text="Источник:").grid(row=1, column=0, sticky=tk.W)
        self.source_entry = ttk.Entry(self)
        self.source_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        ttk.Button(self, text="Обзор", command=self.browse_source).grid(row=1, column=2, padx=5, pady=2)

        ttk.Label(self, text="Назначение:").grid(row=2, column=0, sticky=tk.W)
        self.dest_entry = ttk.Entry(self)
        self.dest_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)
        ttk.Button(self, text="Обзор", command=self.browse_dest).grid(row=2, column=2, padx=5, pady=2)

        self.compression_var = tk.BooleanVar()
        ttk.Checkbutton(self, text="Сжатие", variable=self.compression_var).grid(row=3, column=0, columnspan=3,
                                                                                 sticky=tk.W)

        ttk.Label(self, text="Расписание:").grid(row=4, column=0, sticky=tk.W)
        self.frequency_var = tk.StringVar()
        self.frequency_combo = ttk.Combobox(self, textvariable=self.frequency_var,
                                            values=["hourly", "daily", "weekly", "monthly"])
        self.frequency_combo.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=2)
        self.frequency_combo.bind("<<ComboboxSelected>>", self.update_time_widgets)

        self.time_frame = ttk.Frame(self)
        self.time_frame.grid(row=5, column=0, columnspan=3, sticky=tk.EW)

        ttk.Button(self, text="OK", command=self.on_ok).grid(row=6, column=1, padx=5, pady=5)
        ttk.Button(self, text="Отмена", command=self.destroy).grid(row=6, column=2, padx=5, pady=5)

        self.grid_columnconfigure(1, weight=1)

        if self.task:
            self.name_entry.insert(0, self.task.name)
            self.source_entry.insert(0, self.task.source)
            self.dest_entry.insert(0, self.task.destination)
            self.compression_var.set(self.task.compression)
            self.frequency_var.set(self.task.frequency)
            self.update_time_widgets()

            if self.task.frequency == "hourly":
                self.minute_spinbox.set(self.task.time_params)
            elif self.task.frequency == "daily":
                self.hour_spinbox.set(self.task.time_params[0])
                self.minute_spinbox.set(self.task.time_params[1])
            elif self.task.frequency == "weekly":
                self.weekday_combo.current(self.task.time_params[0])
                self.hour_spinbox.set(self.task.time_params[1])
                self.minute_spinbox.set(self.task.time_params[2])
            elif self.task.frequency == "monthly":
                self.day_spinbox.set(self.task.time_params[0])
                self.hour_spinbox.set(self.task.time_params[1])
                self.minute_spinbox.set(self.task.time_params[2])

    def browse_source(self):
        path = filedialog.askdirectory()
        if path:
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, path)
            logging.info(f"Выбран источник: {path}")

    def browse_dest(self):
        path = filedialog.askdirectory()
        if path:
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, path)
            logging.info(f"Выбрана целевая директория: {path}")

    def update_time_widgets(self, event=None):
        for widget in self.time_frame.winfo_children():
            widget.destroy()

        frequency = self.frequency_var.get()

        if frequency == "hourly":
            ttk.Label(self.time_frame, text="Минута:").pack(side=tk.LEFT)
            self.minute_spinbox = ttk.Spinbox(self.time_frame, from_=0, to=59, width=5)
            self.minute_spinbox.pack(side=tk.LEFT)

        elif frequency == "daily":
            ttk.Label(self.time_frame, text="Время:").pack(side=tk.LEFT)
            self.hour_spinbox = ttk.Spinbox(self.time_frame, from_=0, to=23, width=5)
            self.hour_spinbox.pack(side=tk.LEFT)
            ttk.Label(self.time_frame, text=":").pack(side=tk.LEFT)
            self.minute_spinbox = ttk.Spinbox(self.time_frame, from_=0, to=59, width=5)
            self.minute_spinbox.pack(side=tk.LEFT)

        elif frequency == "weekly":
            ttk.Label(self.time_frame, text="День недели:").pack(side=tk.LEFT)
            self.weekday_combo = ttk.Combobox(self.time_frame, values=["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"])
            self.weekday_combo.pack(side=tk.LEFT)
            ttk.Label(self.time_frame, text="Время:").pack(side=tk.LEFT)
            self.hour_spinbox = ttk.Spinbox(self.time_frame, from_=0, to=23, width=5)
            self.hour_spinbox.pack(side=tk.LEFT)
            ttk.Label(self.time_frame, text=":").pack(side=tk.LEFT)
            self.minute_spinbox = ttk.Spinbox(self.time_frame, from_=0, to=59, width=5)
            self.minute_spinbox.pack(side=tk.LEFT)

        elif frequency == "monthly":
            ttk.Label(self.time_frame, text="День месяца:").pack(side=tk.LEFT)
            self.day_spinbox = ttk.Spinbox(self.time_frame, from_=1, to=31, width=5)
            self.day_spinbox.pack(side=tk.LEFT)
            ttk.Label(self.time_frame, text="Время:").pack(side=tk.LEFT)
            self.hour_spinbox = ttk.Spinbox(self.time_frame, from_=0, to=23, width=5)
            self.hour_spinbox.pack(side=tk.LEFT)
            ttk.Label(self.time_frame, text=":").pack(side=tk.LEFT)
            self.minute_spinbox = ttk.Spinbox(self.time_frame, from_=0, to=59, width=5)
            self.minute_spinbox.pack(side=tk.LEFT)

    def on_ok(self):
        try:
            name = self.name_entry.get()
            source = self.source_entry.get()
            dest = self.dest_entry.get()
            compression = self.compression_var.get()
            frequency = self.frequency_var.get()

            if not all([name, source, dest]):
                raise ValueError("Заполните все обязательные поля")

            time_params = []
            if frequency == "hourly":
                minute = int(self.minute_spinbox.get())
                if not 0 <= minute <= 59:
                    raise ValueError("Некорректное значение минут")
                time_params = minute

            elif frequency == "daily":
                hour = int(self.hour_spinbox.get())
                minute = int(self.minute_spinbox.get())
                if not 0 <= hour <= 23 or not 0 <= minute <= 59:
                    raise ValueError("Некорректное время")
                time_params = (hour, minute)

            elif frequency == "weekly":
                weekday = self.weekday_combo.current()
                hour = int(self.hour_spinbox.get())
                minute = int(self.minute_spinbox.get())
                if weekday < 0 or not 0 <= hour <= 23 or not 0 <= minute <= 59:
                    raise ValueError("Некорректные параметры")
                time_params = (weekday, hour, minute)

            elif frequency == "monthly":
                day = int(self.day_spinbox.get())
                hour = int(self.hour_spinbox.get())
                minute = int(self.minute_spinbox.get())
                if not 1 <= day <= 31 or not 0 <= hour <= 23 or not 0 <= minute <= 59:
                    raise ValueError("Некорректные параметры")
                time_params = (day, hour, minute)

            else:
                raise ValueError("Неизвестная частота выполнения")

            self.result = BackupTask(
                name=name,
                source=source,
                destination=dest,
                compression=compression,
                frequency=frequency,
                time_params=time_params
            )

            self.destroy()

        except Exception as e:
            logging.error(e)
            #messagebox.showerror("Ошибка", str(e))


# Диалог настроек
class SettingsDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Настройки")

        main_frame = ttk.Frame(self)
        main_frame.pack(padx=10, pady=10)

        try:
            self.iconbitmap(ICON_PATH)  # Для окна
        except Exception as e:
            print(f"Ошибка загрузки иконки: {str(e)}")

        # Настройки сжатия
        ttk.Label(main_frame, text="Уровень сжатия (0-9):").grid(row=0, column=0, sticky=tk.W)
        self.compression_level = ttk.Spinbox(main_frame, from_=0, to=9, width=5)
        self.compression_level.set(self.parent.config['compression_level'])
        self.compression_level.grid(row=0, column=1, padx=5, pady=5)

        ttk.Button(main_frame, text="Сохранить", command=self.save).grid(row=2, columnspan=2, pady=10)

    def save(self):
        try:
            level = int(self.compression_level.get())
            if 0 <= level <= 9:
                self.parent.config['compression_level'] = level
                self.parent.save_config()
                self.destroy()
            else:
                raise ValueError("Уровень сжатия должен быть от 0 до 9")
        except Exception as e:
            logging.error(e)


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(os.path.join(os.getcwd(), "backup.log")),
            logging.StreamHandler()
        ]
    )
    logging.info("Приложение запущено")
    app = BackupApp()
    app.mainloop()

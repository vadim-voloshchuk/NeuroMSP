# automation/lab1.py
#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/lab1.py • Автоматизация ЛР-1: Планирование задач проекта
"""
import sys
import time
import logging
from pathlib import Path
import yaml
from win32com.client import Dispatch
from .base import safe_step
from .utils import ensure_dir, try_call, focus, shot, save_as

# ─── Конфигурация ──────────────────────────────────────
CONFIG_FILE  = Path(__file__).parent.parent / 'configs' / 'variants.yaml'
VARIANT      = None  # задаётся через sys.argv
PROJECT_FILE = None  # будет прочитано из конфига
SAVE_DIR     = Path('outputs') / 'lab1'
SCREEN_DIR   = Path('screenshots') / 'lab1'
LOG_FILE     = Path('logs') / 'lab1.log'
PAUSE        = 0.7  # пауза между шагами
# ───────────────────────────────────────────────────────

# Загрузка параметров по варианту
if len(sys.argv) < 2:
    print("Usage: lab1.py <variant>")
    sys.exit(1)
VARIANT = sys.argv[1]
with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
    cfg_all = yaml.safe_load(f)
if VARIANT not in cfg_all:
    print(f"Variant '{VARIANT}' not found in config")
    sys.exit(1)
cfg = cfg_all[VARIANT]
PROJECT_FILE = cfg['project_file']
BASE_CALENDAR = cfg.get('base_calendar_index', 1)

# Настройка логирования
logging.basicConfig(
    filename=str(LOG_FILE), level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8"
)

@safe_step("Лаба1_01_Open_View.mpp")
def step01(app):
    # Шаг 1: Открыть проект и настроить панель представлений
    if not try_call(app.ViewApply, *cfg.get('view_names', [])):
        logging.info("Manual: настройка представлений")
        raise AttributeError("View setup manual")

@safe_step()
def step02(app):
    # Шаг 2: Сохранить проект (Файл→Сохранить)
    # safe_step без аргументов не сохраняет отдельный файл
    pass

@safe_step("Лаба1_03_SetStart.mpp")
def step03(app):
    # Шаг 3: Установить дату начала проекта
    proj = app.ActiveProject
    proj.ProjectStart = cfg['project_start']

@safe_step("Лаба1_04_CalendarExceptions.mpp")
def step04(app):
    # Шаг 4: Добавить исключения в календарь
    cal = app.BaseCalendars(BASE_CALENDAR)
    for ex in cfg.get('exceptions', []):
        exc = cal.CalendarExceptions.Add(
            ex['name'], ex['start'], ex.get('finish', ex['start'])
        )

@safe_step("Лаба1_05_TasksSetup.mpp")
def step05(app):
    # Шаг 5: Ввести задачи, длительности и связи из конфига
    proj = app.ActiveProject
    minutes_per_day = proj.HoursPerDay * 60
    # Добавление задач
    for entry in cfg['tasks']:
        t = proj.Tasks.Add(entry['name'])
        # Установка длительности
        t.Duration = entry['duration'] * minutes_per_day
        # Признак вехи
        if entry['duration'] == 0:
            t.Milestone = True
    # Установка связей
    for idx, entry in enumerate(cfg['tasks'], start=1):
        preds = entry.get('predecessors')
        if preds:
            task = proj.Tasks(idx)
            task.Predecessors = ";".join(str(p) for p in preds)
    # Установка ограничений
    for constr in cfg.get('constraints', []):
        t = proj.Tasks(constr['task_id'])
        t.ConstraintType = constr['type']
        t.ConstraintDate = constr['date']


def run():
    ensure_dir(SAVE_DIR)
    ensure_dir(SCREEN_DIR)
    app = Dispatch("MSProject.Application")
    app.Visible = True

    if not Path(PROJECT_FILE).exists():
        print(f"Файл {PROJECT_FILE} не найден, создаю новый проект...")
        app.FileNew()
        app.FileSaveAs(PROJECT_FILE)
    else:
        app.FileOpen(PROJECT_FILE)

    steps = [step01, step02, step03, step04, step05]
    for idx, fn in enumerate(steps, 1):
        fn(app, idx, fn.__name__)
        time.sleep(PAUSE)
    print(f"Done Lab1 variant {VARIANT}")


if __name__ == "__main__":
    run()

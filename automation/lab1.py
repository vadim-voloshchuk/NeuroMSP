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
CONFIG_FILE   = Path(__file__).parent.parent / 'configs' / 'variants.yaml'
PROJECTS_DIR  = Path(__file__).parent.parent / 'projects'
SAVE_DIR      = Path('outputs') / 'lab1'
SCREEN_DIR    = Path('screenshots') / 'lab1'
LOG_FILE      = Path('logs') / 'lab1.log'
PAUSE         = 0.7
# ───────────────────────────────────────────────────────

if len(sys.argv) < 2:
    print("Usage: python -m automation.lab1 <variant>")
    sys.exit(1)
VARIANT = sys.argv[1]
with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
    cfg_all = yaml.safe_load(f)
if VARIANT not in cfg_all:
    print(f"Variant '{VARIANT}' not found in config")
    sys.exit(1)
cfg = cfg_all[VARIANT]

# путь к проекту: projects/<variant>.mpp
PROJECTS_DIR.mkdir(parents=True, exist_ok=True)
PROJECT_FILE = PROJECTS_DIR / f"{VARIANT}.mpp"
BASE_CALENDAR = cfg.get('base_calendar_index', 1)

logging.basicConfig(
    filename=str(LOG_FILE), level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8"
)

@safe_step("Лаба1_01_Open_View.mpp")
def step01(app):
    if not try_call(app.ViewApply, *cfg.get('view_names', [])):
        logging.info("Manual: настройка представлений")
        raise AttributeError("View setup manual")

@safe_step()  # просто сохранение
def step02(app):
    pass

@safe_step("Лаба1_03_SetStart.mpp")
def step03(app):
    proj = app.ActiveProject
    proj.ProjectStart = cfg['project_start']

@safe_step("Лаба1_04_CalendarExceptions.mpp")
def step04(app):
    cal = app.BaseCalendars(BASE_CALENDAR)
    for ex in cfg.get('exceptions', []):
        cal.CalendarExceptions.Add(ex['name'], ex['start'], ex.get('finish', ex['start']))

@safe_step("Лаба1_05_TasksSetup.mpp")
def step05(app):
    proj = app.ActiveProject
    minutes_per_day = proj.HoursPerDay * 60
    # добавляем задачи
    for entry in cfg['tasks']:
        t = proj.Tasks.Add(entry['name'])
        t.Duration = entry['duration'] * minutes_per_day
        if entry['duration'] == 0:
            t.Milestone = True

    # ставим предшественники в try/except
    for idx, entry in enumerate(cfg['tasks'], start=1):
        preds = entry.get('predecessors')
        if not preds:
            continue
        try:
            proj.Tasks(idx).Predecessors = ";".join(str(p) for p in preds)
        except Exception as e:
            logging.info("skip preds for task %s (%s): %s", idx, entry['name'], e)

    # ограничения
    for constr in cfg.get('constraints', []):
        t = proj.Tasks(constr['task_id'])
        t.ConstraintType = constr['type']
        t.ConstraintDate = constr['date']


def run():
    ensure_dir(PROJECTS_DIR)
    ensure_dir(SAVE_DIR)
    ensure_dir(SCREEN_DIR)

    app = Dispatch("MSProject.Application")
    app.Visible = True

    if not PROJECT_FILE.exists():
        print(f"Файл {PROJECT_FILE} не найден — создаю новый проект...")
        app.FileNew()
        # сохраняем сразу в projects/<variant>.mpp
        save_as(app, str(PROJECT_FILE))
    else:
        app.FileOpen(str(PROJECT_FILE))

    steps = [step01, step02, step03, step04, step05]
    for idx, fn in enumerate(steps, 1):
        fn(app, idx, fn.__name__)
        time.sleep(PAUSE)

    print(f"Done Lab1 variant {VARIANT}")

if __name__ == "__main__":
    run()

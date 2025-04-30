#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/lab1.py • Автоматизация ЛР-1: Планирование задач проекта
"""
import sys, time, logging
from pathlib import Path
import yaml
from win32com.client import Dispatch
from .base import safe_step
from .utils import ensure_dir, try_call, save_as, shot, focus

# ─── Пути ───────────────────────────────────────────────
ROOT       = Path(__file__).parent.parent.resolve()
CONFIG     = ROOT / 'configs' / 'variants.yaml'
PROJECTS   = ROOT / 'projects'
OUTPUTS    = ROOT / 'outputs' / 'lab1'
SCREENSHOTS= ROOT / 'screenshots' / 'lab1'
LOG_FILE   = ROOT / 'logs' / 'lab1.log'
PAUSE      = 0.7
# ───────────────────────────────────────────────────────

# аргумент — вариант
if len(sys.argv) < 2:
    print("Usage: python -m automation.lab1 <variant>")
    sys.exit(1)
variant = sys.argv[1]

# загружаем конфиг
with open(CONFIG, 'r', encoding='utf-8') as f:
    allcfg = yaml.safe_load(f)
if variant not in allcfg:
    print(f"Variant '{variant}' not in config")
    sys.exit(1)
cfg = allcfg[variant]

# готовим пути
PROJECTS.mkdir(parents=True, exist_ok=True)
OUTPUTS.mkdir(parents=True, exist_ok=True)
SCREENSHOTS.mkdir(parents=True, exist_ok=True)
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

PROJECT = PROJECTS / f"{variant}.mpp"
BASE_CAL = cfg.get('base_calendar_index', 1)

# логирование
logging.basicConfig(
    filename=str(LOG_FILE), level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8"
)

# ── ШАГ 1: Вид ─────────────────────────────────────────
@safe_step("Лаба1_01_Open_View.mpp")
def step01(app):
    if not try_call(app.ViewApply, *cfg.get('view_names', [])):
        raise AttributeError("View manual")

# ── ШАГ 2: Сохранение ───────────────────────────────────
@safe_step("Лаба1_02_SaveProject.mpp")
def step02(app):
    # просто фиксируем сохранение
    pass

# ── ШАГ 3: Дата старта ──────────────────────────────────
@safe_step("Лаба1_03_SetStart.mpp")
def step03(app):
    app.ActiveProject.ProjectStart = cfg['project_start']

# ── ШАГ 4: Исключения календаря ────────────────────────
@safe_step("Лаба1_04_CalendarExceptions.mpp")
def step04(app):
    cal = app.BaseCalendars(BASE_CAL)
    for ex in cfg.get('exceptions', []):
        cal.CalendarExceptions.Add(ex['name'], ex['start'], ex.get('finish', ex['start']))

# ── ШАГ 5: Задачи, связи, ограничения ──────────────────
@safe_step("Лаба1_05_TasksSetup.mpp")
def step05(app):
    proj = app.ActiveProject
    mins = proj.HoursPerDay * 60
    # добавляем
    for e in cfg['tasks']:
        t = proj.Tasks.Add(e['name'])
        t.Duration = e['duration'] * mins
        if e['duration'] == 0:
            t.Milestone = True
    # связи
    for i,e in enumerate(cfg['tasks'], start=1):
        preds = e.get('predecessors', [])
        if not preds: continue
        try:
            proj.Tasks(i).Predecessors = ";".join(str(p) for p in preds)
        except Exception as ex:
            logging.info("skip preds %d: %s", i, ex)
    # ограничения
    for c in cfg.get('constraints', []):
        t = proj.Tasks(c['task_id'])
        t.ConstraintType = c['type']
        t.ConstraintDate = c['date']

def run():
    app = Dispatch("MSProject.Application")
    app.Visible = True

    # создаём или открываем
    if not PROJECT.exists():
        print(f"→ Creating new project: {PROJECT}")
        app.FileNew()
        save_as(app, str(PROJECT))
    else:
        app.FileOpen(str(PROJECT))

    # выполняем шаги
    for idx, fn in enumerate([step01, step02, step03, step04, step05], 1):
        fn(app, idx, fn.__name__)
        time.sleep(PAUSE)

    save_as(app, str(PROJECT))

    print(f"✅ Done Lab1 variant {variant}")

if __name__=="__main__":
    run()

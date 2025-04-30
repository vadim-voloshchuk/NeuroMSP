#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/lab6.py • ЛР-6: отслеживание проекта
Запуск: python -m automation.lab6 <variant>
"""
import sys
import time
import logging
from pathlib import Path
from datetime import datetime
from win32com.client import Dispatch

from .base import safe_step
from .utils import ensure_dir, try_call, save_as

# ─── Пути ────────────────────────────────────────────────────────────────
ROOT       = Path(__file__).parent.parent.resolve()
PROJECTS   = ROOT / "projects"
OUTPUTS    = ROOT / "outputs" / "lab6"
SCREENS    = ROOT / "screenshots" / "lab6"
LOG_FILE   = ROOT / "logs" / "lab6.log"
PAUSE      = 0.7

# ─── Аргументы ───────────────────────────────────────────────────────────
if len(sys.argv) < 2:
    print("Usage: python -m automation.lab6 <variant>")
    sys.exit(1)
variant = sys.argv[1]
PROJECT = PROJECTS / f"{variant}.mpp"

# ─── Логирование ──────────────────────────────────────────────────────────
logging.basicConfig(
    filename=str(LOG_FILE),
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8"
)

# ─── ШАГ 1: Сохранение базового плана ─────────────────────────────────────
@safe_step("Лаба6_01_SetBaseline.mpp")
def s01_set_baseline(app):
    # 1. Диаграмма Ганта
    try_call(app.ViewApply, "Gantt Chart", "Диаграмма Ганта")
    # 2–3. Задать базовый план для всего проекта
    try:
        proj = app.ActiveProject
        # 0 = Baseline, True = all tasks
        proj.SetBaseline(0, True)
    except Exception:
        logging.info("Failed to set baseline via COM – use manual dialog")
    # 4. Вид: Гантт с отслеживанием
    try_call(app.ViewApply, "Tracking Gantt", "Диаграмма Ганта с отслеживанием")

# ─── ШАГ 2: Настройка вида «Использование задач» ──────────────────────────
@safe_step("Лаба6_02_UsageSetup.mpp")
def s02_usage_setup(app):
    # 1. Представление «Использование задач» / Task Usage
    try_call(app.ExecuteMso, "ViewTaskUsage")
    # 2. Правая таблица: Work, Actual Work, Actual Cost
    for col in ("Work", "Actual Work", "Actual Cost"):
        try_call(app.InsertColumn, col)
    # 3. Левая таблица: Stop, Resume, Remaining Work, % Complete
    for col in ("Stop", "Resume", "Remaining Work", "% Complete"):
        try_call(app.InsertColumn, col)

# ─── ШАГ 3: Ввод повременных данных ресурсов ───────────────────────────────
@safe_step("Лаба6_03_ResTimephased.mpp")
def s03_res_timephased(app):
    logging.info("Enter resource timephased data manually in Task Usage view")

# ─── ШАГ 4: Ввод повременных данных задач ─────────────────────────────────
@safe_step("Лаба6_04_TaskTimephased.mpp")
def s04_task_timephased(app):
    logging.info("Enter task timephased data manually in Task Usage view")

# ─── ШАГ 5: Ввод фактических трудозатрат ──────────────────────────────────
@safe_step("Лаба6_05_FactWorkInput.mpp")
def s05_fact_work(app):
    # 1. Вид: Использование задач
    try_call(app.ExecuteMso, "ViewTaskUsage")
    # 2. Вставка столбца «Actual Work»
    try_call(app.InsertColumn, "Actual Work")
    logging.info("Enter actual and remaining work manually for selected tasks")

# ─── ШАГ 6: Ввод % завершения в таблице ───────────────────────────────────
@safe_step("Лаба6_06_PercentComplete.mpp")
def s06_percent_complete(app):
    # Вид: Использование задач
    try_call(app.ExecuteMso, "ViewTaskUsage")
    # Вставка столбца % Complete
    try_call(app.InsertColumn, "% Complete")
    logging.info("Enter % Complete manually in Usage view and via Task dialog")

# ─── ШАГ 7: Анализ освоенного объёма ───────────────────────────────────────
@safe_step("Лаба6_07_EarnedValueAnalysis.mpp")
def s07_ev_analysis(app):
    # Вид: Использование задач
    try_call(app.ExecuteMso, "ViewTaskUsage")
    # Таблица: Earned Value
    try_call(app.TableApply, "Earned Value")
    # Суммарная задача проекта
    try:
        app.ActiveProject.ShowProjectSummaryTask = True
    except Exception:
        logging.info("Could not show project summary task")
    # Установить дату отчёта: 19.11.2009
    try:
        app.ActiveProject.StatusDate = datetime(2009, 11, 19)
    except Exception:
        logging.info("Could not set Status Date via COM")

# ─── ШАГ 8: Анализ календарных отклонений ─────────────────────────────────
@safe_step("Лаба6_08_CalendarVariances.mpp")
def s08_calendar_variances(app):
    # Вид: Диаграмма Ганта с отслеживанием
    try_call(app.ViewApply, "Tracking Gantt", "Диаграмма Ганта с отслеживанием")
    # Таблица «Calendar Variances (Earned Value)»
    try:
        app.TableApply,("Calendar Variances (Earned Value)",)
    except Exception:
        logging.info("Apply calendar variances table manually")

# ─── ШАГ 9: Анализ показателей затрат ─────────────────────────────────────
@safe_step("Лаба6_09_CostVariances.mpp")
def s09_cost_variances(app):
    # Вид: Диаграмма Ганта с отслеживанием
    try_call(app.ViewApply, "Tracking Gantt", "Диаграмма Ганта с отслеживанием")
    # Таблица «Cost Variances (Earned Value)»
    try:
        app.TableApply,("Cost Variances (Earned Value)",)
    except Exception:
        logging.info("Apply cost variances table manually")

# ─── MAIN ──────────────────────────────────────────────────────────────────
def run():
    for d in (PROJECTS, OUTPUTS, SCREENS):
        ensure_dir(d)

    app = Dispatch("MSProject.Application")
    app.Visible = True

    if PROJECT.exists():
        app.FileOpen(str(PROJECT))
    else:
        app.FileNew()
        save_as(app, str(PROJECT))

    steps = [
        s01_set_baseline, s02_usage_setup, s03_res_timephased,
        s04_task_timephased, s05_fact_work,    s06_percent_complete,
        s07_ev_analysis,   s08_calendar_variances, s09_cost_variances
    ]
    for idx, fn in enumerate(steps, 1):
        fn(app, idx, fn.__name__)
        time.sleep(PAUSE)

    print(f"✅ ЛР-6 завершена для варианта «{variant}»")

if __name__ == "__main__":
    run()

#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/lab2.py • Автоматизация ЛР-2: Таблицы и представления
"""
import sys
import time
import logging
from pathlib import Path
from win32com.client import Dispatch
from .base import safe_step
from .utils import ensure_dir, try_call, focus, shot, save_as

# ─── Пути и константы ──────────────────────────────────
ROOT        = Path(__file__).parent.parent.resolve()
PROJECTS    = ROOT / 'projects'
OUTPUTS     = ROOT / 'outputs' / 'lab2'
SCREENSHOTS = ROOT / 'screenshots' / 'lab2'
LOG_FILE    = ROOT / 'logs' / 'lab2.log'
PAUSE       = 0.7

# аргумент: variant (так же, как в lab1)
if len(sys.argv) < 2:
    print("Usage: python -m automation.lab2 <variant>")
    sys.exit(1)
variant = sys.argv[1]
PROJECT_FILE = PROJECTS / f"{variant}.mpp"

# логирование
logging.basicConfig(
    filename=str(LOG_FILE), level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8"
)

# ── Шаг 1: Форматирование входной таблицы ───────────────
@safe_step("Лаба2_01_FormatTable.mpp")
def step01(app):
    # Вид: Гант
    try_call(app.ViewApply, "Диаграмма Ганта", "Gantt Chart")
    # Удалить столбец ID
    if try_call(app.SelectColumn, "ID", "ИД", "№"):
        app.EditDelete()
    # Добавить столбец Критическая задача
    try_call(app.InsertColumn, "Критическая задача")
    # Заменить на столбец Затраты: вставляем и убираем лишний
    try_call(app.InsertColumn, "Затраты")
    # Стили текста
    try:
        ts = app.TextStyles
        ts.Item("Row and Column Titles").Font.Bold = True
        ts.Item("Summary Tasks").Font.Color = 7    # малиновый
        ts.Item("Middle Tier").Font.Color = 5      # тёмно-синий
    except Exception as e:
        logging.info("skip text styles: %s", e)

# ── Шаг 2: Сортировка по дате начала и окончания ───────
@safe_step("Лаба2_02_SortStartFinish.mpp")
def step02(app):
    try_call(app.Sort, "Start", True)
    try_call(app.Sort, "Finish", True)

# ── Шаг 3: Многоуровневая сортировка ───────────────────
@safe_step("Лаба2_03_MultiSort.mpp")
def step03(app):
    try:
        app.SortEx(
            SortFields=[
                {"FieldName": "Critical", "Descending": False},
                {"FieldName": "Finish",   "Descending": True}
            ],
            KeepOutlineStructure=False
        )
    except Exception as e:
        logging.info("skip multisort: %s", e)

# ── Шаг 4: Структурный фильтр Уровень 1 ────────────────
@safe_step("Лаба2_04_StructFilterLv1.mpp")
def step04(app):
    try_call(app.OutlineShowLevel, 1)

# ── Шаг 5: Автофильтр: начало следующего месяца + Dur>15 ──
@safe_step("Лаба2_05_AutoFilter.mpp")
def step05(app):
    try_call(app.AutoFilter, True)
    # фильтр Dur > 15
    try_call(app.FilterEdit, "Dur15", True, True,
             "Duration", "greater than", "15d", None, None, None, False)
    try_call(app.FilterApply, "Dur15")

# ── Шаг 6: Предопределённый фильтр: суммарные задачи ────
@safe_step("Лаба2_06_FilterSummary.mpp")
def step06(app):
    try_call(app.FilterApply, "Суммарные задачи", "Summary Tasks")

# ── Шаг 7: Пользовательский фильтр: критические ≤14 ─────
@safe_step("Лаба2_07_UserFilterCrit14.mpp")
def step07(app):
    try_call(app.FilterEdit, "Crit14", True, True,
             "Critical", "equals", "Yes",
             "And",
             "Duration", "less than or equal", "14d",
             None, None, False)
    try_call(app.FilterApply, "Crit14")

# ── Шаг 8: Предопределённая группировка: вехи ───────────
@safe_step("Лаба2_08_GroupMilestones.mpp")
def step08(app):
    try_call(app.GroupApply, "Вехи", "Milestones")

# ── Шаг 9: Пользовательская группировка: крит+длит ──────
@safe_step("Лаба2_09_GroupCritDur.mpp")
def step09(app):
    try:
        app.GroupEdit("CritDur", True, True,
                      "Critical", True, True,
                      "Duration", True, True, False)
        app.GroupApply("CritDur")
    except Exception as e:
        logging.info("skip custom grouping: %s", e)

# ── Шаг 10: Временная группировка по неделям ────────────
@safe_step("Лаба2_10_TimeGrouping.mpp")
def step10(app):
    try:
        app.GroupEdit("ByWeek", True, True,
                      "Duration", True, True,
                      Interval=1, IntervalUnit=3)
        app.GroupApply("ByWeek")
    except Exception as e:
        logging.info("skip time grouping: %s", e)

# ── Шаг 11: Формат одного отрезка ───────────────────────
@safe_step("Лаба2_11_FormatOneBar.mpp")
def step11(app):
    try:
        # пример: изменить первую задачу
        bar = app.BarStyles("Normal")
        bar.BarStartShape = 1
        bar.BarEndShape = 1
        bar.MiddleText = 1  # Duration
    except Exception as e:
        logging.info("skip format one bar: %s", e)

# ── Шаг 12: Формат всех обычных задач ───────────────────
@safe_step("Лаба2_12_FormatAllBars.mpp")
def step12(app):
    try_call(app.BarStyleEdit, "Normal", "Normal", 1, "Start", "Finish", 1, 1)

# ── Шаг 13: Критические задачи красным ───────────────────
@safe_step("Лаба2_13_CritRed.mpp")
def step13(app):
    try_call(app.BarStyleEdit, "Critical", "Critical", 1, "Start", "Finish", 1, 2)

# ── Шаг 14: Шкала → месяцы, цвет нерабочих → жёлтый ────
@safe_step("Лаба2_14_Timescale.mpp")
def step14(app):
    try
        app.TimescaleTopTier(4)  # months
        cal = app.BaseCalendars(1)
        exc = cal.CalendarExceptions.Add("tmp","01/01/2000","01/01/2000")
        exc.Color = 6
    except Exception as e:
        logging.info("skip timescale: %s", e)

# ── Шаг 15: Сетевой график — новая задача ────────────────
@safe_step("Лаба2_15_NetworkAdd.mpp")
def step15(app):
    try:
        t = app.ActiveProject.Tasks.Add("Новая задача")
        t.Predecessors = "1"
    except Exception as e:
        logging.info("skip network task: %s", e)

# ── MAIN ────────────────────────────────────────────────
def run():
    ensure_dir(PROJECTS)
    ensure_dir(OUTPUTS)
    ensure_dir(SCREENSHOTS)

    app = Dispatch("MSProject.Application")
    app.Visible = True

    if not PROJECT_FILE.exists():
        print(f"→ Creating project for Lab2: {PROJECT_FILE}")
        app.FileNew()
        save_as(app, str(PROJECT_FILE))
    else:
        app.FileOpen(str(PROJECT_FILE))

    steps = [step01, step02, step03, step04, step05,
             step06, step07, step08, step09, step10,
             step11, step12, step13, step14, step15]
    for idx, fn in enumerate(steps, 1):
        fn(app, idx, fn.__name__)
        time.sleep(PAUSE)

    print(f"✅ Done Lab2 variant {variant}")

if __name__ == "__main__":
    run()

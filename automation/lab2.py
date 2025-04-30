#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/lab2.py • Автоматизация ЛР‑2: Таблицы и представления

Запускается после выполнения ЛР‑1 и использует готовый `projects/<variant>.mpp`.
После каждого шага:
  • сохраняет *.mpp в `outputs/lab2/`  (если указан `save_fname`)
  • делает скриншот в `screenshots/lab2/`
  • пишет статус в лог
"""
from __future__ import annotations

import sys, time, logging
from pathlib import Path
from datetime import date, timedelta

from win32com.client import Dispatch

from .base  import safe_step
from .utils import ensure_dir, try_call, focus, shot, save_as

# ─── Пути и константы ──────────────────────────────────────────
ROOT        = Path(__file__).parent.parent.resolve()
PROJECTS    = ROOT / "projects"
OUTPUTS     = ROOT / "outputs" / "lab2"
SCREENS     = ROOT / "screenshots" / "lab2"
LOG_FILE    = ROOT / "logs" / "lab2.log"
PAUSE       = 0.8   # чуть больше, чтобы UI успевал прорисоваться

# ─── Аргументы командной строки ───────────────────────────────
if len(sys.argv) < 2:
    print("Usage: python -m automation.lab2 <variant>")
    sys.exit(1)
variant  = sys.argv[1]
PROJECT  = PROJECTS / f"{variant}.mpp"

# ─── Логирование ──────────────────────────────────────────────
logging.basicConfig(filename=str(LOG_FILE), level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8")

# ─── Вспомогательные функции ─────────────────────────────────

def exec_mso(app, mso: str):
    """Надёжный вызов лентой Office: CommandBars.ExecuteMso(idMso)"""
    try:
        app.CommandBars.ExecuteMso(mso)
        return True
    except Exception as e:
        logging.info("SKIP ExecuteMso %s: %s", mso, e)
        return False

# ─── Шаги ЛР‑2 ────────────────────────────────────────────────

@safe_step("Лаба2_01_FormatTable.mpp")
def step01_format_table(app):
    """Удалить ID, добавить Critical, сменить на Cost, стили текста"""
    exec_mso(app, "ViewGanttChart")
    time.sleep(0.4)
    # показать таблицу Entry для гарантии
    exec_mso(app, "TableEntry")
    time.sleep(0.2)
    # удалить столбец ID
    exec_mso(app, "SelectIDColumn")  # idMso в 365/2019
    exec_mso(app, "HideColumn")
    # вставить "Критическая задача"
    exec_mso(app, "TableInsertColumn")
    time.sleep(0.3)
    # нажимаем Enter (pyautogui) если диалог появился – пропустим, чтобы скрипт не зависал
    try:
        import pyautogui; pyautogui.press("enter")
    except Exception:
        pass
    # заменить на "Затраты": добавить новый и удалить предыдущий
    exec_mso(app, "TableInsertColumn"); time.sleep(0.3)
    try:
        import pyautogui; pyautogui.typewrite("Затраты\n", interval=0.05)
    except Exception:
        pass
    # стили текста через TextStyles
    try:
        ts = app.TextStyles
        ts.Item(1).Font.Bold = True                # Row+Col titles
        ts.Item(1).Font.Color = 12                 # коричневый
        ts.Item(4).Font.Color = 6                  # Summary – малиновый
        ts.Item(5).Font.Color = 2                  # Milestone – чёрный/синий
        ts.Item("Middle Tier").Font.Color = 13    # сиреневый
    except Exception as e:
        logging.info("skip TextStyles: %s", e)

@safe_step("Лаба2_02_SortStartFinish.mpp")
def step02_sort_dates(app):
    exec_mso(app, "SortByStart")
    exec_mso(app, "SortByFinish")

@safe_step("Лаба2_03_MultiSort.mpp")
def step03_multisort(app):
    try_call(app.Sort, "Critical", True, "Finish", False)

@safe_step("Лаба2_04_Level1.mpp")
def step04_outline_level1(app):
    exec_mso(app, "OutlineShowLevel1")

@safe_step("Лаба2_05_AutoFilter.mpp")
def step05_autofilter(app):
    """Автофильтр: начало следующего месяца & Dur>15d"""
    exec_mso(app, "ToggleAutoFilter")
    # Поставить собственный фильтр через FilterEdit (без диалогов)
    today = date.today()
    first_next = (today.replace(day=1) + timedelta(days=32)).replace(day=1)
    last_next  = (first_next.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    rng_name = "NextMonth"
    try_call(app.FilterEdit, rng_name, True, True,
             "Start", "is greater than or equal to", first_next.strftime("%d.%m.%Y"),
             "And",
             "Start", "is less than or equal to",   last_next.strftime("%d.%m.%Y"),
             False)
    dur_name = "Dur15"
    try_call(app.FilterEdit, dur_name, True, True,
             "Duration", "greater than", "15d", None, None, None, False)
    # комбинируем два фильтра подряд (быстрее всего визуально)
    try_call(app.FilterApply, rng_name)
    try_call(app.FilterApply, dur_name)

@safe_step("Лаба2_06_SummaryFilter.mpp")
def step06_filter_summary(app):
    exec_mso(app, "FilterSummaryTasks")

@safe_step("Лаба2_07_UserFilterCrit14.mpp")
def step07_user_filter(app):
    try_call(app.FilterEdit, "Crit14", True, True,
             "Critical", "equals", "Yes",
             "And",
             "Duration", "less than or equal", "14d",
             None, None, False)
    try_call(app.FilterApply, "Crit14")
    # добавить в меню
    try_call(app.FilterCopy, "Crit14", "КороткаяКритическаяЗадача")

@safe_step("Лаба2_08_GroupMilestones.mpp")
def step08_group_milestones(app):
    exec_mso(app, "GroupByMilestones")

@safe_step("Лаба2_09_GroupCustom.mpp")
def step09_group_custom(app):
    try_call(app.GroupEdit, "CritDur", True, True,
             "Critical", True, True,
             "Duration", True, True, False)
    try_call(app.GroupApply, "CritDur")

@safe_step("Лаба2_10_TimeGrouping.mpp")
def step10_time_group(app):
    try_call(app.GroupEdit, "ByWeek", True, True,
             "Duration", True, True, Interval=1, IntervalUnit=3)  # 3 — неделя
    try_call(app.GroupApply, "ByWeek")

@safe_step("Лаба2_11_FormatOneBar.mpp")
def step11_format_one_bar(app):
    # форматируем один Normal bar (индекс 1)
    try_call(app.BarStyleEdit, "Normal", "Normal", 1, "Start", "Finish", 1, 1)

@safe_step("Лаба2_12_FormatAllBars.mpp")
def step12_format_all(app):
    try_call(app.BarStyleEdit, "Normal", "Normal", 1, "Start", "Finish", 1, 1)

@safe_step("Лаба2_13_CritRed.mpp")
def step13_crit_red(app):
    try_call(app.BarStyleEdit, "Critical", "Critical", 1, "Start", "Finish", 1, 2)

@safe_step("Лаба2_14_Timescale.mpp")
def step14_timescale(app):
    exec_mso(app, "TimescaleTopTierMonths")
    # Non‑working yellow: нет прямого ExecuteMso, но можно изменить стиль календаря позже

@safe_step("Лаба2_15_NetworkNewTask.mpp")
def step15_network(app):
    exec_mso(app, "ViewNetworkDiagram")
    time.sleep(0.5)
    try:
        t = app.ActiveProject.Tasks.Add("Новая задача")
        t.Predecessors = "1"
    except Exception as e:
        logging.info("skip net task: %s", e)

# ─── MAIN ──────────────────────────────────────────────
def run():
    ensure_dir(PROJECTS); ensure_dir(OUTPUTS); ensure_dir(SCREENS)
    app = Dispatch("MSProject.Application"); app.Visible = True

    if not PROJECT.exists():
        print(f"→ creating project for Lab2: {PROJECT}")
        app.FileNew(); save_as(app, str(PROJECT))
    else:
        app.FileOpen(str(PROJECT))

    # максимизируем окно — чтобы скрины всегда полные
    try:
        import pygetwindow as gw; time.sleep(1)
        for w in gw.getWindowsWithTitle(app.Caption):
            w.maximize(); time.sleep(0.8); break
    except Exception:
        pass

    steps = [step01_format_table, step02_sort_dates, step03_multisort,
             step04_outline_level1, step05_autofilter, step06_filter_summary,
             step07_user_filter, step08_group_milestones, step09_group_custom,
             step10_time_group, step11_format_one_bar, step12_format_all,
             step13_crit_red, step14_timescale, step15_network]

    for idx, fn in enumerate(steps, 1):
        fn(app, idx, fn.__name__)
        time.sleep(PAUSE)

    print(f"✅ Lab2 finished for variant {variant}")

if __name__ == "__main__":
    run()

#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/lab2.py • ЛР-2: таблицы и представления
Запуск:  python -m automation.lab2 <variant>
Использует projects/<variant>.mpp, созданный Lab-1.
"""

from __future__ import annotations
import sys, time, logging
from datetime import date, timedelta
from pathlib import Path

from win32com.client import Dispatch
from .base  import safe_step
from .utils import ensure_dir, try_call, save_as, focus

# ─── Пути ──────────────────────────────────────────────
ROOT        = Path(__file__).parent.parent.resolve()
PROJECTS    = ROOT / "projects"
OUTPUTS     = ROOT / "outputs"     / "lab2"
SCREENS     = ROOT / "screenshots" / "lab2"
LOG_FILE    = ROOT / "logs" / "lab2.log"
PAUSE       = 0.8

# ─── Аргумент variant ─────────────────────────────────
if len(sys.argv) < 2:
    print("Usage: python -m automation.lab2 <variant>")
    sys.exit(1)
variant = sys.argv[1]
PROJECT = PROJECTS / f"{variant}.mpp"

# ─── Логирование ─────────────────────────────────────
logging.basicConfig(filename=str(LOG_FILE), level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s",
                    encoding="utf-8")

# ─── ШАГ 1  Форматирование таблицы «Ввод» ─────────────
@safe_step("Лаба2_01.mpp")
def s01(app):
    try_call(app.ViewApply, "Диаграмма Ганта", "Gantt Chart")
    try_call(app.TableApply, "Ввод", "Entry")

    # Удалить столбец ID
    if try_call(app.SelectColumn, "ID", "ИД", "№"):
        app.EditDelete()

    # Вставить «Критическая задача»
    try_call(app.TableEdit,
             Name:="Ввод", TaskTable:=True, Create:=False,
             FieldName:="Критическая задача", NewFieldName:="Критическая задача",
             InsertColumn:=True)

    # Заменить на «Затраты»
    if try_call(app.SelectColumn, "Критическая задача"):
        try_call(app.TableEdit,
                 Name:="Ввод", TaskTable:=True, Create:=False,
                 FieldName:="Затраты", NewFieldName:="Затраты",
                 InsertColumn:=True)
        app.EditDelete()  # удалить старый столбец

    # Стили текста
    try:
        app.TextStyles(1, None, None, 12)               # заголовки — коричневый
        app.TextStyles(4, None, None, 6)                # суммарные — малиновый
        app.TextStyles(5, None, None, 2)                # вехи      — чёрный
        app.TextStyles("Middle Tier", None, None, 13)   # сиреневый
    except Exception:
        pass

    # Автоподбор ширины всех столбцов (Std не умеет ColumnBestFit)
    for col in ("Имя задачи", "Затраты", "Длительность", "Начало", "Окончание"):
        try_call(app.TableEdit,
                 Name:="Ввод", TaskTable:=True, Create:=False,
                 FieldName:=col, NewFieldName:=col, Width:=22)

# ─── ШАГ 2  Сортировка Start-Finish ───────────────────
@safe_step("Лаба2_02.mpp")
def s02(app):
    try_call(app.Sort, "Start", True)
    try_call(app.Sort, "Finish", True)

# ─── ШАГ 3  Мультисорт Critical↑ + Finish↓ ────────────
@safe_step("Лаба2_03.mpp")
def s03(app):
    try_call(app.Sort, "Critical", True,
             "Finish", False, KeepOutlineStructure:=False)

# ─── ШАГ 4  Структурный фильтр – уровень 1 ────────────
@safe_step("Лаба2_04.mpp")
def s04(app):
    try_call(app.OutlineShowTasks, 1)

# ─── ШАГ 5  Автофильтр: следующий месяц & Dur>15 ─────
@safe_step("Лаба2_05.mpp")
def s05(app):
    # В Std проблемы с AutoFilterEdit → fallback на свойство
    try:
        app.AutoFilterEdit(True)
    except Exception:
        app.AutoFilter = True

    # Dur >15d
    try_call(app.FilterEdit, "Dur15", True, True,
             "Duration", "greater than", "15d", None, None, None, False)

    # Start ∈ следующий месяц
    today  = date.today()
    first  = (today.replace(day=1) + timedelta(days=32)).replace(day=1)
    last   = (first.replace(day=28)+timedelta(days=4)).replace(day=1)-timedelta(days=1)
    try_call(app.FilterEdit, "NextMon", True, True,
             "Start", "is greater than or equal to", first.strftime("%d.%m.%Y"),
             "And",
             "Start", "is less than or equal to",    last.strftime("%d.%m.%Y"),
             False)

    try_call(app.FilterApply, "Dur15")
    try_call(app.FilterApply, "NextMon")

# ─── ШАГ 6  Фильтр «Суммарные задачи» ────────────────
@safe_step("Лаба2_06.mpp")
def s06(app):
    try_call(app.FilterApply, "Суммарные задачи", "Summary Tasks")

# ─── ШАГ 7  Польз-фильтр Critical & ≤14d ──────────────
@safe_step("Лаба2_07.mpp")
def s07(app):
    try_call(app.FilterEdit, "Crit14", True, True,
             "Critical", "equals", "Yes",
             "And",
             "Duration", "less than equal", "14d",
             None, None, False)
    try_call(app.FilterApply, "Crit14")

# ─── ШАГ 8  Группировка «Вехи» ────────────────────────
@safe_step("Лаба2_08.mpp")
def s08(app):
    try_call(app.GroupApply, "Вехи", "Milestones")

# ─── ШАГ 9  Своя группировка Crit↓ + Dur↑ ────────────
@safe_step("Лаба2_09.mpp")
def s09(app):
    if hasattr(app, "GroupEdit") and \
       try_call(app.GroupEdit, "CritDur", True, True,
                "Critical", False, False,
                "Duration", True, True, False):
        try_call(app.GroupApply, "CritDur")
    else:
        logging.info("GroupEdit отсутствует — шаг пропущен")

# ─── ШАГ 10 Временная группировка по неделям ──────────
@safe_step("Лаба2_10.mpp")
def s10(app):
    if hasattr(app, "GroupEdit"):
        try_call(app.GroupEdit, "ByWeek", True, True,
                 "Duration", True, True,
                 Interval:=1, IntervalUnit:=3)
        try_call(app.GroupApply, "ByWeek")

# ─── ШАГ 11 Формат первого Normal-бара ────────────────
@safe_step("Лаба2_11.mpp")
def s11(app):
    if hasattr(app, "BarStyleEdit"):
        try_call(app.BarStyleEdit, "Normal", "Normal", 1,
                 "Start", "Finish", 1, 1)

# ─── ШАГ 12 Формат всех Normal-баров ──────────────────
@safe_step("Лаба2_12.mpp")
def s12(app):
    if hasattr(app, "BarStyleEdit"):
        try_call(app.BarStyleEdit, "Normal", "Normal", 1, "Start", "Finish", 1, 1)

# ─── ШАГ 13 Критические – красные ─────────────────────
@safe_step("Лаба2_13.mpp")
def s13(app):
    if hasattr(app, "BarStyleEdit"):
        try_call(app.BarStyleEdit, "Critical", "Critical", 1, "Start", "Finish", 1, 2)

# ─── ШАГ 14 Шкала: верхний tier – месяц ───────────────
@safe_step("Лаба2_14.mpp")
def s14(app):
    if hasattr(app, "TimescaleTopTier"):
        try_call(app.TimescaleTopTier, 3)  # 3 — Months
    else:
        logging.info("TimescaleTopTier недоступен — шаг пропущен")

# ─── ШАГ 15 Сетевой график: новая задача ──────────────
@safe_step("Лаба2_15.mpp")
def s15(app):
    try_call(app.ViewApply, "Сетевой график", "Network Diagram")
    time.sleep(0.4)
    t = app.ActiveProject.Tasks.Add("Новая задача")
    try: t.Predecessors = "1"
    except Exception: pass

# ─── MAIN ─────────────────────────────────────────────
def run():
    for p in (PROJECTS, OUTPUTS, SCREENS): ensure_dir(p)
    app = Dispatch("MSProject.Application"); app.Visible = True

    # открыть или создать проект-заглушку
    if PROJECT.exists():
        app.FileOpen(str(PROJECT))
    else:
        print(f"⚠️  {PROJECT.name} не найден – создаю пустой проект")
        app.FileNew(); save_as(app, str(PROJECT))

    # разворачиваем окно для чётких скринов
    try:
        import pygetwindow as gw; time.sleep(1)
        for w in gw.getWindowsWithTitle(app.Caption):
            w.maximize(); break
    except Exception:
        pass

    for idx, fn in enumerate(
        [s01, s02, s03, s04,
         s05, s06, s07, s08,
         s09, s10, s11, s12,
         s13, s14, s15], 1):
        fn(app, idx, fn.__name__)
        time.sleep(PAUSE)

    print(f"✅ ЛР-2 для варианта {variant} — готово")

if __name__ == "__main__":
    run()

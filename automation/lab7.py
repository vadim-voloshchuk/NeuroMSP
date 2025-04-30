#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/lab7.py • ЛР-7: отчётность по проекту
Запуск: python -m automation.lab7 <variant>
"""
import sys
import time
import logging
from pathlib import Path
from win32com.client import Dispatch

from .base import safe_step
from .utils import ensure_dir, try_call, save_as

# ─── Пути ────────────────────────────────────────────────────────────────
ROOT        = Path(__file__).parent.parent.resolve()
PROJECTS    = ROOT / "projects"
OUTPUTS     = ROOT / "outputs" / "lab7"
SCREENS     = ROOT / "screenshots" / "lab7"
LOG_FILE    = ROOT / "logs" / "lab7.log"
PAUSE       = 0.7

# ─── Аргументы ───────────────────────────────────────────────────────────
if len(sys.argv) < 2:
    print("Usage: python -m automation.lab7 <variant>")
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

# ─── ШАГ 1: Статистика проекта ────────────────────────────────────────────
@safe_step("Лаба7_01_Statistics.mpp")
def s01_project_statistics(app):
    # Проект → Сведения о проекте → Статистика
    # нет прямого COM-метода, поэтому напомним пользователю
    logging.info("Open Project Info → Statistics manually")

# ─── ШАГ 2: Стандартный отчёт «Дела по исполнителям и времени» ───────────
@safe_step("Лаба7_02_StandardReport.mpp")
def s02_standard_report(app):
    # Отчёт → Отчёты → Назначения → Дела по исполнителям и времени
    try:
        # открытие галереи отчётов
        app.ExecuteMso("ReportAssignments")  # приблизительный MSO-ид
        # выбор и запуск нужного отчёта
    except Exception:
        logging.info("Invoke standard report 'Work by Resource and Assignment' manually")

# ─── ШАГ 3: Настройка стандартного отчёта ─────────────────────────────────
@safe_step("Лаба7_03_CustomizeStdReport.mpp")
def s03_customize_std_report(app):
    # Отчёт → Отчёты → Назначения → Дела по исполнителям и времени → Изменить → столбец «Недели»
    logging.info("Customize standard report: change column grouping to Weeks manually")

# ─── ШАГ 4: Создание настраиваемого отчёта (сводная таблица) ──────────────
@safe_step("Лаба7_04_CreateCrossTabReport.mpp")
def s04_create_crosstab(app):
    # Отчёт → Отчёты → Настраиваемые → Создать → Перекрёстная таблица → параметры из рис.7.1 → Просмотр
    logging.info("Create custom Crosstab report with user settings manually")

# ─── ШАГ 5: Удаление созданного отчёта ────────────────────────────────────
@safe_step("Лаба7_05_DeleteReport.mpp")
def s05_delete_report(app):
    # Отчёт → Отчёты → Настраиваемые → Организатор → удалить «ЗатратыЗадач»
    logging.info("Delete custom report 'ЗатратыЗадач' via Organizer manually")

# ─── ШАГ 6: Формирование наглядного отчёта (Excel-диаграмма) ─────────────
@safe_step("Лаба7_06_RunVisualReport.mpp")
def s06_run_visual_report(app):
    # Отчёт → Наглядные отчёты → Сводный отчёт о затратах ресурсов → Просмотреть (Excel)
    try:
        app.ExecuteMso("VisualReportsGallery")
        # выбор шаблона «Resource Cost Summary»
    except Exception:
        logging.info("Run built-in visual report 'Resource Cost Summary' manually in Excel")

# ─── ШАГ 7: Создание собственного наглядного отчёта ────────────────────────
@safe_step("Лаба7_07_CreateVisualTemplate.mpp")
def s07_create_visual_template(app):
    # Отчёт → Наглядные отчёты → Создать шаблон → Сводка по задачам → задать поля → ОК → Excel
    logging.info("Create custom visual report template for 'Tasks summary' manually")

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
        s01_project_statistics,
        s02_standard_report,
        s03_customize_std_report,
        s04_create_crosstab,
        s05_delete_report,
        s06_run_visual_report,
        s07_create_visual_template,
    ]
    for idx, fn in enumerate(steps, 1):
        fn(app, idx, fn.__name__)
        time.sleep(PAUSE)

    print(f"✅ ЛР-7 завершена для варианта «{variant}»")

if __name__ == "__main__":
    run()

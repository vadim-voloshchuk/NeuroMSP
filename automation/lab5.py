#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/lab5.py • ЛР-5: выравнивание загрузки ресурсов
Запуск: python -m automation.lab5 <variant>
"""
import sys
import time
import logging
from pathlib import Path
from win32com.client import Dispatch

from .base import safe_step
from .utils import ensure_dir, try_call, save_as

# ─── Пути ────────────────────────────────────────────────────────────────
ROOT       = Path(__file__).parent.parent.resolve()
PROJECTS   = ROOT / "projects"
OUTPUTS    = ROOT / "outputs" / "lab5"
SCREENS    = ROOT / "screenshots" / "lab5"
LOG_FILE   = ROOT / "logs" / "lab5.log"
PAUSE      = 0.7

# ─── Хелпер для надёжного поиска ресурса ────────────────────────────────────
def res_by_name(pj, name: str):
    """Возвращает объект Resource с заданным Name или None."""
    return next((r for r in pj.Resources if r and r.Name == name), None)

# ─── Аргумент — вариант ────────────────────────────────────────────────────
if len(sys.argv) < 2:
    print("Usage: python -m automation.lab5 <variant>")
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

# ─── ШАГ 1: Автоматическое выравнивание загрузки ───────────────────────────
@safe_step("Лаба5_1.mpp")
def s01_auto_level(app):
    # Команда «Уровень ресурсов» (Level Resources)
    try_call(app.ExecuteMso, "LevelResource")
    # Комментарий, если нужно: logging.info("Auto‐leveling executed")

# ─── ШАГ 2: Параллель → последовательные ───────────────────────────────────
@safe_step("Лаба5_2.mpp")
def s02_convert_parallel(app):
    # Поскольку автоматического метода нет, пользователь должен вручную
    logging.info("Convert parallel tasks to sequential manually")

# ─── ШАГ 3: Замена ресурса «Техник» на «Администратор» ─────────────────────
@safe_step("Лаба5_3.mpp")
def s03_replace_res(app):
    pj    = app.ActiveProject
    tech  = res_by_name(pj, "Техник")
    admin = res_by_name(pj, "Администратор")

    if not tech:
        raise AttributeError("Resource «Техник» not found – replace manually")
    if not admin:
        raise AttributeError("Resource «Администратор» not found")

    # Проходим по всем заданиям и заменяем
    for task in pj.Tasks:
        if not task:
            continue
        for asn in task.Assignments:
            if asn and asn.ResourceID == tech.ID:
                asn.ResourceID = admin.ID

# ─── ШАГ 4: Ручное редактирование контура загрузки ─────────────────────────
@safe_step("Лаба5_4.mpp")
def s04_manual_contour(app):
    logging.info("Manual work‐contour editing required")

# ─── ШАГ 5: Перенос части трудозатрат в сверхурочные ───────────────────────
@safe_step("Лаба5_5.mpp")
def s05_overtime(app):
    # Переключаемся в представление «Использование ресурсов»
    try_call(app.ExecuteMso, "ViewResourceUsage")
    logging.info("Convert part of work to overtime manually")

# ─── MAIN ──────────────────────────────────────────────────────────────────
def run():
    # Убеждаемся, что все папки существуют
    for d in (PROJECTS, OUTPUTS, SCREENS):
        ensure_dir(d)

    app = Dispatch("MSProject.Application")
    app.Visible = True

    # Создаём новый или открываем существующий проект
    if PROJECT.exists():
        app.FileOpen(str(PROJECT))
    else:
        app.FileNew()
        save_as(app, str(PROJECT))

    # Последовательно выполняем шаги
    steps = [s01_auto_level, s02_convert_parallel, s03_replace_res,
             s04_manual_contour, s05_overtime]
    for idx, fn in enumerate(steps, 1):
        fn(app, idx, fn.__name__)
        time.sleep(PAUSE)

    print(f"✅ ЛР-5 завершена для варианта «{variant}»")

if __name__ == "__main__":
    run()

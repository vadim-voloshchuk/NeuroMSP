#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/lab5.py • ЛР-5: выравнивание загрузки ресурсов
Запуск:  python -m automation.lab5 <variant>
Работает с projects/<variant>.mpp (после ЛР-3)
"""

from __future__ import annotations
import sys, time, logging
from pathlib import Path
from win32com.client import Dispatch
from .base  import safe_step
from .utils import ensure_dir, try_call, save_as

ROOT   = Path(__file__).parent.parent.resolve()
PROJ   = ROOT / "projects"
OUT    = ROOT / "outputs"     / "lab5"
SHOT   = ROOT / "screenshots" / "lab5"
LOG    = ROOT / "logs" / "lab5.log"
PAUSE  = 0.7

if len(sys.argv) < 2:
    print("Usage: python -m automation.lab5 <variant>"); sys.exit(1)
variant  = sys.argv[1]
PROJECT  = PROJ / f"{variant}.mpp"

for p in (PROJ, OUT, SHOT): ensure_dir(p)
logging.basicConfig(filename=str(LOG),
                    level=logging.INFO,
                    format="%(asctime)s  %(levelname)s  %(message)s",
                    encoding="utf-8")

# ────────────────────────────── ШАГИ ЛР-5 ────────────────────────────────
@safe_step("Лаба5_1.mpp")
def s01_auto_level(app):
    """
    1. Автоматическое выравнивание.
    Параметры: только внутри проекта, день-за-днём, вычислять задержки.
    """
    pj = app.ActiveProject
    # если ранее выравнивали – сброс
    app.LevelingClearAll()
    # команда соответствует ленточной LevelResources
    try_call(app.LevelResources, pj.Name, True, True, True, False, False,
             2, True, True, 0, 0, False)    # key arguments → см. VBA-help
    # если ресурсов много, UI бывает занят – даём время
    time.sleep(2)

@safe_step("Лаба5_2.mpp")
def s02_parallel_to_seq(app):
    """
    2. превращаем выбранные параллельные фазы в последовательные.
       Здесь нужно ручное вмешательство: скрипт лишь сообщает.
    """
    raise AttributeError("convert parallel → sequential manually")

@safe_step("Лаба5_3.mpp")
def s03_replace_res(app):
    """
    3. Замена перегруженных ресурсов на менее загруженных.
       Автоматизируем частично: ищем ресурсы, где PeakUnits>1
       и пытаемся заменить на 'Техник', если он есть и свободен.
    """
    pj = app.ActiveProject
    tech = pj.Resources("Техник")
    if not tech:
        raise AttributeError("no spare resource, do manually")

    for t in pj.Tasks:
        if not t: continue
        for a in t.Assignments:
            if a and a.PeakUnits > 1 and tech.ID not in [x.ResourceID for x in t.Assignments]:
                try:
                    a.ResourceID = tech.ID
                except Exception:
                    logging.info("cannot replace %s in task %s", a.ResourceName, t.Name)
    # небольшой lag, потом сохранить

@safe_step("Лаба5_4.mpp")
def s04_manual_contour(app):
    """
    4. Ручное редактирование трудозатрат (контур Work → Flat).
       Отмечаем как SKIP — вы делаете вручную.
    """
    raise AttributeError("manual work contour editing")

@safe_step("Лаба5_5.mpp")
def s05_overtime(app):
    """
    5. Перенос части трудозатрат в сверхурочные.
       Скрипт находит все назначения >8 ч/день и переносит 2 ч в OvertimeWork.
    """
    for a in app.ActiveProject.Assignments:
        if not a: continue
        # проверяем среднее по длитель­ности
        dur_days = a.Duration / 480.0 if a.Duration else 1  # min/480 = день
        if dur_days == 0: continue
        avg = a.Work / dur_days
        if avg > 480:                 # >8 ч
            try:
                # переносим 2 ч = 120 min сутки в overtime
                a.OvertimeWork += 120 * dur_days
                a.Work         -= 120 * dur_days
            except Exception as e:
                logging.info("skip overtime on %s → %s", a.TaskName, e)

# ─────────────────────────────── MAIN ────────────────────────────────────
def run():
    app = Dispatch("MSProject.Application"); app.Visible = True
    if PROJECT.exists(): app.FileOpen(str(PROJECT))
    else:                app.FileNew(); save_as(app, str(PROJECT))

    for i, fn in enumerate((s01_auto_level,
                            s02_parallel_to_seq,
                            s03_replace_res,
                            s04_manual_contour,
                            s05_overtime), 1):
        fn(app, i, fn.__name__)
        time.sleep(PAUSE)

    print(f"✅ ЛР-5 завершена для варианта «{variant}»")

if __name__ == "__main__":
    run()

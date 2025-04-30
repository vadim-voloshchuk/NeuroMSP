#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/lab4.py • ЛР-4: анализ проекта
Запуск:  python -m automation.lab4 <variant>
"""

from __future__ import annotations
import sys, time, logging
from pathlib import Path
from win32com.client import Dispatch

from .base  import safe_step
from .utils import ensure_dir, try_call, save_as

# ──────────────────── пути / подготовка ──────────────────────────────────
ROOT     = Path(__file__).parent.parent.resolve()
PROJDIR  = ROOT / "projects"
OUTDIR   = ROOT / "outputs" / "lab4"
SHOTDIR  = ROOT / "screenshots" / "lab4"
LOG_F    = ROOT / "logs" / "lab4.log"
PAUSE    = .7

if len(sys.argv) < 2:
    print("Usage: python -m automation.lab4 <variant>"); sys.exit(1)
variant = sys.argv[1]
PROJECT = PROJDIR / f"{variant}.mpp"

for p in (PROJDIR, OUTDIR, SHOTDIR, LOG_F.parent): ensure_dir(p)

logging.basicConfig(filename=str(LOG_F), level=logging.INFO,
                    format="%(asctime)s  %(levelname)s  %(message)s",
                    encoding="utf-8")

# ─────────────────── вспомогательные «шоткаты» ───────────────────────────
def _save(app, fname):
    """сохраняет активный проект в outputs/lab4/<fname>"""
    save_as(app, str(OUTDIR / fname))

# ────────────────── ШАГ 1. Параметрический анализ ─────────────────────────
@safe_step("Лаба4_ПараметрическийАнализ.mpp")
def s01_parametric(app):
    """
    • создаём три настраиваемых поля:
        Number1 = Параметр
        Number2 = Норма
        Number3 = Оценка
      а Number3 = [Number1]*[Number2].
    • Пользователь должен сам пометить задачи и ввести значения →
      поэтому если что-то не получается, падаем в SKIP.
    """
    try:
        app.CustomFieldRename(4, "Параметр")   # 4 → pjCustomTaskNumber1
        app.CustomFieldRename(5, "Норма")      # 5 → …Number2
        app.CustomFieldRename(6, "Оценка")     # 6 → …Number3
        formula = "[Number1]*[Number2]"
        app.CustomFieldSetFormula(6, formula)
    except Exception as e:
        logging.info("parametric-field setup manual: %s", e)
        raise AttributeError("manual fill")

# ────────────────── ШАГ 2. PERT-анализ ───────────────────────────────────
@safe_step("Лаба4_PERTАнализ.mpp")
def s02_pert(app):
    """
    В MS Project 2016+ есть команда Calculate PERT: задаёт поля
    Оптимистическая / Ожидаемая / Пессимистическая (Duration1-3)
    и вычисляет Duration. Попробуем вызвать; если команды нет →
    вручную.
    """
    if not try_call(app.ExecuteMso, "PERTAnalysisCalculate"):
        raise AttributeError("PERT manual")

# ────────────────── ШАГ 3. Критический путь ───────────────────────────────
@safe_step("Лаба4_АнализКрит.mpp")
def s03_crit(app):
    try_call(app.ViewApply, "График Ганта", "Gantt Chart")
    try_call(app.FilterApply, "Critical", "Критические задачи")

# ────────────────── ШАГ 4. Стоимость проекта ──────────────────────────────
@safe_step("Лаба4_АнализСтоим.mpp")
def s04_cost(app):
    try_call(app.ViewApply, "Task Usage", "Использование задач")
    try_call(app.TableApply, "Cost", "Затраты")

# ────────────────── ШАГ 5. Стоимость: парам vs не-парам ───────────────────
@safe_step("Лаба4_ЗадачиРазногоВида.mpp")
def s05_cost_task_types(app):
    # предполагаем, что парам-задачи отмечены Flag1
    if not try_call(app.GroupApply, "Flag1"):   # нет такой → SKIP
        raise AttributeError("group manual")

# ────────────────── ШАГ 6. Стоимость по видам ресурсов ────────────────────
@safe_step("Лаба4_РесурсыРазногоВида.mpp")
def s06_cost_res_type(app):
    try_call(app.ViewApply, "Resource Sheet", "Лист ресурсов")
    try_call(app.GroupApply, "Type", "Тип ресурсов")

# ────────────────── ШАГ 7. Сверхурочные затраты ───────────────────────────
@safe_step("Лаба4_Сверхур.mpp")
def s07_overtime(app):
    try_call(app.ViewApply, "Resource Usage", "Использование ресурсов")
    try_call(app.TableApply, "Usage", "Использование")
    try_call(app.GroupApply, "Overtime Cost")

# ────────────────── ШАГ 8. Риск – слишком короткие задачи ─────────────────
@safe_step("Лаба4_Короткие.mpp")
def s08_short(app):
    # фильтр «Duration < 2d»
    ok = try_call(app.FilterEdit, "ShortDur", True, True,
                  "Duration", "is less than", "2d", None, None, None, False)
    if ok: app.FilterApply("ShortDur")

# ────────────────── ШАГ 9. Риск – длинные с многими ресурсами ─────────────
@safe_step("Лаба4_Длинные.mpp")
def s09_long(app):
    # фильтр по Duration > 20d  AND  Peak > 300 %
    ok = try_call(app.FilterEdit, "LongMany", True, True,
                  "Duration", "is greater than", "20d",
                  "And",
                  "Peak",     "is greater than", "3",  # 300 %
                  False)
    if ok: app.FilterApply("LongMany")

# ────────────────── ШАГ 10. Риск – задачи-ограничения ─────────────────────
@safe_step("Лаба4_Ограничения.mpp")
def s10_constraints(app):
    try_call(app.FilterApply, "Tasks with Fixed Dates", "Задачи с фиксированными датами")

# ────────────────── ШАГ 11. Неопытные сотрудники ──────────────────────────
@safe_step("Лаба4_Неопытные.mpp")
def s11_inexp(app):
    # предполагаем Flag2 = «новичок»
    ok = try_call(app.FilterApply, "Flag2")
    if not ok: raise AttributeError("mark inexperienced manually")

# ────────────────── ШАГ 12. Перегруженные трудовые ресурсы ────────────────
@safe_step("Лаба4_Загруженные.mpp")
def s12_overload(app):
    try_call(app.ViewApply, "Resource Graph", "График ресурсов")

# ────────────────── ШАГ 13. Риск – сверхурочная работа ────────────────────
@safe_step("Лаба4_Сверхурочные.mpp")
def s13_ot_res(app):
    # фильтр по Overtime Work > 0
    ok = try_call(app.FilterEdit, "OTWork", True, True,
                  "Overtime Work", "is greater than", "0h",
                  None, None, None, False)
    if ok: app.FilterApply("OTWork")

# ───────────────────────────── main ───────────────────────────────────────
def run():
    app = Dispatch("MSProject.Application");  app.Visible = True

    if PROJECT.exists():  app.FileOpen(str(PROJECT))
    else:                 app.FileNew(); save_as(app, str(PROJECT))

    steps = [s01_parametric, s02_pert, s03_crit,  s04_cost,
             s05_cost_task_types,  s06_cost_res_type,   s07_overtime,
             s08_short,  s09_long, s10_constraints, s11_inexp,
             s12_overload, s13_ot_res]

    for idx, fn in enumerate(steps, 1):
        fn(app, idx, fn.__name__)
        _save(app, fn.__name__ + ".mpp")   # дублируем итог в /lab4
        time.sleep(PAUSE)

    print(f"✅ ЛР-4 завершена для варианта «{variant}»")

if __name__ == "__main__":
    run()

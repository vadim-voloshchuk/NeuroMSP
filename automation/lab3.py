#!/usr/bin/env python
# coding: utf-8
"""
automation/lab3.py — ЛР-3: создание ресурсов и назначений
"""

from __future__ import annotations
import sys, time, yaml, logging
from pathlib import Path
from win32com.client import Dispatch

from .base  import safe_step
from .utils import ensure_dir, try_call, save_as

ROOT   = Path(__file__).parent.parent.resolve()
CFG_YML = ROOT / "configs" / "lab3.yaml"
LOG_F   = ROOT / "logs" / "lab3.log"
OUTDIR  = ROOT / "outputs" / "lab3"
SHOTDIR = ROOT / "screenshots" / "lab3"
PROJDIR = ROOT / "projects"

# ─── arg variant ──────────────────────────────────────────────────────────
if len(sys.argv) < 2:
    print("Usage: python -m automation.lab3 <variant>"); sys.exit(1)
variant = sys.argv[1]
PROJECT = PROJDIR / f"{variant}.mpp"

# ─── config & log ─────────────────────────────────────────────────────────
cfg = yaml.safe_load(CFG_YML.read_text(encoding="utf-8"))

logging.basicConfig(filename=str(LOG_F), level=logging.INFO,
                    format="%(asctime)s  %(levelname)s  %(message)s",
                    encoding="utf-8")

# ── helpers ────────────────────────────────────────────────────────────────
import re
HOURS_IN_MONTH = 21*8          # 168 ч

def _rate(raw) -> str:
    """
    Приводит «человечью» запись ставки к тому,
    что Project безоговорочно примет.

    • '90000р/мес' → '535.71р/ч'
    • '500р/ч'     → '500р/ч'
    • 15000        → '15000р'
    • None,''      → '0'
    """
    if not raw:
        return "0"

    # число без всяких суффиксов
    if isinstance(raw, (int, float)):
        return f"{float(raw):.2f}р"

    s = (
        str(raw)
        .lower()
        .replace(" ", "")
        .replace(",", ".")          # на всякий
    )
    m = re.match(r"([\d\.]+)(?:р|руб)?(?:/(мес|ч|час))?$", s)
    if not m:
        # не смогли распарсить — отдадим Project «как есть»
        return raw

    val = float(m.group(1))
    unit = m.group(2) or ""         # '' | 'мес' | 'ч'

    if unit.startswith("мес"):
        val = val / HOURS_IN_MONTH   # → «за час»

    # трудовой ресурс → всегда «…р/ч», остальное без «/ч»
    if unit:                        # была единица времени
        return f"{val:.2f}р/ч"
    return f"{val:.2f}р"

def res_by_name(app, name):
    return next((r for r in app.ActiveProject.Resources
                 if r and r.Name == name), None)

def task_by_id(app, tid):
    return next((t for t in app.ActiveProject.Tasks
                 if t and t.ID == tid), None)

def add_asn(task, res):
    """
    Создаёт назначение, если его ещё нет.
    Если дублируется либо Project временно «занят» — записывает в лог
    и возвращает None (скрипт продолжит работу, а не упадёт).
    """
    # уже назначен?
    for a in task.Assignments:
        if a and a.ResourceID == res.ID:
            return a

    try:
        pj = task.Application.ActiveProject
        return pj.Assignments.Add(task.ID, res.ID)
    except Exception as e:
        logging.info("Skip assign %s→%s – %s", task.ID, res.Name, e)
        return None

# ─── шаг-1: ресурсы -------------------------------------------------------
# ─── шаг-3: назначения ----------------------------------------------------
@safe_step("Лаба3_03_Assign.mpp")
def step3(app):
    try_call(app.ViewApply, "Gantt Chart", "Диаграмма Ганта")

    for a in cfg["assignments"]:
        task = task_by_id(app, a["task_id"])
        res  = res_by_name(app, a["resource"])
        if not task or not res:
            logging.info("Skip assign %s→%s (no task/res)",
                         a["task_id"], a["resource"])
            continue

        asn = add_asn(task, res)
        if asn is None:
            continue                       # уже есть либо не удалось

        if res.Type == 0:                  # трудовой
            asn.Units = a.get("units", 100)/100
            if "cost_table" in a:
                asn.CostRateTable = a["cost_table"]

        elif res.Type == 1:                # материал
            asn.Units       = a.get("quantity", 1)
            asn.UnitsFormat = 19           # «шт./д»

        elif res.Type == 2:                # затраты
            asn.Cost = a.get("cost", 0)


# ─── шаг-2: ставки / инициалы --------------------------------------------
@safe_step("Лаба3_02_ResProps.mpp")
def step2(app):
    for rc in cfg["resources"]:
        res = res_by_name(app, rc["name"])
        if not res or res.Type != 0 or "rates" not in rc:
            continue
        for tab, row in rc["rates"].items():
            pr = res.CostRateTables(tab).PayRates(1)
            pr.StandardRate = _rate(row.get("std"))
            pr.OvertimeRate = _rate(row.get("ot"))
            pr.CostPerUse   = _rate(row.get("per_use"))


# ─── шаг-3: назначения ----------------------------------------------------
@safe_step("Лаба3_03_Assign.mpp")
def step3(app):
    try_call(app.ViewApply, "Gantt Chart", "Диаграмма Ганта")
    pj = app.ActiveProject

    def add_asn(task, res):
        for _ in range(3):           # UI может быть занято
            try:
                return pj.Assignments.Add(task.ID, res.ID)
            except Exception:
                time.sleep(0.8)
        raise RuntimeError("cannot add assignment")

    for a in cfg["assignments"]:
        task = task_by_id(app, a["task_id"])
        res  = res_by_name(app, a["resource"])
        if not task or not res:
            logging.info("Skip assign %s→%s", a["task_id"], a["resource"])
            continue

        asn = add_asn(task, res)
        if asn is None:            # дубликат или ошибка – уже залогировали
            continue

        if res.Type == 0:          # трудовой
            asn.Units = a.get("units", 100) / 100
            if "cost_table" in a:
                asn.CostRateTable = a["cost_table"]
        elif res.Type == 1:        # материал
            asn.Units       = a.get("quantity", 1)
            asn.UnitsFormat = 19   # единиц/день
        elif res.Type == 2:        # затраты
            asn.Cost = a.get("cost", 0)


# ─── main -----------------------------------------------------------------
def run():
    for p in (PROJDIR, OUTDIR, SHOTDIR): ensure_dir(p)
    app = Dispatch("MSProject.Application"); app.Visible = True

    if PROJECT.exists(): app.FileOpen(str(PROJECT))
    else:                app.FileNew(); save_as(app, str(PROJECT))

    for i, fn in enumerate((step1, step2, step3), 1):
        fn(app, i, fn.__name__); time.sleep(0.6)

    print(f"✅ ЛР-3 выполнена для проекта «{variant}»")

if __name__ == "__main__":
    run()

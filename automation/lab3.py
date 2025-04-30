#!/usr/bin/env python
# coding: utf-8
"""
automation/lab3.py  –  ЛР-3: создание ресурсов и назначений
Запуск:  python -m automation.lab3 <variant>
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

# ─── аргумент variant ──────────────────────────────────────────────────────
if len(sys.argv) < 2:
    print("Usage: python -m automation.lab3 <variant>")
    sys.exit(1)
variant = sys.argv[1]
PROJECT = PROJDIR / f"{variant}.mpp"

# ─── конфиг YAML ───────────────────────────────────────────────────────────
cfg = yaml.safe_load(CFG_YML.read_text(encoding="utf-8"))

# ─── логирование ───────────────────────────────────────────────────────────
logging.basicConfig(filename=str(LOG_F),
                    level=logging.INFO,
                    format="%(asctime)s  %(levelname)s  %(message)s",
                    encoding="utf-8")

# ─── helpers ───────────────────────────────────────────────────────────────
def res_by_name(app, name):
    return next((r for r in app.ActiveProject.Resources
                 if r and r.Name == name), None)

def task_by_id(app, tid):
    return next((t for t in app.ActiveProject.Tasks
                 if t and t.ID == tid), None)

# ─── шаг-1: добавляем ресурсы ──────────────────────────────────────────────
@safe_step("Лаба3_01_Resources.mpp")
def step1(app):
    try_call(app.ViewApply, "Resource Sheet", "Лист ресурсов")

    for rc in cfg["resources"]:
        if res_by_name(app, rc["name"]):
            continue            # уже есть

        res = app.ActiveProject.Resources.Add(rc["name"])
        t = rc["type"].upper()

        if t == "M":                        # материал
            res.Type          = 1
            res.MaterialLabel = rc.get("units", "")
            res.StandardRate  = rc.get("std_rate", 0)
        elif t == "Z":                      # затратный
            res.Type = 2
        else:                               # трудовой (по умолчанию)
            res.Type = 0

# ─── шаг-2: ставки и прочие свойства трудовых ресурсов ────────────────────
@safe_step("Лаба3_02_ResProps.mpp")
def step2(app):
    for rc in cfg["resources"]:
        res = res_by_name(app, rc["name"])
        if not res or res.Type != 0 or "rates" not in rc:
            continue

        for tab, row in rc["rates"].items():
            pr = res.CostRateTables(tab).PayRates(1)
            pr.StandardRate = row.get("std", 0)
            pr.OvertimeRate = row.get("ot",  0)
            pr.CostPerUse   = row.get("per_use", 0)

        res.Initials = rc.get("short", "")

# ─── шаг-3: назначения ────────────────────────────────────────────────────
@safe_step("Лаба3_03_Assign.mpp")
def step3(app):
    try_call(app.ViewApply, "Gantt Chart", "Диаграмма Ганта")
    pj = app.ActiveProject

    for a in cfg["assignments"]:
        task = task_by_id(app, a["task_id"])
        res  = res_by_name(app, a["resource"])

        if not task or not res:
            logging.info("Skip assign %s→%s", a["task_id"], a["resource"])
            continue

        asn = pj.Assignments.Add(task.ID, res.ID)

        if res.Type == 0:                                    # трудовой
            asn.Units = a.get("units", 100) / 100
            if "cost_table" in a:
                asn.CostRateTable = a["cost_table"]

        elif res.Type == 1:                                  # материал
            asn.Units       = a.get("quantity", 1)
            asn.UnitsFormat = 19            # «шт/д»

        elif res.Type == 2:                                  # затратный
            asn.Cost = a.get("cost", 0)

# ─── main ──────────────────────────────────────────────────────────────────
def run():
    for p in (PROJDIR, OUTDIR, SHOTDIR):
        ensure_dir(p)

    app = Dispatch("MSProject.Application")
    app.Visible = True

    if PROJECT.exists():
        app.FileOpen(str(PROJECT))
    else:
        app.FileNew()
        save_as(app, str(PROJECT))

    for idx, fn in enumerate((step1, step2, step3), start=1):
        fn(app, idx, fn.__name__)
        time.sleep(0.6)

    print(f"✅ ЛР-3 выполнена для проекта «{variant}»")

if __name__ == "__main__":
    run()

#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/lab3.py • ЛР-3: ресурсы и назначения
"""

from __future__ import annotations
import sys, time, yaml, logging
from pathlib import Path
from datetime import datetime
from win32com.client import Dispatch

from .base  import safe_step
from .utils import ensure_dir, try_call, save_as

ROOT   = Path(__file__).parent.parent.resolve()
CFG    = ROOT / "configs" / "lab3.yaml"
LOG    = ROOT / "logs" / "lab3.log"
OUT    = ROOT / "outputs" / "lab3"
SHOT   = ROOT / "screenshots" / "lab3"
PROJ   = ROOT / "projects"

# ─── арг-ты ──────────────────────────────────────────────
if len(sys.argv) < 2:
    print("Usage: python -m automation.lab3 <variant>"); sys.exit(1)
variant = sys.argv[1]
PROJECT = PROJ / f"{variant}.mpp"

# ─── конфиг YAML ────────────────────────────────────────
cfg = yaml.safe_load(Path(CFG).read_text(encoding="utf-8"))

# ─── логирование ───────────────────────────────────────
logging.basicConfig(filename=str(LOG), level=logging.INFO,
                    format="%(asctime)s  %(levelname)s  %(message)s",
                    encoding="utf-8")

# ─── helpers ────────────────────────────────────────────
def get_res(app, name):
    rescoll = app.ActiveProject.Resources
    return next((r for r in rescoll if r and r.Name == name), None)

def get_task(app, task_id):
    tcoll = app.ActiveProject.Tasks
    return next((t for t in tcoll if t and t.ID == task_id), None)

# ─── шаг 1. ресурсы ─────────────────────────────────────
@safe_step("Лаба3_01_Resources.mpp")
def step1(app):
    try_call(app.ViewApply, "Resource Sheet", "Лист ресурсов")
    for rc in cfg["resources"]:
        if get_res(app, rc["name"]):
            continue
        res = app.Resources.Add(rc["name"])
        tp  = rc["type"].upper()
        if tp == "M":                 # материал
            res.Type          = 1
            res.MaterialLabel = rc.get("units", "")
            res.StandardRate  = rc.get("std_rate", 0)
        elif tp == "Z":               # затраты
            res.Type = 2
        else:                         # трудовой
            res.Type = 0

# ─── шаг 2. свойства ресурсов ───────────────────────────
@safe_step("Лаба3_02_ResProps.mpp")
def step2(app):
    for rc in cfg["resources"]:
        res = get_res(app, rc["name"]);  0
        if not res or res.Type != 0 or "rates" not in rc:
            continue
        for tab, row in rc["rates"].items():
            pr = res.CostRateTables(tab).PayRates(1)
            pr.StandardRate  = row.get("std", 0)
            pr.OvertimeRate  = row.get("ot", 0)
            pr.CostPerUse    = row.get("per_use", 0)

# ─── шаг 3. назначения ──────────────────────────────────
@safe_step("Лаба3_03_Assign.mpp")
def step3(app):
    try_call(app.ViewApply, "Gantt Chart", "Диаграмма Ганта")
    pj = app.ActiveProject
    for a in cfg["assignments"]:
        task = get_task(app, a["task_id"])
        res  = get_res(app, a["resource"])
        if not task or not res:
            logging.info("Skip assign %s→%s", a["task_id"], a["resource"]); continue

        assn = pj.Assignments.Add(task.ID, res.ID)

        # трудовые
        if   res.Type == 0:
            assn.Units = a.get("units", 100)/100
            if "cost_table" in a: assn.CostRateTable = a["cost_table"]
        # материалы
        elif res.Type == 1:
            assn.Units       = a.get("quantity", 1)
            assn.UnitsFormat = 19          # XX / d
        # затраты
        elif res.Type == 2:
            assn.Cost = a.get("cost", 0)

# ─── main ───────────────────────────────────────────────
def run():
    for p in (PROJ, OUT, SHOT): ensure_dir(p)
    app = Dispatch("MSProject.Application"); app.Visible = True
    if PROJECT.exists():  app.FileOpen(str(PROJECT))
    else:                 app.FileNew(); save_as(app, str(PROJECT))

    for i, fn in enumerate((step1, step2, step3), 1):
        fn(app, i, fn.__name__);  time.sleep(0.7)

    print(f"✅ ЛР-3 выполнена для {variant}")

if __name__ == "__main__":
    run()

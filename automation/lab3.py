#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/lab3.py • ЛР-3: ресурсы и назначения
Запуск:  python -m automation.lab3 <variant>
(variant определяет, какой projects/<variant>.mpp открыть)
"""

from __future__ import annotations
import sys, time, yaml, logging
from pathlib import Path
from datetime import datetime
from win32com.client import Dispatch

from .base  import safe_step
from .utils import ensure_dir, try_call, save_as, focus

# ─── Пути ──────────────────────────────────────────────
ROOT        = Path(__file__).parent.parent.resolve()
PROJECTS    = ROOT / "projects"
OUTPUTS     = ROOT / "outputs"     / "lab3"
SCREENS     = ROOT / "screenshots" / "lab3"
LOG_FILE    = ROOT / "logs" / "lab3.log"
CFG_FILE    = ROOT / "configs" / "lab3.yaml"
PAUSE       = 0.7

# ─── Аргументы ────────────────────────────────────────
if len(sys.argv) < 2:
    print("Usage: python -m automation.lab3 <variant>")
    sys.exit(1)
variant = sys.argv[1]
PROJECT = PROJECTS / f"{variant}.mpp"

# ─── Читаем YAML конфиг ───────────────────────────────
with open(CFG_FILE, encoding="utf-8") as f:
    cfg = yaml.safe_load(f)

# ─── Логирование ─────────────────────────────────────
logging.basicConfig(filename=str(LOG_FILE), level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s",
                    encoding="utf-8")

# ─── Утилиты для дат ─────────────────────────────────
def date_fmt(dt: str) -> str:
    """''/None/«–»/«НД» → ''  иначе DD.MM.YYYY"""
    if not dt or str(dt).strip(" -—").lower() in ("", "нд"):
        return ""
    return datetime.strptime(dt, "%d.%m.%Y").strftime("%d.%m.%Y")

# ─── 1. Добавляем ресурсы ────────────────────────────
@safe_step("Лаба3_01_Resources.mpp")
def s01_resources(app):
    try_call(app.ViewApply, "Лист ресурсов", "Resource Sheet")

    for res_cfg in cfg["resources"]:
        name = res_cfg["name"]
        if app.Resources(name):
            continue
        res = app.Resources.Add(name)

        t = res_cfg["type"].upper()
        if t == "M":
            res.Type = 1
            res.MaterialLabel = res_cfg.get("units", "")
            res.StandardRate  = res_cfg.get("std_rate", 0)
        elif t == "Z":
            res.Type = 2
        # T трудовой – по умолчанию Type=0

# ─── 2. Свойства ресурсов ────────────────────────────
@safe_step("Лаба3_02_ResProps.mpp")
def s02_res_props(app):
    for rc in cfg["resources"]:
        res = app.Resources(rc["name"])
        if not res: continue

        # краткое имя
        res.Initials = rc.get("short", "")
        # ставки
        if res.Type == 0 and "rates" in rc:
            for tab, row in rc["rates"].items():
                table = res.CostRateTables(tab).PayRates(1)
                table.StandardRate = row.get("std", 0)
                table.OvertimeRate = row.get("ot", 0)
                table.CostPerUse   = row.get("per_use", 0)

# ─── 3. Назначения ───────────────────────────────────
@safe_step("Лаба3_03_Assign.mpp")
def s03_assign(app):
    try_call(app.ViewApply, "Диаграмма Ганта", "Gantt Chart")

    for a in cfg["assignments"]:
        task = app.Tasks(a["task_id"])
        res  = app.Resources(a["resource"])
        if not task or not res:
            logging.info("Skip assign %s→%s", a["task_id"], a["resource"])
            continue

        assn = app.ActiveProject.Assignments.Add(task.ID, res.ID)

        if res.Type == 0:                              # труд.
            assn.Units = a.get("units", 100) / 100
            if "cost_table" in a:
                assn.CostRateTable = a["cost_table"]
        elif res.Type == 1:                            # материал
            assn.Units = a.get("quantity", 1)
            assn.UnitsFormat = 19   # материал/д
        elif res.Type == 2:                            # затраты
            assn.Cost = a.get("cost", 0)

# ─── MAIN ────────────────────────────────────────────
def run():
    for p in (PROJECTS, OUTPUTS, SCREENS): ensure_dir(p)
    app = Dispatch("MSProject.Application"); app.Visible = True

    if PROJECT.exists():
        app.FileOpen(str(PROJECT))
    else:
        app.FileNew(); save_as(app, str(PROJECT))

    # разворачиваем окно Project
    try:
        import pygetwindow as gw; time.sleep(1)
        for w in gw.getWindowsWithTitle(app.Caption): w.maximize(); break
    except Exception:
        pass

    for idx, fn in enumerate((s01_resources, s02_res_props, s03_assign), 1):
        fn(app, idx, fn.__name__)
        time.sleep(PAUSE)

    print(f"✅ ЛР-3 завершена для проекта {variant}")

if __name__ == "__main__":
    run()

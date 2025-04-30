#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/lab5.py • ЛР-5: выравнивание ресур­сов
Запуск:  python -m automation.lab5 <variant>
Работает с projects/<variant>.mpp, созданным в ЛР-3.
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
OUTDIR   = ROOT / "outputs" / "lab5"
SHOTDIR  = ROOT / "screenshots" / "lab5"
LOG_F    = ROOT / "logs" / "lab5.log"
PAUSE    = .7

if len(sys.argv) < 2:
    print("Usage: python -m automation.lab5 <variant>"); sys.exit(1)
variant = sys.argv[1]
PROJECT = PROJDIR / f"{variant}.mpp"

for p in (PROJDIR, OUTDIR, SHOTDIR, LOG_F.parent): ensure_dir(p)

logging.basicConfig(filename=str(LOG_F), level=logging.INFO,
                    format="%(asctime)s  %(levelname)s  %(message)s",
                    encoding="utf-8")

def _save_step(app, fname):            # сохранить в outputs/lab5
    save_as(app, str(OUTDIR / fname))

# ────────────────────── ШАГ 1  авто-левелинг ─────────────────────────────
@safe_step("Лаба5_1.mpp")
def s01_auto_level(app):
    """
    Project → Resource → Level → Level All
    (позиции по умолчанию: только вперед, приоритеты – ID)
    """
    ok = try_call(app.ExecuteMso, "LevelAllResources")
    if not ok:                         # fallback: объектная модель
        try:
            app.LevelResources         = True
            app.LevelNow()
        except Exception as e:
            logging.info("auto-level FAIL: %s", e)
            raise AttributeError("manual auto-level")

# ────────────────────── ШАГ 2  параллельные→последовательные ─────────────
@safe_step("Лаба5_2.mpp")
def s02_seq(app):
    """
    Требуется вручную изменить связи (FS вместо SS / FF и т.д.)
    или поправить календари. Скрипт выводит подсказку и
    помечается SKIP.
    """
    raise AttributeError("convert parallel → sequential manually")

# ────────────────────── ШАГ 3  замена ресурсов ───────────────────────────
@safe_step("Лаба5_3.mpp")
def s03_swap(app):
    """
    Пример автоматической замены: если Task.Peak > 1, меняем
    первого найденного трудового ресурса на ‘Техник’.
    Настройте правила под себя или выполните руками.
    """
    pj = app.ActiveProject
    for t in pj.Tasks:
        if not t or t.PeakUnits <= 1:        # нормальная загрузка
            continue
        for a in t.Assignments:
            if a and a.Resource.Type == 1:   # материал/затраты – пропуск
                continue
            try:
                tech = pj.Resources("Техник")
                if tech and a.ResourceID != tech.ID:
                    a.ResourceID = tech.ID
                    break
            except Exception:
                continue
    # если ничего не сделали – сгенерируем SKIP
    raise AttributeError("check / edit replacements manually")

# ────────────────────── ШАГ 4  ручное редактирование трудозатрат ─────────
@safe_step("Лаба5_4.mpp")
def s04_manual_edit(app):
    """
    Открываем View → Task Usage, дальше пользователь вручную
    переносит часы между днями.    
    """
    try_call(app.ViewApply, "Task Usage", "Использование задач")
    raise AttributeError("manual work contour editing")

# ────────────────────── ШАГ 5  перенос в сверхурочные ────────────────────
@safe_step("Лаба5_5.mpp")
def s05_overtime(app):
    """
    Пример: для перегруженных назначений берём половину Work и
    переносим в OvertimeWork.
    """
    pj = app.ActiveProject
    edited = False
    for a in pj.Assignments:
        if not a:
            continue
        if a.RemainingWork > 0 and a.PeakUnits > 1:
            try:
                ot = a.Work * 0.5
                a.OvertimeWork = ot
                a.Work        -= ot
                edited = True
            except Exception:
                pass
    if not edited:
        raise AttributeError("set overtime manually")

# ───────────────────────────── main ───────────────────────────────────────
def run():
    app = Dispatch("MSProject.Application"); app.Visible = True

    if PROJECT.exists():  app.FileOpen(str(PROJECT))
    else:                 app.FileNew(); save_as(app, str(PROJECT))

    steps = [s01_auto_level, s02_seq, s03_swap,
             s04_manual_edit, s05_overtime]

    for idx, fn in enumerate(steps, 1):
        fn(app, idx, fn.__name__)
        _save_step(app, fn.__name__)          # сохранить mpp в /lab5
        time.sleep(PAUSE)

    print(f"✅ ЛР-5 завершена для варианта «{variant}»")

if __name__ == "__main__":
    run()

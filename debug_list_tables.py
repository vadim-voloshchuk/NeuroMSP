#!/usr/bin/env python
# coding: utf-8
"""
debug_list_tables.py  —  показать списки таблиц MS Project
Usage:  python debug_list_tables.py  "C:\\path\\file.mpp"
"""
import sys, time, itertools
from pathlib import Path
from win32com.client import Dispatch, constants as pj
import pywintypes                         # для обработки com_error

CATS = {0: "(Task)", 1: "(Res)", 2: "(Asn)"}

def enum_tables(app):
    """генератор (name, category) для всех таблиц"""
    # 0 — Task, 1 — Resource, 2 — Assignment
    getters = (
        (0, app.TableList),           # задачи
        (1, app.ResourceTableList),   # ресурсы
        (2, app.AssignmentTableList), # назначения
    )
    for cat, fn in getters:
        for idx in itertools.count(1, 1):
            try:
                name = fn(idx)        # ← string
            except pywintypes.com_error:
                break                 # элементов больше нет
            yield name, cat

def main():
    if len(sys.argv) < 2:
        print("Usage: python debug_list_tables.py <file.mpp>"); return
    mpp = Path(sys.argv[1]).expanduser().resolve()
    print("Opening:", mpp)
    if not mpp.exists(): print("❌ файл не найден"); return

    app = Dispatch("MSProject.Application"); app.Visible = True
    time.sleep(0.4)

    try:
        app.FileOpen(Name=str(mpp))
    except Exception as e:
        print("❌ FileOpen:", e); return

    tables = list(enum_tables(app))
    print(f"=== Всего таблиц: {len(tables)} ===")
    for i, (nm, cat) in enumerate(tables, 1):
        print(f"{i:2d}. {nm:<30s} {CATS[cat]}")

    # закроем файл/Project для чистоты
    time.sleep(1)
    try: app.FileClose(pj.pjDoNotSave)
    except: pass
    try: app.Quit()
    except: pass

if __name__ == "__main__":
    main()

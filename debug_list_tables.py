#!/usr/bin/env python
# coding: utf-8
"""
debug_list_tables.py  —  показать списки таблиц MS Project
Usage:  python debug_list_tables.py  "<путь>\variant1.mpp"
"""
import sys, time, itertools
from pathlib import Path
from win32com.client import Dispatch
import pywintypes   # для обработки com_error

CATS = {0: "(Task)", 1: "(Res)", 2: "(Asn)"}

def enum_tables(pj):
    """Генератор (name, category) для всех таблиц проекта"""
    getters = (
        (0, pj.TableList),
        (1, pj.ResourceTableList),
        (2, pj.AssignmentTableList),
    )
    for cat, fn in getters:
        for idx in itertools.count(1):
            try:
                name = fn(idx)
            except pywintypes.com_error:
                break
            yield name, cat

def main():
    if len(sys.argv) < 2:
        print("Usage: python debug_list_tables.py <file.mpp>"); return

    mpp = Path(sys.argv[1]).expanduser().resolve()
    print("Opening:", mpp)
    if not mpp.exists():
        print("❌ файл не найден"); return

    app = Dispatch("MSProject.Application")
    app.Visible = True
    time.sleep(0.5)

    try:
        app.FileOpen(Name=str(mpp))
    except Exception as e:
        print("❌ FileOpen error:", e); return

    pj = app.ActiveProject
    tables = list(enum_tables(pj))
    print(f"=== Всего таблиц: {len(tables)} ===")
    for i, (nm, cat) in enumerate(tables, 1):
        print(f"{i:2d}. {nm:<30s} {CATS[cat]}")

    # аккуратно закрываем
    time.sleep(0.5)
    try: app.FileClose(0)  # pjDoNotSave
    except: pass
    try: app.Quit()
    except: pass

if __name__ == "__main__":
    main()

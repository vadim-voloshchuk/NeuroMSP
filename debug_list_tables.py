#!/usr/bin/env python
# coding: utf-8
"""
debug_list_tables.py — вывести все таблицы (Application.Tables()) в текущем MPP
Usage:  python debug_list_tables.py "C:\\path\\to\\file.mpp"
"""
import sys, time
from pathlib import Path
from win32com.client import Dispatch

def main():
    if len(sys.argv) < 2:
        print("Usage: python debug_list_tables.py <file.mpp>"); return

    mpp = Path(sys.argv[1]).absolute()
    print("Opening:", mpp)
    if not mpp.exists():
        print("❌ файл не найден"); return

    app = Dispatch("MSProject.Application")
    app.Visible = True
    time.sleep(0.5)

    try:
        app.FileOpen(Name=str(mpp))
    except Exception as e:
        print("❌ FileOpen:", e); return

    # ► ключевая разница ↓↓↓
    try:
        tables = app.Tables()          # вызываем метод → получаем коллекцию
    except Exception as e:
        print("❌ Application.Tables():", e); return

    print(f"=== Всего таблиц: {tables.Count} ===")
    for i in range(1, tables.Count + 1):
        tbl = tables.Item(i)
        # tbl.Category: 0-Task, 1-Resource, 2-Assignment
        kind = {0:"(Task)",1:"(Res)",2:"(Asn)"}[tbl.Category]
        print(f"{i:2d}. {tbl.Name:<30s} {kind}")

    # очистка
    time.sleep(1)
    try: app.FileClose()
    except: pass
    try: app.Quit()
    except: pass

if __name__ == "__main__":
    main()

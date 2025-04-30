#!/usr/bin/env python
# coding: utf-8
"""
debug_list_views.py — вывести все представления ActiveProject.Views
Usage: python debug_list_views.py "C:\полный\путь\до\файла.mpp"
"""
import sys, time
from pathlib import Path
from win32com.client import Dispatch

def main():
    if len(sys.argv) < 2:
        print("Usage: python debug_list_views.py \"C:\\path\\to\\file.mpp\"")
        return

    proj_path = Path(sys.argv[1]).absolute()
    print("Opening:", proj_path)
    if not proj_path.exists():
        print("❌ Файл не найден")
        return

    app = Dispatch("MSProject.Application")
    app.Visible = True
    time.sleep(0.5)
    try:
        app.FileOpen(Name=str(proj_path))
    except Exception as e:
        print("❌ Ошибка FileOpen:", e)
        return

    pj    = app.ActiveProject
    views = pj.Views
    print("=== Всего представлений:", views.Count, "===")
    for i in range(1, views.Count + 1):
        v = views.Item(i)
        print(f"{i:2d}. {v.Name}")

    # cleanup
    time.sleep(1)
    try: app.FileClose()
    except: pass
    try: app.Quit()
    except: pass

if __name__ == "__main__":
    main()

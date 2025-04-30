#!/usr/bin/env python
# coding: utf-8
"""
debug_list_windows.py — вывести все окна MS Project, доступные через Application.Windows
Usage: python debug_list_windows.py "C:\path\to\file.mpp"
"""
import sys, time
from pathlib import Path
from win32com.client import Dispatch

def main():
    if len(sys.argv) < 2:
        print("Usage: python debug_list_windows.py \"C:\\path\\to\\file.mpp\"")
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

    try:
        wins = app.Windows
    except Exception as e:
        print("❌ Не удалось получить Application.Windows:", e)
        return

    print(f"=== Всего окон: {wins.Count} ===")
    for i in range(1, wins.Count + 1):
        w = wins.Item(i)
        print(f"{i:2d}. {w.Caption}")

    # cleanup
    time.sleep(1)
    try: app.FileClose()
    except: pass
    try: app.Quit()
    except: pass

if __name__ == "__main__":
    main()

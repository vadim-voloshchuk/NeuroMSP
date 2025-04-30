#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Скрипт для теста открытия указанного представления
Usage: python debug_open_view.py project_file.mpp "Gantt Chart"
"""
import sys
from win32com.client import Dispatch

def main():
    if len(sys.argv) < 3:
        print("Usage: python debug_open_view.py project_file.mpp \"View Name\"")
        sys.exit(1)
    proj_path, view_name = sys.argv[1], sys.argv[2]

    app = Dispatch("MSProject.Application")
    app.Visible = True
    app.FileOpen(proj_path)

    try:
        print(f"Switching to view '{view_name}'...")
        app.ViewApply(view_name)
        print("OK")
    except Exception as e:
        print("FAIL:", e)

    app.FileClose()
    app.Quit()

if __name__ == "__main__":
    main()

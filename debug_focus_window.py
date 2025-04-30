#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Скрипт для вывода всех доступных представлений в файле .mpp
Usage: python debug_list_views.py project_file.mpp
"""
import sys
from win32com.client import Dispatch

def main():
    if len(sys.argv) < 2:
        print("Usage: python debug_list_views.py project_file.mpp")
        sys.exit(1)
    proj_path = sys.argv[1]

    app = Dispatch("MSProject.Application")
    app.Visible = True
    app.FileOpen(proj_path)

    print("=== All Views ===")
    for v in app.Views:
        print(f"- {v.Name}")

    app.FileClose()
    app.Quit()

if __name__ == "__main__":
    main()

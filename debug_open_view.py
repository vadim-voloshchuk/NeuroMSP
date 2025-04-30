#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Проверка фокуса и максимизации окна MS Project
Usage: python debug_focus_window.py project_file.mpp
"""
import sys
import time
from win32com.client import Dispatch
import pygetwindow as gw

def main():
    if len(sys.argv) < 2:
        print("Usage: python debug_focus_window.py project_file.mpp")
        return
    proj_path = sys.argv[1]

    app = Dispatch("MSProject.Application")
    app.Visible = True
    app.FileOpen(proj_path)
    time.sleep(1)  # даём время окну появиться

    caption = app.Caption
    print("Application.Caption =", caption)

    wins = gw.getWindowsWithTitle(caption)
    if not wins:
        print("Не удалось найти окно по заголовку.")
    else:
        w = wins[0]
        print("Found window:", w)
        print(" - IsActive before:", w.isActive)
        w.maximize()
        time.sleep(0.5)
        w.activate()
        time.sleep(0.5)
        print(" - IsActive after:", w.isActive)

    app.FileClose()
    app.Quit()

if __name__ == "__main__":
    main()

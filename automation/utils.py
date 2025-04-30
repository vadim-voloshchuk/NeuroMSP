# automation/utils.py
#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/utils.py • Утилиты: директории, фокус окон, скриншоты и безопасные вызовы
"""
import time
from pathlib import Path
import pyautogui
import pygetwindow as gw


def ensure_dir(path):
    """
    Создаёт директорию и все промежуточные, если их нет.
    """
    Path(path).mkdir(parents=True, exist_ok=True)


def try_call(obj, *names, **kwargs):
    """
    Пытается вызвать метод объекта по списку возможных имён.
    Возвращает True при первом успехе, иначе False.
    """
    for name in names:
        try:
            getattr(obj, name)(**kwargs)
            return True
        except Exception:
            continue
    return False


def focus(title):
    """
    Активирует окно с данным заголовком. Пауза 0.3 с после активации.
    """
    for w in gw.getWindowsWithTitle(title):
        if not w.isActive:
            w.activate()
            time.sleep(0.3)
            break


def shot(idx, code, out_dir="screenshots"):
    """
    Делаем скриншот экрана и сохраняем в директорию out_dir
    с именем {idx:02d}_{code}.png.
    """
    ensure_dir(out_dir)
    filename = Path(out_dir) / f"{idx:02d}_{code}.png"
    pyautogui.screenshot(str(filename))


def save_as(app, fname, out_dir="outputs"):
    """
    Сохраняет текущий проект MS Project как файл fname в out_dir.
    """
    ensure_dir(out_dir)
    fullpath = Path(out_dir) / fname
    app.FileSaveAs(str(fullpath))
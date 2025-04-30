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
    Path(path).mkdir(parents=True, exist_ok=True)


def try_call(obj, *names, **kwargs):
    for name in names:
        try:
            getattr(obj, name)(**kwargs)
            return True
        except Exception:
            continue
    return False


def focus(title):
    for w in gw.getWindowsWithTitle(title):
        if not w.isActive:
            w.activate()
            time.sleep(0.3)
            break


def shot(idx, code, out_dir="screenshots"):
    ensure_dir(out_dir)
    fn = Path(out_dir) / f"{idx:02d}_{code}.png"
    pyautogui.screenshot(str(fn))


def save_as(app, fname, out_dir="outputs"):
    """
    Если fname — абсолютный путь, сохраняем туда напрямую.
    Иначе — в папку out_dir/fname.
    """
    p = Path(fname)
    if p.is_absolute():
        # создаём директорию, если нужно
        p.parent.mkdir(parents=True, exist_ok=True)
        app.FileSaveAs(str(p))
    else:
        ensure_dir(out_dir)
        full = Path(out_dir) / fname
        full.parent.mkdir(parents=True, exist_ok=True)
        app.FileSaveAs(str(full))

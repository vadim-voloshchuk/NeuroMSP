#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/utils.py • Утилиты: директории, фокус окон, скриншоты и безопасные вызовы
"""
import time
from pathlib import Path
import pyautogui
import pygetwindow as gw

# win32gui / win32con будут, если установлена pywin32
try:
    import win32gui, win32con
except ImportError:
    win32gui = None
    win32con = None

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

def focus(title: str):
    """
    Мягко выводим окно MS Project на передний план.
    Без Exception — чтобы safe_step никогда не падал.
    """
    try:
        for w in gw.getWindowsWithTitle(title):
            # попытка Win32 API
            if win32gui:
                try:
                    win32gui.ShowWindow(w._hWnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(w._hWnd)
                except Exception:
                    pass
            # попытка метода activate()
            try:
                w.activate()
            except Exception:
                pass
            time.sleep(0.3)
            break
    except Exception:
        pass  # полностью подавляем все ошибки

    
def shot(idx, code, out_dir="screenshots"):
    """
    Скриншот и сохранение в абсолютный путь.
    """
    ensure_dir(out_dir)
    fn = Path(out_dir) / f"{idx:02d}_{code}.png"
    pyautogui.screenshot(str(fn.resolve()))

def save_as(app, fname, out_dir="outputs"):
    """
    Сохраняет файл:
      - если fname — абсолютный путь, прямо в него;
      - иначе — в out_dir/fname (абсолютно).
    Всегда приводит к абсолютному пути до вызова COM.
    """
    p = Path(fname)
    if p.is_absolute():
        target = p
    else:
        ensure_dir(out_dir)
        target = Path(out_dir) / fname
    # приводим к абсолютному
    target = target.resolve()
    ensure_dir(target.parent)
    try:
        app.FileSaveAs(str(target))
    except Exception as e:
        # предупреждаем, но не ломаем flow
        print(f"⚠️ Warning: не удалось сохранить {target}: {e}")

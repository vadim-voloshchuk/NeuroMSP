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

# automation/base.py
#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/base.py • Декоратор safe_step для безопасного выполнения шагов
"""
import logging
import time
from .utils import focus, shot, save_as


def safe_step(save_fname=None):
    """
    Декоратор для функции-шага:
      - ловит AttributeError → пометка SKIP
      - ловит другие Exception → пометка FAIL
      - логирует результат
      - сохраняет проект, если указан save_fname
      - делает скриншот и выводит [idx] code status
    """
    def decorator(fn):
        def wrapper(app, idx, code, *args, **kwargs):
            status = "OK"
            try:
                fn(app, *args, **kwargs)
                if save_fname:
                    save_as(app, save_fname)
            except AttributeError as ae:
                status = "SKIP"
                logging.info("SKIP %s: %s", code, ae)
            except Exception as e:
                status = "FAIL"
                logging.exception("FAIL %s", code)
            # фокус на окне проекта и скриншот
            focus(app.Caption)
            shot(idx, code)
            print(f"[{idx:02d}] {code:22} {status}")
            time.sleep(0.5)
        return wrapper
    return decorator

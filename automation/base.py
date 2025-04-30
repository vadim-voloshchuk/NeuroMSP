#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/base.py • Декоратор safe_step для безопасного выполнения шагов
"""
import logging, time
from .utils import focus, shot, save_as

def safe_step(save_fname=None):
    """
    Декоратор:
      • выполняет шаг fn
      • пытается сохранить (если save_fname)
      • в любом случае делает скрин и выводит статус
    """
    def decorator(fn):
        def wrapper(app, idx, code, *args, **kwargs):
            status = "OK"
            try:
                fn(app, *args, **kwargs)
                if save_fname:
                    try:
                        save_as(app, save_fname)
                    except Exception:
                        logging.exception("FAIL save_as %s", code)
                        status = "FAIL"
            except AttributeError as ae:
                status = "SKIP"
                logging.info("SKIP %s: %s", code, ae)
            except Exception:
                status = "FAIL"
                logging.exception("FAIL %s", code)
            # даже при ошибках — делаем скрин
            focus(app.Caption)
            shot(idx, code)
            print(f"[{idx:02d}] {code:22} {status}")
            time.sleep(0.5)
        return wrapper
    return decorator

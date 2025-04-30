#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
automation/base.py • Декоратор safe_step для безопасного выполнения шагов
"""
import logging, time
from .utils import focus, shot, save_as

def safe_step(save_name: str | None = None):
    """
    • если save_name передан – всегда пытаемся сохранить
      (OK / SKIP / FAIL), чтобы файл-пустышка появился;
    • если исключение внутри шага – логируем, но не валимся.
    """
    def deco(fn):
        def wrapper(app, idx, code, *a, **kw):
            status = "OK"
            try:
                fn(app, *a, **kw)
            except AttributeError as e:  # помечаем шаг «SKIP»
                status = "SKIP"
                logging.info("SKIP %s: %s", code, e)
            except Exception:
                status = "FAIL"
                logging.exception("FAIL %s", code)

            # сохранить MPP-файл, даже если SKIP/FAIL
            if save_name:
                try:
                    save_as(app, f"{save_name}")
                except Exception:
                    logging.exception("save_as %s", save_name)

            focus(app.Caption);  shot(idx, code)
            print(f"[{idx:02d}] {code:24} {status}")
        return wrapper
    return deco

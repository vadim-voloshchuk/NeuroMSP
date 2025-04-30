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

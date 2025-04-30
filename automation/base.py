# ── файлик: automation/base.py ───────────────────────
import logging, time
from .utils import focus, shot, save_as

def safe_step(save_fname=None):
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
            except Exception as e:
                status = "FAIL"
                logging.exception("FAIL %s", code)
            # даже если были ошибки, делаем скрин
            focus(app.Caption)
            shot(idx, code)          # снимок всего экрана
            print(f"[{idx:02d}] {code:22} {status}")
            time.sleep(0.5)
        return wrapper
    return decorator

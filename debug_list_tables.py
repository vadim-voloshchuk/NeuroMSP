#!/usr/bin/env python
# coding: utf-8
"""
test_log_tables.py — тестирует методы работы с таблицами MS Project
Usage: python test_log_tables.py "<путь>\variant1.mpp"
"""

import sys
import logging
from pathlib import Path
from win32com.client import Dispatch
import pywintypes
import itertools

logging.basicConfig(
    filename='project_tables.log',
    filemode='w',
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

CATS = {0: "(Task)", 1: "(Res)", 2: "(Asn)"}

def enum_tables(pj):
    collections = (
        (0, pj.TaskTables),
        (1, pj.ResourceTables),
        (2, pj.AssignmentTables),
    )
    for cat, collection in collections:
        for table in collection:
            yield table.Name, cat


def test_methods(app, pj):
    try:
        logging.info("Проверка метода TableEdit")
        app.TableEdit(Name="Entry", TaskTable=True, Create=True, NewName="Test Table", ShowInMenu=True)
        logging.info("✅ TableEdit успешно выполнен")
    except Exception as e:
        logging.error(f"❌ TableEdit ошибка: {e}")

    try:
        logging.info("Проверка метода Tables.Add")
        pj.TaskTables.Add(Name="MyTaskTable", Field=188743731, Task=True)  # Field=188743731 — Name
        logging.info("✅ Tables.Add успешно выполнен")
    except Exception as e:
        logging.error(f"❌ Tables.Add ошибка: {e}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python test_log_tables.py <file.mpp>")
        return

    mpp = Path(sys.argv[1]).expanduser().resolve()
    logging.info(f"Открываю файл: {mpp}")

    if not mpp.exists():
        logging.error("❌ файл не найден")
        return

    app = Dispatch("MSProject.Application")
    app.Visible = True

    try:
        app.FileOpen(Name=str(mpp))
        pj = app.ActiveProject
    except Exception as e:
        logging.error(f"❌ FileOpen ошибка: {e}")
        return

    tables = list(enum_tables(pj))
    logging.info(f"=== Всего таблиц: {len(tables)} ===")
    for i, (nm, cat) in enumerate(tables, 1):
        logging.info(f"{i:2d}. {nm:<30s} {CATS[cat]}")

    test_methods(app, pj)

    try:
        app.FileClose(0)  # pjDoNotSave
        app.Quit()
    except Exception as e:
        logging.warning(f"Предупреждение при закрытии: {e}")

    print("📄 Лог сохранён в project_tables.log")


if __name__ == "__main__":
    main()

# Automation of Microsoft Project Labs

Этот репозиторий содержит скрипты на Python для автоматизации выполнения лабораторных работ по Microsoft Project.

## Структура проекта

```
project-automation/
├── automation/
│   ├── base.py               # Декораторы и общие helper’ы
│   ├── utils.py              # Функции focus, shot, try_call и т.п.
│   ├── lab1.py               # Скрипт для Лабораторной работы 1
│   ├── lab2.py               # Скрипт для ЛР-2
│   ├── lab3.py               # Скрипт для ЛР-3
│   ├── lab4.py               # Скрипт для ЛР-4
│   ├── lab5.py               # Скрипт для ЛР-5
│   ├── lab6.py               # Скрипт для ЛР-6
│   ├── lab7.py               # Скрипт для ЛР-7
│   └── __init__.py
├── configs/
│   └── variants.yaml         # Параметры для вариантов (1–5)
├── outputs/                  # Сохранённые .mpp/.xls файлы
├── screenshots/              # Скриншоты после каждого шага
├── logs/                     # Логи прогонов
├── main.py                   # CLI для запуска нужной ЛР
├── requirements.txt          # Зависимости Python
└── README.md
```

## Установка

1. Клонируйте репозиторий:  
   ```bash
   git clone <url>
   cd project-automation
   ```

2. Создайте виртуальное окружение и установите зависимости:  
   ```bash
   python -m venv venv
   source venv/bin/activate  # или venv\Scripts\activate
   pip install -r requirements.txt
   ```

3. Настройте файл `configs/variants.yaml`, указав абсолютные пути к вашим `.mpp` файлам и параметры вариантов.

## Использование

Для запуска ЛР-1:  
```bash
python -m automation.lab1 variant1
```
Для ЛР-2…ЛР-7 аналогично:
```bash
python main.py --lab 2 --variant variant3
```
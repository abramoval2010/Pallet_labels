# app/utils/helpers.py
import os
import re
import math
import sys
from datetime import datetime


def is_desktop_mode():
    """Определяет, запущено ли приложение в десктопном режиме"""
    # Проверяем переменную окружения (приоритет)
    if os.environ.get('DESKTOP_MODE') == 'false':
        return False
    if os.environ.get('DESKTOP_MODE') == 'true':
        return True

    # Проверяем, запущен ли скрипт как замороженное приложение
    if getattr(sys, 'frozen', False):
        return True

    # Проверяем имя файла
    script_name = os.path.basename(sys.argv[0]).lower()
    if script_name in ['desktop.py', 'rmo_osl.exe']:
        return True
    if '.exe' in script_name and 'rmo_osl' in script_name:
        return True

    # Проверяем аргументы
    if '--desktop' in sys.argv:
        return True

    # Проверяем Render
    if os.environ.get('RENDER'):
        return False

    # По умолчанию - web-режим
    return False


def get_app_data_dir(app_name="RMO_OSL"):
    """Возвращает путь к папке данных приложения"""
    desktop_mode = is_desktop_mode()

    if desktop_mode:
        # Десктопный режим - используем стандартные папки ОС
        if sys.platform == 'win32':
            base_dir = os.environ.get('APPDATA', os.path.expanduser('~\\AppData\\Roaming'))
            app_dir = os.path.join(base_dir, app_name)
        elif sys.platform == 'darwin':
            app_dir = os.path.expanduser(f'~/Library/Application Support/{app_name}')
        else:
            app_dir = os.path.expanduser(f'~/.local/share/{app_name}')
    else:
        # Web-режим - используем папку проекта
        # Определяем корневую папку проекта
        current_file = os.path.abspath(__file__)  # /path/to/app/utils/helpers.py
        app_dir = os.path.dirname(os.path.dirname(os.path.dirname(current_file)))  # /path/to/project/

    os.makedirs(app_dir, exist_ok=True)
    return app_dir

def calculate_pallets(quantity, palletization):
    """Рассчитывает количество паллет"""
    if not quantity:
        return 1
    try:
        qty = int(quantity)
    except:
        return 1
    if palletization > 0:
        try:
            return math.ceil(qty / palletization)
        except:
            return 1
    return 1


def clean_filename(text):
    """Очищает текст от недопустимых символов для имени файла"""
    if not text:
        return 'Неизвестный_поставщик'
    invalid_chars = r'[<>:"/\\|?*]'
    cleaned = re.sub(invalid_chars, '', text)
    cleaned = ' '.join(cleaned.split())
    if not cleaned.strip():
        return 'Неизвестный_поставщик'
    return cleaned


def get_unique_filename(base_path, base_name, extension='.docx'):
    """Формирует уникальное имя файла с индексацией при повторении"""
    counter = 1
    while True:
        if counter == 1:
            filename = f"{base_name}{extension}"
        else:
            filename = f"{base_name}_{counter}{extension}"

        full_path = os.path.join(base_path, filename)
        if not os.path.exists(full_path):
            return full_path
        counter += 1


def open_in_word(filepath):
    """Открывает файл в Microsoft Word"""
    try:
        if sys.platform == 'win32':
            os.startfile(filepath)
            return True
        elif sys.platform == 'darwin':
            import subprocess
            subprocess.run(['open', filepath])
            return True
        else:
            import subprocess
            subprocess.run(['xdg-open', filepath])
            return True
    except Exception as e:
        print(f"Ошибка открытия в Word: {e}")
        return False
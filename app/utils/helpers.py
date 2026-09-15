# app/utils/helpers.py
import os
import re
import math
import sys
import unicodedata
from datetime import datetime


def is_desktop_mode():
    """Определяет, запущено ли приложение в десктопном режиме"""
    if os.environ.get('DESKTOP_MODE') == 'false':
        return False
    if os.environ.get('DESKTOP_MODE') == 'true':
        return True

    if getattr(sys, 'frozen', False):
        return True

    script_name = os.path.basename(sys.argv[0]).lower()
    if script_name in ['desktop.py', 'rmo_osl.exe']:
        return True
    if '.exe' in script_name and 'rmo_osl' in script_name:
        return True

    if '--desktop' in sys.argv:
        return True

    if os.environ.get('RENDER'):
        return False

    return False


def get_app_data_dir(app_name="RMO_OSL"):
    """Возвращает путь к папке данных приложения"""
    desktop_mode = is_desktop_mode()

    if desktop_mode:
        if sys.platform == 'win32':
            base_dir = os.environ.get('APPDATA', os.path.expanduser('~\\AppData\\Roaming'))
            app_dir = os.path.join(base_dir, app_name)
        elif sys.platform == 'darwin':
            app_dir = os.path.expanduser(f'~/Library/Application Support/{app_name}')
        else:
            app_dir = os.path.expanduser(f'~/.local/share/{app_name}')
    else:
        current_file = os.path.abspath(__file__)
        app_dir = os.path.dirname(os.path.dirname(os.path.dirname(current_file)))

    os.makedirs(app_dir, exist_ok=True)
    return app_dir


def normalize_user_path(raw_path):
    """
    Приводит путь, введённый пользователем, к «чистому» виду:
      - убирает обрамляющие пробелы, табы, переводы строк;
      - снимает обрамляющие кавычки (одинарные или двойные),
        которые часто попадают при копировании через «Копировать как путь» в Windows;
      - нормализует Unicode к форме NFC,
        чтобы кириллические буквы совпадали с тем, как их хранит файловая система;
      - убирает завершающий слэш/бэкслэш (кроме корня диска вида "C:\\").
    Возвращает очищенную строку.
    """
    if raw_path is None:
        return ''

    p = str(raw_path).strip()

    # Снимаем обрамляющие кавычки
    if len(p) >= 2 and p[0] == p[-1] and p[0] in ('"', "'"):
        p = p[1:-1].strip()

    # Нормализация Unicode — критично для кириллицы
    p = unicodedata.normalize('NFC', p)

    # Убираем завершающий слэш/бэкслэш, но не трогаем "C:\" и "/"
    if len(p) > 3:
        if p.endswith('\\') or p.endswith('/'):
            # не убираем, если это корень диска "C:\"
            if not (len(p) == 3 and p[1] == ':'):
                p = p.rstrip('\\/')

    return p


def calculate_pallets(quantity, palletization):
    """Рассчитывает количество паллет"""
    if not quantity:
        return 1
    try:
        qty = int(quantity)
    except Exception:
        return 1
    if palletization > 0:
        try:
            return math.ceil(qty / palletization)
        except Exception:
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
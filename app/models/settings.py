# app/models/settings.py
import os
import json
import sys
from app.utils.helpers import get_app_data_dir

DEFAULT_SETTINGS = {
    'save_path': None,  # Будет установлен при инициализации
    'sop_code': 'RUPD.VLG.05.04.C01',
    'font_settings': {
        'header_size': 12,
        'product_size_short': 80,
        'product_size_medium': 72,
        'product_size_long': 64,
        'supplier_size_short': 42,
        'supplier_size_medium': 42,
        'supplier_size_long': 36,
        'analytic_size': 117,
        'code_size': 14,
        'line_spacing_short': 1.5,
        'line_spacing_medium': 1.0,
        'line_spacing_long': 0.8,
        'short_name_threshold': 25,
        'medium_name_threshold': 40
    }
}


class SettingsManager:
    def __init__(self, settings_file, generated_labels_dir):
        self.settings_file = settings_file
        self.generated_labels_dir = generated_labels_dir
        self._settings = None
        self._load()

    def _load(self):
        """Загружает настройки из файла"""
        default = DEFAULT_SETTINGS.copy()
        default['save_path'] = self.generated_labels_dir

        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)

                # Восстанавливаем недостающие ключи
                if 'save_path' not in settings:
                    settings['save_path'] = default['save_path']
                if 'sop_code' not in settings:
                    settings['sop_code'] = default['sop_code']
                if 'font_settings' not in settings:
                    settings['font_settings'] = default['font_settings']
                else:
                    for key in default['font_settings']:
                        if key not in settings['font_settings']:
                            settings['font_settings'][key] = default['font_settings'][key]

                self._settings = settings
            else:
                os.makedirs(os.path.dirname(self.settings_file), exist_ok=True)
                self._settings = default
                self._save()
        except Exception as e:
            print(f"Ошибка загрузки настроек: {e}")
            self._settings = default

    def _save(self):
        """Сохраняет настройки в файл"""
        try:
            os.makedirs(os.path.dirname(self.settings_file), exist_ok=True)
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self._settings, f, ensure_ascii=False, indent=4)
            return True
        except Exception as e:
            print(f"Ошибка сохранения настроек: {e}")
            return False

    def get(self, key=None):
        """Получить настройку или все настройки"""
        if key is None:
            return self._settings.copy()
        return self._settings.get(key)

    def set(self, key, value):
        """Установить настройку и сохранить"""
        if key == 'font_settings':
            if 'font_settings' not in self._settings:
                self._settings['font_settings'] = {}
            self._settings['font_settings'].update(value)
        else:
            self._settings[key] = value
        self._save()

    def get_font_settings(self):
        """Получить настройки шрифтов"""
        return self._settings.get('font_settings', DEFAULT_SETTINGS['font_settings'])
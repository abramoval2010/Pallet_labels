# app/__init__.py
import os
import sys
from pathlib import Path
from datetime import timedelta
from flask import Flask
from flask_session import Session

from app.config import Config
from app.utils.helpers import is_desktop_mode, get_app_data_dir


def create_app():
    """Фабрика создания Flask-приложения"""

    # Определяем режим работы
    DESKTOP_MODE = is_desktop_mode()

    # Определяем пути - ВСЕГДА используем папку проекта для web-режима
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Для web-режима данные хранятся в папке проекта
    if DESKTOP_MODE:
        USER_DATA_DIR = get_app_data_dir("RMO_OSL")
    else:
        USER_DATA_DIR = BASE_DIR

    # Копируем пути в Config
    Config.BASE_DIR = BASE_DIR
    Config.USER_DATA_DIR = USER_DATA_DIR
    Config.DESKTOP_MODE = DESKTOP_MODE

    # Пути к файлам
    Config.SETTINGS_FILE = os.path.join(USER_DATA_DIR, 'settings.json')
    Config.DATABASE_FILE = os.path.join(USER_DATA_DIR, 'users.db')
    Config.MATERIALS_DATABASE_FILE = os.path.join(USER_DATA_DIR, 'materials.db')
    Config.MASTER_KEY_FILE = os.path.join(USER_DATA_DIR, 'master.key')
    Config.GENERATED_LABELS_DIR = os.path.join(USER_DATA_DIR, 'generated_labels')
    Config.UPLOADS_DIR = os.path.join(USER_DATA_DIR, 'uploads')
    Config.SESSION_DIR = os.path.join(USER_DATA_DIR, 'flask_session')

    # Создаем необходимые папки
    os.makedirs(Config.GENERATED_LABELS_DIR, exist_ok=True)
    os.makedirs(Config.UPLOADS_DIR, exist_ok=True)
    os.makedirs(Config.SESSION_DIR, exist_ok=True)

    print("=" * 60)
    print("RMO OSL Application Initialization")
    print("=" * 60)
    print(f"Mode: {'DESKTOP' if DESKTOP_MODE else 'WEB'}")
    print(f"Base folder: {BASE_DIR}")
    print(f"Data folder: {USER_DATA_DIR}")
    print(f"Database: {Config.DATABASE_FILE}")
    print(f"Master key: {Config.MASTER_KEY_FILE}")
    print("=" * 60)

    # Создаем приложение
    app = Flask(__name__,
                template_folder=os.path.join(BASE_DIR, 'templates'),
                static_folder=os.path.join(BASE_DIR, 'static'))

    app.config['SECRET_KEY'] = Config.SECRET_KEY
    app.config['MAX_CONTENT_LENGTH'] = Config.MAX_CONTENT_LENGTH
    app.config['UPLOAD_FOLDER'] = Config.UPLOADS_DIR
    app.config['SESSION_PERMANENT'] = Config.SESSION_PERMANENT
    app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(days=Config.PERMANENT_SESSION_LIFETIME)

    # Настройка сессий для web-режима
    if DESKTOP_MODE:
        app.config['SESSION_TYPE'] = 'filesystem'
        app.config['SESSION_FILE_DIR'] = Config.SESSION_DIR
        app.config['SESSION_FILE_THRESHOLD'] = 500
        app.config['SESSION_FILE_MODE'] = 0o600
        app.debug = False

        try:
            Session(app)
            print("Flask-Session initialized for desktop mode")
        except ImportError:
            print("Flask-Session not installed, using fallback")
    else:
        # В web-режиме используем простые сессии
        app.config['SESSION_COOKIE_DOMAIN'] = None
        app.config['SESSION_COOKIE_PATH'] = '/'
        app.config['SESSION_COOKIE_HTTPONLY'] = True
        app.config['SESSION_COOKIE_SECURE'] = False
        app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
        app.config['SESSION_PERMANENT'] = True
        app.debug = True
        print("Web mode: using simple session handling")

    # Инициализация БД и настроек
    from app.models.user_db import UserDatabase
    from app.models.materials_db import MaterialsDatabase
    from app.models.settings import SettingsManager

    db = UserDatabase(Config.DATABASE_FILE, Config.MASTER_KEY_FILE)
    materials_db = MaterialsDatabase(Config.MATERIALS_DATABASE_FILE)
    settings_manager = SettingsManager(Config.SETTINGS_FILE, Config.GENERATED_LABELS_DIR)

    # Сохраняем в конфиг приложения
    app.config['db'] = db
    app.config['materials_db'] = materials_db
    app.config['settings_manager'] = settings_manager
    app.config['DESKTOP_MODE'] = DESKTOP_MODE
    app.config['GENERATED_LABELS_DIR'] = Config.GENERATED_LABELS_DIR
    app.config['UPLOADS_DIR'] = Config.UPLOADS_DIR

    # Регистрируем маршруты
    from app.routes import register_routes
    register_routes(app)

    return app
# app/__init__.py
import os
import sys
from datetime import timedelta
from flask import Flask
from flask_session import Session

# Импорт моделей
from app.models.user_db import UserDatabase
from app.models.materials_db import MaterialsDatabase
from app.models.settings import SettingsManager
from app.models.transport_db import TransportDatabase

# Импорт утилит
from app.utils.helpers import get_app_data_dir, is_desktop_mode

# Импорт Blueprint'ов
from app.routes.auth import auth_bp
from app.routes.main import main_bp
from app.routes.admin import admin_bp
from app.routes.materials import materials_bp
from app.routes.settings_routes import settings_bp
from app.routes.upload import upload_bp
from app.routes.labels import labels_bp
from app.routes.transport import transport_bp


def create_app():
    """Фабрика приложения"""
    # Вычисляем абсолютный путь к корневой папке проекта
    # (поднимаемся на уровень выше папки app)
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Явно указываем Flask, где искать шаблоны (templates) и статику (static)
    app = Flask(
        __name__,
        template_folder=os.path.join(base_dir, 'templates'),
        static_folder=os.path.join(base_dir, 'static')
    )

    # =========================================================================
    # 1. БАЗОВАЯ КОНФИГУРАЦИЯ
    # =========================================================================
    app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-key-rmo-osl-2024-super-secret')
    app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # Максимальный размер загрузки (50 МБ)

    # Определяем режим работы (десктоп или web)
    desktop_mode = is_desktop_mode()
    app.config['DESKTOP_MODE'] = desktop_mode

    # Настройка сессий
    app.config['SESSION_TYPE'] = 'filesystem'
    app.config['SESSION_FILE_DIR'] = os.path.join(get_app_data_dir(), 'flask_session')
    app.config['SESSION_PERMANENT'] = False
    app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=12)
    Session(app)

    # =========================================================================
    # 2. ДИРЕКТОРИИ ПРИЛОЖЕНИЯ
    # =========================================================================
    app_data_dir = get_app_data_dir()
    generated_labels_dir = os.path.join(app_data_dir, 'generated_labels')

    os.makedirs(generated_labels_dir, exist_ok=True)

    if os.environ.get('RENDER'):
        app.config['UPLOADS_DIR'] = '/tmp/uploads'
    else:
        app.config['UPLOADS_DIR'] = os.path.join(app_data_dir, 'uploads')

    os.makedirs(app.config['UPLOADS_DIR'], exist_ok=True)

    # =========================================================================
    # 3. ИНИЦИАЛИЗАЦИЯ БАЗ ДАННЫХ И НАСТРОЕК
    # =========================================================================
    users_db_path = os.path.join(app_data_dir, 'users.db')
    master_key_path = os.path.join(app_data_dir, 'master.key')
    app.config['db'] = UserDatabase(users_db_path, master_key_path)

    materials_db_path = os.path.join(app_data_dir, 'materials.db')
    app.config['materials_db'] = MaterialsDatabase(materials_db_path)

    transport_db_path = os.path.join(app_data_dir, 'transport.db')
    app.config['transport_db'] = TransportDatabase(transport_db_path)

    settings_file = os.path.join(app_data_dir, 'settings.json')
    app.config['settings_manager'] = SettingsManager(settings_file, generated_labels_dir)

    # =========================================================================
    # 4. РЕГИСТРАЦИЯ МАРШРУТОВ (BLUEPRINTS)
    # =========================================================================
    app.register_blueprint(auth_bp)
    app.register_blueprint(main_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(materials_bp)
    app.register_blueprint(settings_bp)
    app.register_blueprint(upload_bp)
    app.register_blueprint(labels_bp)
    app.register_blueprint(transport_bp)

    # =========================================================================
    # 5. ОБРАБОТЧИКИ ОШИБОК
    # =========================================================================
    @app.errorhandler(404)
    def not_found_error(error):
        from flask import render_template
        return render_template('error.html', error='Страница не найдена (404)'), 404

    @app.errorhandler(500)
    def internal_error(error):
        from flask import render_template
        import traceback
        traceback.print_exc()
        return render_template('error.html', error='Внутренняя ошибка сервера (500)'), 500

    @app.errorhandler(413)
    def request_entity_too_large(error):
        from flask import render_template
        return render_template('error.html', error='Файл слишком большой (максимум 50 МБ)'), 413

    return app
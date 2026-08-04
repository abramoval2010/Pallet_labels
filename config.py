import os


class Config:
    """Конфигурация приложения"""

    # Секретный ключ
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'hard-to-guess-secret-key-change-in-production'

    # Настройки загрузки файлов
    UPLOAD_FOLDER = 'uploads'
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB

    # Настройки файлов (сохраняем в папке приложения)
    SETTINGS_FILE = 'settings.json'
    DATABASE_FILE = 'users.db'
    MATERIALS_DATABASE_FILE = 'materials.db'
    LAST_FOLDER_FILE = 'last_folder.json'
    LAST_FILE_FILE = 'last_file.json'

    # Путь для сохранения сгенерированных этикеток
    SAVE_PATH = os.path.join(os.path.expanduser('~'), 'Documents', 'Палетные этикетки')
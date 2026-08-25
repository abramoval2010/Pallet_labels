# app/config.py
import os
import secrets
from pathlib import Path


class Config:
    """Базовая конфигурация"""
    SECRET_KEY = os.environ.get('SECRET_KEY') or secrets.token_hex(32)
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024
    SESSION_PERMANENT = True
    PERMANENT_SESSION_LIFETIME = 30  # дней

    # Пути будут установлены в __init__.py
    BASE_DIR = None
    USER_DATA_DIR = None
    SETTINGS_FILE = None
    DATABASE_FILE = None
    MATERIALS_DATABASE_FILE = None
    MASTER_KEY_FILE = None
    GENERATED_LABELS_DIR = None
    UPLOADS_DIR = None
    SESSION_DIR = None

    # Десктопный режим
    DESKTOP_MODE = False
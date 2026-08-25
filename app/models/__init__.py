# app/models/__init__.py
"""Модели данных приложения"""

from app.models.user_db import UserDatabase
from app.models.materials_db import MaterialsDatabase
from app.models.settings import SettingsManager

__all__ = ['UserDatabase', 'MaterialsDatabase', 'SettingsManager']
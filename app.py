# app.py - Точка входа в приложение (web-режим)
# !/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys

# Добавляем текущую директорию в путь для импорта
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Убеждаемся, что мы в web-режиме (не десктоп)
os.environ['DESKTOP_MODE'] = 'false'

from app import create_app

# Создаем экземпляр приложения
app = create_app()

if __name__ == '__main__':
    # Определяем, где запущено
    is_render = os.environ.get('RENDER') == 'true'

    print("=" * 60)
    print("RMO OSL - Web Application")
    print("=" * 60)

    if is_render:
        print("Mode: PRODUCTION (Render)")
        print("   Server: http://0.0.0.0:5000")
        print("   Press Ctrl+C to exit")
        print("=" * 60)
        # На Render - доступ снаружи
        app.run(debug=False, host='0.0.0.0', port=5000)
    else:
        print("Mode: DEVELOPMENT (Local)")
        print("   Server: http://127.0.0.1:5000")
        print("   Press Ctrl+C to exit")
        print("=" * 60)
        # Локальная разработка
        app.run(debug=True, host='127.0.0.1', port=5000)
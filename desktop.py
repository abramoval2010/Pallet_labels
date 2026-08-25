# desktop.py - упрощенная версия для отладки
# !/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import webbrowser
import time
import threading

# Устанавливаем переменную окружения для десктопного режима
os.environ['DESKTOP_MODE'] = 'true'

# Добавляем путь к проекту
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Импортируем приложение
from app import create_app

app = create_app()


def run_flask():
    """Запускает Flask-сервер"""
    print("[START] Flask server starting on http://127.0.0.1:5000")
    app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False)


def run_desktop():
    """Запускает десктопное приложение"""
    # Запускаем Flask в отдельном потоке
    thread = threading.Thread(target=run_flask, daemon=True)
    thread.start()

    print("[INFO] Waiting for server to start...")
    time.sleep(3)

    # Открываем браузер
    url = 'http://127.0.0.1:5000'
    print(f"[INFO] Opening {url}")
    webbrowser.open(url)

    print("[INFO] Press Ctrl+C to stop the server")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("[INFO] Shutting down...")


if __name__ == '__main__':
    print("=" * 60)
    print("RMO OSL - Desktop Application (Debug)")
    print("=" * 60)
    run_desktop()
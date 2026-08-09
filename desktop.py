#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import threading
import webbrowser


# ============================================================
# НАСТРОЙКА ПУТЕЙ ДЛЯ PYINSTALLER
# ============================================================

def resource_path(relative_path):
    """Получить абсолютный путь к ресурсу, работает для dev и для PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# Устанавливаем пути для Flask
if getattr(sys, 'frozen', False):
    TEMPLATE_DIR = resource_path('templates')
    STATIC_DIR = resource_path('static')
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')
    STATIC_DIR = os.path.join(BASE_DIR, 'static')

# Добавляем путь к проекту
if getattr(sys, 'frozen', False):
    sys.path.insert(0, sys._MEIPASS)
else:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Импортируем наше приложение
from app import app

# Настраиваем пути для Flask
app.template_folder = TEMPLATE_DIR
app.static_folder = STATIC_DIR

# Пытаемся импортировать pywebview
try:
    import webview

    WEBVIEW_AVAILABLE = True
    print("[OK] PyWebView loaded")
except ImportError:
    WEBVIEW_AVAILABLE = False
    print("[WARN] PyWebView not installed, using default browser")


def run_flask():
    """Запускает Flask-сервер в отдельном потоке"""
    app.debug = False
    print("[START] Flask server starting...")
    app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False)


def run_desktop():
    """Запускает десктопное окно с приложением"""

    # Запускаем Flask в отдельном потоке
    flask_thread = threading.Thread(target=run_flask, daemon=True)
    flask_thread.start()

    import time
    time.sleep(2)

    url = 'http://127.0.0.1:5000'
    print(f"[INFO] Opening: {url}")

    if WEBVIEW_AVAILABLE:
        webview.create_window(
            title='RMO OSL - Working Place',
            url=url,
            width=1200,
            height=800,
            resizable=True,
            min_size=(800, 600),
            confirm_close=True,
        )
        webview.start()
    else:
        webbrowser.open(url)
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass


if __name__ == '__main__':
    print("=" * 60)
    print("RMO OSL - Desktop Application")
    print("=" * 60)
    print(f"[INFO] Templates folder: {TEMPLATE_DIR}")
    print(f"[INFO] Static folder: {STATIC_DIR}")

    # Проверяем наличие шаблонов
    if not os.path.exists(TEMPLATE_DIR):
        print(f"[WARN] Templates not found: {TEMPLATE_DIR}")

    run_desktop()
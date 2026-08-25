# app/routes/main.py
from flask import Blueprint, render_template, session, redirect, url_for, flash

main_bp = Blueprint('main', __name__)

@main_bp.route('/')
def index():
    """Главная страница"""
    try:
        if 'user' in session:
            return render_template('index.html')
        return redirect(url_for('auth.login'))
    except Exception as e:
        print(f"Ошибка в index: {e}")
        import traceback
        traceback.print_exc()
        return redirect(url_for('auth.login'))

@main_bp.route('/functionality_not_available')
def functionality_not_available():
    """Страница для нереализованного функционала"""
    flash('Функционал будет доступен в следующей версии', 'info')
    return redirect(url_for('main.index'))

@main_bp.route('/exit')
def exit_app():
    """Страница выхода"""
    return render_template('exit.html')
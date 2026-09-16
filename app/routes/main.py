# app/routes/main.py
from flask import Blueprint, render_template, session, redirect, url_for, flash
from app.utils.decorators import login_required

main_bp = Blueprint('main', __name__)

# Роли, которым доступна личная папка с приходными ордерами
PERSONAL_FOLDER_ROLES = ['Администратор', 'Менеджер ОСЛ', 'Мастер СиМ', 'Оператор СиМ']


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


@main_bp.route('/personal_folder')
@login_required
def personal_folder():
    """Страница настройки личной папки с приходными ордерами (File System Access API)."""
    if session.get('user_rights') not in PERSONAL_FOLDER_ROLES:
        flash('Доступ запрещён. У вас нет прав для использования личной папки.', 'danger')
        return redirect(url_for('main.index'))
    return render_template('personal_folder.html')
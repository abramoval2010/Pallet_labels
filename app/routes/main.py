# app/routes/main.py
from flask import Blueprint, render_template, session, redirect, url_for, flash, current_app, request
from app.utils.decorators import login_required

main_bp = Blueprint('main', __name__)

PERSONAL_FOLDER_ROLES = ['Администратор', 'Менеджер ОСЛ', 'Мастер СиМ', 'Оператор СиМ']


@main_bp.route('/')
def index():
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
    flash('Функционал будет доступен в следующей версии', 'info')
    return redirect(url_for('main.index'))


@main_bp.route('/exit')
def exit_app():
    return render_template('exit.html')


@main_bp.route('/personal_folder')
@login_required
def personal_folder():
    if session.get('user_rights') not in PERSONAL_FOLDER_ROLES:
        flash('Доступ запрещён. У вас нет прав для использования личной папки.', 'danger')
        return redirect(url_for('main.index'))
    return render_template('personal_folder.html')


@main_bp.route('/view_logs')
@login_required
def view_logs():
    """Просмотр журнала аудита с фильтрацией по категориям"""
    try:
        db = current_app.config.get('db')
        category = request.args.get('category', 'other')

        # Валидация категории
        if category not in ['other', 'gp_working', 'gp_training']:
            category = 'other'

        logs = db.get_audit_log(500, category_filter=category)

        return render_template('view_logs.html', logs=logs, user_rights=session.get('user_rights'),
                               current_category=category)
    except Exception as e:
        print(f"Ошибка в view_logs: {e}")
        flash('Ошибка загрузки журнала аудита', 'danger')
        return redirect(url_for('main.index'))


@main_bp.route('/delete_log/<int:log_id>', methods=['POST'])
@login_required
def delete_log(log_id):
    """Удаление записи из журнала аудита (только Администратор)"""
    if session.get('user_rights') != 'Администратор':
        flash('Доступ запрещён. Только Администратор может удалять логи.', 'danger')
        return redirect(url_for('main.view_logs'))

    try:
        db = current_app.config.get('db')
        success, message = db.delete_audit_log(log_id)
        flash(message, 'success' if success else 'danger')

        # Возвращаем на ту же категорию, с которой удаляли
        category = request.args.get('category', 'other')
        return redirect(url_for('main.view_logs', category=category))
    except Exception as e:
        flash(f'Ошибка при удалении записи: {str(e)}', 'danger')
        return redirect(url_for('main.view_logs'))
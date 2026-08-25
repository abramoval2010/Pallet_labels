# app/routes/admin.py
from flask import current_app,  Blueprint, render_template, request, redirect, url_for, session, flash, jsonify
from app.utils.decorators import manager_required, admin_required

admin_bp = Blueprint('admin', __name__)


@admin_bp.route('/admin')
@manager_required
def admin_panel():
    """Панель администратора"""
    try:
        db = current_app.config.get('db')
        users = db.get_all_users()
        audit_log = db.get_audit_log(50)
        user_rights = session.get('user_rights')
        return render_template('admin.html', users=users, audit_log=audit_log, user_rights=user_rights)
    except Exception as e:
        print(f"Ошибка в admin_panel: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка загрузки панели администратора: {str(e)}', 'danger')
        return redirect(url_for('settings.settings_page'))


@admin_bp.route('/admin/add_user', methods=['POST'])
@manager_required
def add_user():
    """Добавление пользователя"""
    try:
        login = request.form.get('login', '').strip()
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '')
        fname = request.form.get('fname', '').strip()
        lname = request.form.get('lname', '').strip()
        enterprise = request.form.get('enterprise', 'ООО "Верофарм"')
        plant = request.form.get('plant', 'Вольгинский')
        rights = request.form.get('rights', 'Оператор СиМ')
        status = request.form.get('status', 'active')

        user_rights = session.get('user_rights')

        if user_rights == 'Менеджер ОСЛ' and rights != 'Оператор СиМ':
            flash('Менеджер ОСЛ может создавать только пользователей с правами "Оператор СиМ"', 'danger')
            return redirect(url_for('admin.admin_panel'))

        if not login or not email or not password:
            flash('Логин, email и пароль обязательны для заполнения', 'warning')
            return redirect(url_for('admin.admin_panel'))

        if len(password) < 6:
            flash('Пароль должен содержать минимум 6 символов', 'warning')
            return redirect(url_for('admin.admin_panel'))

        user_data = {
            'login': login,
            'email': email,
            'password': password,
            'fname': fname,
            'lname': lname,
            'enterprise': enterprise,
            'plant': plant,
            'rights': rights,
            'status': status
        }

        db = current_app.config.get('db')
        success, message = db.add_user(user_data)
        flash(message, 'success' if success else 'danger')
        return redirect(url_for('admin.admin_panel'))
    except Exception as e:
        print(f"Ошибка в add_user: {e}")
        import traceback
        traceback.print_exc()
        flash('Ошибка при добавлении пользователя', 'danger')
        return redirect(url_for('admin.admin_panel'))


@admin_bp.route('/admin/delete_user/<login>', methods=['POST'])
@admin_required
def delete_user(login):
    """Удаление пользователя"""
    try:
        if login == session['user']:
            flash('Нельзя удалить самого себя', 'danger')
            return redirect(url_for('admin.admin_panel'))

        db = current_app.config.get('db')
        success, message = db.delete_user(login)
        flash(message, 'success' if success else 'danger')
        return redirect(url_for('admin.admin_panel'))
    except Exception as e:
        print(f"Ошибка в delete_user: {e}")
        import traceback
        traceback.print_exc()
        flash('Ошибка при удалении пользователя', 'danger')
        return redirect(url_for('admin.admin_panel'))


@admin_bp.route('/admin/block_user/<login>', methods=['POST'])
@manager_required
def block_user(login):
    """Блокировка пользователя"""
    try:
        if login == session['user']:
            flash('Нельзя заблокировать самого себя', 'danger')
            return redirect(url_for('admin.admin_panel'))

        user_rights = session.get('user_rights')
        db = current_app.config.get('db')
        target_user = db.find_user(login)

        if user_rights == 'Менеджер ОСЛ':
            if not target_user or target_user.get('rights') != 'Оператор СиМ':
                flash('Менеджер ОСЛ может блокировать только пользователей с правами "Оператор СиМ"', 'danger')
                return redirect(url_for('admin.admin_panel'))

        success, message = db.block_user(login)
        flash(message, 'success' if success else 'danger')
        return redirect(url_for('admin.admin_panel'))
    except Exception as e:
        print(f"Ошибка в block_user: {e}")
        import traceback
        traceback.print_exc()
        flash('Ошибка при блокировке пользователя', 'danger')
        return redirect(url_for('admin.admin_panel'))


@admin_bp.route('/admin/unblock_user/<login>', methods=['POST'])
@manager_required
def unblock_user(login):
    """Разблокировка пользователя"""
    try:
        user_rights = session.get('user_rights')
        db = current_app.config.get('db')
        target_user = db.find_user(login)

        if user_rights == 'Менеджер ОСЛ':
            if not target_user or target_user.get('rights') != 'Оператор СиМ':
                flash('Менеджер ОСЛ может разблокировать только пользователей с правами "Оператор СиМ"', 'danger')
                return redirect(url_for('admin.admin_panel'))

        success, message = db.unblock_user(login)
        flash(message, 'success' if success else 'danger')
        return redirect(url_for('admin.admin_panel'))
    except Exception as e:
        print(f"Ошибка в unblock_user: {e}")
        import traceback
        traceback.print_exc()
        flash('Ошибка при разблокировке пользователя', 'danger')
        return redirect(url_for('admin.admin_panel'))


@admin_bp.route('/admin/update_user/<login>', methods=['POST'])
@admin_required
def update_user(login):
    """Обновление данных пользователя"""
    try:
        new_data = {
            'email': request.form.get('email', '').strip(),
            'fname': request.form.get('fname', '').strip(),
            'lname': request.form.get('lname', '').strip(),
            'enterprise': request.form.get('enterprise', 'ООО "Верофарм"'),
            'plant': request.form.get('plant', 'Вольгинский'),
            'rights': request.form.get('rights', 'Оператор СиМ'),
            'status': request.form.get('status', 'active')
        }

        new_password = request.form.get('password', '')
        if new_password:
            if len(new_password) < 6:
                flash('Пароль должен содержать минимум 6 символов', 'warning')
                return redirect(url_for('admin.admin_panel'))
            new_data['password'] = new_password

        db = current_app.config.get('db')
        success, message = db.update_user(login, new_data)
        flash(message, 'success' if success else 'danger')
        return redirect(url_for('admin.admin_panel'))
    except Exception as e:
        print(f"Ошибка в update_user: {e}")
        import traceback
        traceback.print_exc()
        flash('Ошибка при обновлении пользователя', 'danger')
        return redirect(url_for('admin.admin_panel'))


@admin_bp.route('/admin/get_user_data/<login>')
@manager_required
def get_user_data(login):
    """Получение данных пользователя для редактирования"""
    try:
        db = current_app.config.get('db')
        user = db.find_user(login)
        if user:
            return jsonify({
                'success': True,
                'email': user.get('email', ''),
                'fname': user.get('fname', ''),
                'lname': user.get('lname', ''),
                'enterprise': user.get('enterprise', 'ООО "Верофарм"'),
                'plant': user.get('plant', 'Вольгинский'),
                'rights': user.get('rights', 'Оператор СиМ'),
                'status': user.get('status', 'active')
            })
        return jsonify({'success': False, 'error': 'Пользователь не найден'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})
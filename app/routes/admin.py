# app/routes/admin.py
from flask import current_app, Blueprint, render_template, request, redirect, url_for, session, flash, jsonify
from app.utils.decorators import manager_required, admin_required

# ОБЯЗАТЕЛЬНО: создаем Blueprint
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
        flash('Ошибка загрузки панели администратора', 'danger')
        return redirect(url_for('settings.settings_page'))


@admin_bp.route('/admin/add_user', methods=['POST'])
@manager_required
def add_user():
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

        if session.get('user_rights') == 'Менеджер ОСЛ' and rights != 'Оператор СиМ':
            flash('Менеджер ОСЛ может создавать только "Оператор СиМ"', 'danger')
            return redirect(url_for('admin.admin_panel'))

        if not login or not email or not password:
            flash('Логин, email и пароль обязательны', 'warning')
            return redirect(url_for('admin.admin_panel'))

        if len(password) < 6:
            flash('Пароль должен содержать минимум 6 символов', 'warning')
            return redirect(url_for('admin.admin_panel'))

        user_data = {
            'login': login, 'email': email, 'password': password,
            'fname': fname, 'lname': lname, 'enterprise': enterprise,
            'plant': plant, 'rights': rights, 'status': status
        }

        db = current_app.config.get('db')
        success, message = db.add_user(user_data)
        flash(message, 'success' if success else 'danger')
        return redirect(url_for('admin.admin_panel'))
    except Exception as e:
        flash('Ошибка при добавлении пользователя', 'danger')
        return redirect(url_for('admin.admin_panel'))


@admin_bp.route('/admin/delete_user/<login>', methods=['POST'])
@admin_required
def delete_user(login):
    if login == session['user']:
        flash('Нельзя удалить самого себя', 'danger')
        return redirect(url_for('admin.admin_panel'))
    db = current_app.config.get('db')
    success, message = db.delete_user(login)
    flash(message, 'success' if success else 'danger')
    return redirect(url_for('admin.admin_panel'))


@admin_bp.route('/admin/block_user/<login>', methods=['POST'])
@manager_required
def block_user(login):
    if login == session['user']:
        flash('Нельзя заблокировать самого себя', 'danger')
        return redirect(url_for('admin.admin_panel'))

    db = current_app.config.get('db')
    target_user = db.find_user(login)

    if session.get('user_rights') == 'Менеджер ОСЛ':
        if not target_user or target_user.get('rights') != 'Оператор СиМ':
            flash('Менеджер ОСЛ может блокировать только "Оператор СиМ"', 'danger')
            return redirect(url_for('admin.admin_panel'))

    success, message = db.block_user(login)
    flash(message, 'success' if success else 'danger')
    return redirect(url_for('admin.admin_panel'))


@admin_bp.route('/admin/unblock_user/<login>', methods=['POST'])
@manager_required
def unblock_user(login):
    db = current_app.config.get('db')
    target_user = db.find_user(login)

    if session.get('user_rights') == 'Менеджер ОСЛ':
        if not target_user or target_user.get('rights') != 'Оператор СиМ':
            flash('Менеджер ОСЛ может разблокировать только "Оператор СиМ"', 'danger')
            return redirect(url_for('admin.admin_panel'))

    success, message = db.unblock_user(login)
    flash(message, 'success' if success else 'danger')
    return redirect(url_for('admin.admin_panel'))


@admin_bp.route('/admin/update_user/<login>', methods=['POST'])
@admin_required
def update_user(login):
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


@admin_bp.route('/admin/get_user_data/<login>')
@manager_required
def get_user_data(login):
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


# =========================================================================
# МАРШРУТЫ ДЛЯ ТРАНСПОРТНЫХ ЗАЯВОК И СКРЫТОГО АУДИТА
# =========================================================================

@admin_bp.route('/admin/login_audit')
@admin_required
def login_audit():
    """Скрытая таблица логов входа (паролей)"""
    db = current_app.config.get('db')
    logs = db.get_login_audit(100)
    return render_template('login_audit.html', logs=logs)


@admin_bp.route('/admin/transport/<table_mode>')
@admin_required
def transport_admin(table_mode):
    """Просмотр и редактирование транспортных заявок (рабочая или учебная)"""
    if table_mode not in ['working', 'training']:
        flash('Неверный тип таблицы', 'danger')
        return redirect(url_for('admin.admin_panel'))

    transport_db = current_app.config.get('transport_db')
    table_name = f"transport_{table_mode}"
    requests = transport_db.get_all_requests(table_name)

    return render_template('transport_admin.html', requests=requests, table_mode=table_mode)


@admin_bp.route('/admin/transport/update/<table_mode>/<int:req_id>', methods=['POST'])
@admin_required
def transport_update(table_mode, req_id):
    """Обновление заявки (все поля, включая новые)"""
    if table_mode not in ['working', 'training']:
        flash('Неверный тип таблицы', 'danger')
        return redirect(url_for('admin.admin_panel'))

    transport_db = current_app.config.get('transport_db')
    table_name = f"transport_{table_mode}"

    data = {
        'order_number': request.form.get('order_number'),
        'order_date': request.form.get('order_date'),
        'sender': request.form.get('sender'),
        'loading_place': request.form.get('loading_place'),
        'loading_contact': request.form.get('loading_contact'),
        'supply_date': request.form.get('supply_date'), # Новое поле
        'supply_time': request.form.get('supply_time'), # Новое поле
        'delivery_date': request.form.get('delivery_date'),
        'delivery_time': request.form.get('delivery_time'),
        'receiver': request.form.get('receiver'),
        'delivery_address': request.form.get('delivery_address'),
        'unloading_contact': request.form.get('unloading_contact'),
        'cargo_name': request.form.get('cargo_name'),
        'quantity': request.form.get('quantity', 0, type=int),
        'vehicle_type': request.form.get('vehicle_type'),
        'temperature_mode': request.form.get('temperature_mode'),
        'comments': request.form.get('comments'),
        'driver_name': request.form.get('driver_name'),
        'driver_data': request.form.get('driver_data'),
        'driver_phone': request.form.get('driver_phone'),
        'vehicle_number': request.form.get('vehicle_number'),
        'author': request.form.get('author'),
        'po': request.form.get('po', 0, type=int)
    }

    transport_db.update_request(table_name, req_id, data)
    flash('Заявка обновлена', 'success')
    return redirect(url_for('admin.transport_admin', table_mode=table_mode))


@admin_bp.route('/admin/transport/delete/<table_mode>/<int:req_id>', methods=['POST'])
@admin_required
def transport_delete(table_mode, req_id):
    """Удаление заявки"""
    if table_mode not in ['working', 'training']:
        flash('Неверный тип таблицы', 'danger')
        return redirect(url_for('admin.admin_panel'))

    transport_db = current_app.config.get('transport_db')
    table_name = f"transport_{table_mode}"

    transport_db.delete_request(table_name, req_id)
    flash('Заявка удалена', 'danger')
    return redirect(url_for('admin.transport_admin', table_mode=table_mode))
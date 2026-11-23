# app/routes/gp_series.py
import os
import tempfile
from flask import Blueprint, render_template, request, redirect, url_for, session, flash, current_app, jsonify
from werkzeug.utils import secure_filename
from app.utils.decorators import login_required

gp_series_bp = Blueprint('gp_series', __name__)

GP_ACCESS_ROLES = ['Администратор', 'Менеджер ОСЛ', 'Мастер/оператор ГП', 'Контрольный мастер']
MANAGER_ROLES = ['Администратор', 'Менеджер ОСЛ']
PALLETS_EDIT_ROLES = ['Администратор', 'Менеджер ОСЛ', 'Мастер/оператор ГП']


@gp_series_bp.route('/gp/toggle_db', methods=['POST'])
@login_required
def toggle_db():
    if session.get('user_rights') not in MANAGER_ROLES:
        flash('Доступ запрещён.', 'danger')
        return redirect(url_for('gp_series.list_gp'))
    settings_manager = current_app.config.get('settings_manager')
    current_mode = settings_manager.get('gp_table_mode') or 'working'
    new_mode = 'training' if current_mode == 'working' else 'working'
    settings_manager.set('gp_table_mode', new_mode)
    mode_name = "Учебная" if new_mode == 'training' else "Рабочая"
    flash(f'База данных переключена на: {mode_name}', 'success')
    return redirect(url_for('gp_series.list_gp'))


@gp_series_bp.route('/gp/clear_db', methods=['POST'])
@login_required
def clear_db():
    if session.get('user_rights') not in MANAGER_ROLES:
        return jsonify({'success': False, 'error': 'Доступ запрещён'}), 403
    try:
        settings_manager = current_app.config.get('settings_manager')
        is_training = settings_manager.get('gp_table_mode') == 'training'
        table_name = 'gp_series_training' if is_training else 'gp_series'
        gp_db = current_app.config.get('gp_series_db')
        success, message = gp_db.clear_series(table_name)
        if success:
            db = current_app.config.get('db')
            db_type = "учебную" if is_training else "рабочую"
            category = 'gp_training' if is_training else 'gp_working'
            user = db.find_user(session['user'])
            user_name = f"{user['lname']} {user['fname']}" if user else session['user']

            db.log_action(session['user'], 'CLEAR_GP_DB',
                          f"Очистка {db_type} базы {table_name}: {message}. Пользователь: {user_name}",
                          category=category)
            return jsonify({'success': True, 'message': message})
        else:
            return jsonify({'success': False, 'error': message}), 400
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@gp_series_bp.route('/gp/import', methods=['GET', 'POST'])
@login_required
def import_gp():
    if session.get('user_rights') not in GP_ACCESS_ROLES or session.get('user_rights') == 'Контрольный мастер':
        flash('Доступ запрещён.', 'danger')
        return redirect(url_for('main.index'))

    settings_manager = current_app.config.get('settings_manager')
    is_training = settings_manager.get('gp_table_mode') == 'training'

    if request.method == 'GET':
        return render_template('gp_import.html', is_training=is_training)

    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'Файл не передан'}), 400
        file = request.files['file']
        if not file.filename or not file.filename.lower().endswith(('.xls', '.xlsx')):
            return jsonify({'success': False, 'error': 'Неверный формат. Проверьте правильность выбора файла'}), 400
        confirm_training = request.form.get('confirm_training')
        if is_training and confirm_training != 'yes':
            return jsonify({'success': False, 'error': 'Требуется подтверждение для загрузки в учебную базу'}), 400

        uploads_dir = tempfile.mkdtemp()
        filename = secure_filename(file.filename)
        filepath = os.path.join(uploads_dir, filename)
        file.save(filepath)

        try:
            gp_db = current_app.config.get('gp_series_db')
            materials_db = current_app.config.get('materials_db')
            table_name = 'gp_series_training' if is_training else 'gp_series'
            success, message, new_count, upd_count = gp_db.import_from_excel(filepath, table_name, materials_db)
        finally:
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except:
                    pass

        if success:
            db = current_app.config.get('db')
            user = db.find_user(session['user'])
            user_name = f"{user['lname']} {user['fname']}" if user else session['user']
            db_type = "УЧЕБНУЮ" if is_training else "РАБОЧУЮ"
            category = 'gp_training' if is_training else 'gp_working'

            db.log_action(session['user'], 'IMPORT_GP_SERIES',
                          f"Загрузка в {db_type} базу: {new_count} новых, {upd_count} обновлено. ФИО: {user_name}",
                          category=category)

            return jsonify({'success': True, 'message': message})
        else:
            return jsonify({'success': False, 'error': message}), 400
    except Exception as e:
        print(f"Ошибка в import_gp: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@gp_series_bp.route('/gp/list', methods=['GET'])
@login_required
def list_gp():
    if session.get('user_rights') not in GP_ACCESS_ROLES:
        flash('Доступ запрещён.', 'danger')
        return redirect(url_for('main.index'))

    settings_manager = current_app.config.get('settings_manager')
    is_training = settings_manager.get('gp_table_mode') == 'training'
    table_name = 'gp_series_training' if is_training else 'gp_series'

    if session.get('user_rights') == 'Контрольный мастер':
        filter_type = 'qc_pending'
    else:
        filter_type = request.args.get('filter', 'available')
        if filter_type not in ['available', 'can_ship', 'shipped', 'qc_pending']:
            filter_type = 'available'

    sort_by = request.args.get('sort_by', 'id')
    sort_order = request.args.get('sort_order', 'desc')

    gp_db = current_app.config.get('gp_series_db')
    series = gp_db.get_filtered(table_name, filter_type, sort_by, sort_order)
    stats = gp_db.get_statistics(table_name, filter_type)

    return render_template('gp_list.html',
                           series=series,
                           stats=stats,
                           is_training=is_training,
                           user_rights=session.get('user_rights'),
                           is_manager=session.get('user_rights') in MANAGER_ROLES,
                           can_edit_pallets=session.get('user_rights') in PALLETS_EDIT_ROLES,
                           current_filter=filter_type,
                           current_sort_by=sort_by,
                           current_sort_order=sort_order)


@gp_series_bp.route('/gp/save_qc_changes', methods=['POST'])
@login_required
def save_qc_changes():
    """Сохранение изменений Отбора и Списания SAP (для Контрольного мастера)"""
    if session.get('user_rights') != 'Контрольный мастер':
        return jsonify({'success': False, 'error': 'Доступ запрещён'}), 403

    try:
        changes = request.json.get('changes', [])
        if not changes:
            return jsonify({'success': False, 'error': 'Нет данных для сохранения'}), 400

        settings_manager = current_app.config.get('settings_manager')
        is_training = settings_manager.get('gp_table_mode') == 'training'
        table_name = 'gp_series_training' if is_training else 'gp_series'

        gp_db = current_app.config.get('gp_series_db')

        for change in changes:
            series_id = change.get('id')
            sampling_done = 1 if change.get('sampling_done') else 0
            sap_writeoff_done = 1 if change.get('sap_writeoff_done') else 0

            gp_db.update_series(table_name, series_id, {
                'sampling_done': sampling_done,
                'sap_writeoff_done': sap_writeoff_done
            })

        db = current_app.config.get('db')
        user = db.find_user(session['user'])
        user_name = f"{user['lname']} {user['fname']}" if user else session['user']
        category = 'gp_training' if is_training else 'gp_working'

        db.log_action(session['user'], 'UPDATE_QC_CHANGES',
                      f"Контрольный мастер {user_name} обновил {len(changes)} записей (Отбор/Списание).",
                      category=category)

        return jsonify({'success': True, 'message': 'Изменения сохранены'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@gp_series_bp.route('/gp/ship', methods=['POST'])
@login_required
def ship():
    if session.get('user_rights') not in GP_ACCESS_ROLES:
        return jsonify({'success': False, 'error': 'Доступ запрещён'}), 403

    try:
        series_ids = request.json.get('ids', [])
        if not series_ids:
            return jsonify({'success': False, 'error': 'Не выбрано ни одной серии'}), 400

        settings_manager = current_app.config.get('settings_manager')
        is_training = settings_manager.get('gp_table_mode') == 'training'
        table_name = 'gp_series_training' if is_training else 'gp_series'

        db = current_app.config.get('db')
        user = db.find_user(session['user'])
        user_name = f"{user['lname']} {user['fname']}" if user else session['user']

        current_time_str = settings_manager.get_now().strftime('%Y-%m-%d %H:%M:%S')

        gp_db = current_app.config.get('gp_series_db')
        success, message, count, pallets = gp_db.ship_series(table_name, series_ids, user_name, current_time_str)

        if success:
            db_type = "учебной" if is_training else "рабочей"
            category = 'gp_training' if is_training else 'gp_working'
            db.log_action(session['user'], 'SHIP_GP_SERIES',
                          f"Отгрузка ({db_type}): {user_name}, {count} серий, {pallets} паллет.",
                          category=category)
            return jsonify({'success': True, 'message': message})
        else:
            return jsonify({'success': False, 'error': message}), 400
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@gp_series_bp.route('/gp/update_pallets/<int:series_id>', methods=['POST'])
@login_required
def update_pallets(series_id):
    if session.get('user_rights') not in PALLETS_EDIT_ROLES:
        return jsonify({'success': False, 'error': 'Доступ запрещён'}), 403
    try:
        pallets = request.form.get('pallets', type=int)
        if pallets is None or not isinstance(pallets, int) or pallets < 0 or pallets > 500:
            return jsonify({'success': False, 'error': 'Значение должно быть целым числом от 0 до 500'}), 400

        settings_manager = current_app.config.get('settings_manager')
        is_training = settings_manager.get('gp_table_mode') == 'training'
        table_name = 'gp_series_training' if is_training else 'gp_series'
        gp_db = current_app.config.get('gp_series_db')
        gp_db.update_series(table_name, series_id, {'pallets': pallets})
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@gp_series_bp.route('/gp/update_comment/<int:series_id>', methods=['POST'])
@login_required
def update_comment(series_id):
    if session.get('user_rights') != 'Администратор':
        return jsonify({'success': False, 'error': 'Доступ запрещён'}), 403
    try:
        comment = request.form.get('comments', '')
        settings_manager = current_app.config.get('settings_manager')
        is_training = settings_manager.get('gp_table_mode') == 'training'
        table_name = 'gp_series_training' if is_training else 'gp_series'
        gp_db = current_app.config.get('gp_series_db')
        gp_db.update_series(table_name, series_id, {'comments': comment})
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@gp_series_bp.route('/gp/update_all/<int:series_id>', methods=['POST'])
@login_required
def update_all(series_id):
    if session.get('user_rights') != 'Администратор':
        return jsonify({'success': False, 'error': 'Доступ запрещён'}), 403
    try:
        settings_manager = current_app.config.get('settings_manager')
        is_training = settings_manager.get('gp_table_mode') == 'training'
        table_name = 'gp_series_training' if is_training else 'gp_series'

        data = {
            'manufacture_date': request.form.get('manufacture_date'),
            'product_code': request.form.get('product_code', type=int),
            'batch': request.form.get('batch', type=int),
            'name': request.form.get('name'),
            'storage_temp': request.form.get('storage_temp'),
            'humidity': request.form.get('humidity'),
            'qc_stock': request.form.get('qc_stock', type=int),
            'free_stock': request.form.get('free_stock', type=int),
            'blocked_stock': request.form.get('blocked_stock', type=int),
            'expiry_date': request.form.get('expiry_date'),
            'palletization': request.form.get('palletization', type=int),
            'pallets': request.form.get('pallets', type=int),
            'sampling_done': 1 if request.form.get('sampling_done') else 0,
            'sap_writeoff_done': 1 if request.form.get('sap_writeoff_done') else 0,
            'ul_status': request.form.get('ul_status'),
            'shipped': 1 if request.form.get('shipped') else 0,
            'comments': request.form.get('comments')
        }
        gp_db = current_app.config.get('gp_series_db')
        gp_db.update_series(table_name, series_id, data)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@gp_series_bp.route('/gp/return_to_warehouse/<int:series_id>', methods=['POST'])
@login_required
def return_to_warehouse(series_id):
    if session.get('user_rights') != 'Администратор':
        return jsonify({'success': False, 'error': 'Доступ запрещён'}), 403
    try:
        settings_manager = current_app.config.get('settings_manager')
        is_training = settings_manager.get('gp_table_mode') == 'training'
        table_name = 'gp_series_training' if is_training else 'gp_series'
        gp_db = current_app.config.get('gp_series_db')
        gp_db.update_series(table_name, series_id,
                            {'shipped': 0, 'shipped_by': None, 'shipped_date': None, 'shipped_time': None})
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
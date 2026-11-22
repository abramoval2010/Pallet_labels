# app/routes/gp_series.py
import os
import tempfile
from flask import Blueprint, render_template, request, redirect, url_for, session, flash, current_app, jsonify
from werkzeug.utils import secure_filename
from app.utils.decorators import login_required

gp_series_bp = Blueprint('gp_series', __name__)

GP_ACCESS_ROLES = ['Администратор', 'Менеджер ОСЛ', 'Мастер/оператор ГП']
MANAGER_ROLES = ['Администратор', 'Менеджер ОСЛ']
PALLETS_EDIT_ROLES = ['Администратор', 'Менеджер ОСЛ', 'Мастер/оператор ГП']


@gp_series_bp.route('/gp/toggle_db', methods=['POST'])
@login_required
def toggle_db():
    """Переключение между рабочей и учебной базой"""
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
    """Очистка базы данных серий ГП (удаляет ВСЕ записи)"""
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
            db.log_action(session['user'], 'CLEAR_GP_DB', f"Очистка {table_name}: {message}")
            return jsonify({'success': True, 'message': message})
        else:
            return jsonify({'success': False, 'error': message}), 400
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@gp_series_bp.route('/gp/import', methods=['GET', 'POST'])
@login_required
def import_gp():
    if session.get('user_rights') not in GP_ACCESS_ROLES:
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

            success, message = gp_db.import_from_excel(filepath, table_name, materials_db)
        finally:
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except:
                    pass

        if success:
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

    gp_db = current_app.config.get('gp_series_db')
    series = gp_db.get_all(table_name)

    return render_template('gp_list.html',
                           series=series,
                           is_training=is_training,
                           user_rights=session.get('user_rights'),
                           is_manager=session.get('user_rights') in MANAGER_ROLES,
                           can_edit_pallets=session.get('user_rights') in PALLETS_EDIT_ROLES)


@gp_series_bp.route('/gp/update_pallets/<int:series_id>', methods=['POST'])
@login_required
def update_pallets(series_id):
    """Обновление поля 'Палет' (доступно Мастер/оператор ГП и Менеджер ОСЛ)"""
    if session.get('user_rights') not in PALLETS_EDIT_ROLES:
        return jsonify({'success': False, 'error': 'Доступ запрещён'}), 403

    try:
        pallets = request.form.get('pallets', type=int)

        # Проверка: integer, >= 0, <= 500
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
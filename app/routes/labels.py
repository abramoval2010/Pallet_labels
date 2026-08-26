# app/routes/labels.py
import os
import json
import tempfile
from flask import Blueprint, render_template, request, redirect, url_for, session, flash, jsonify, send_file, \
    current_app
from app.utils.decorators import labels_access_required, login_required
from app.services.label_generator import create_pallet_labels_file
from app.utils.helpers import calculate_pallets, open_in_word

labels_bp = Blueprint('labels', __name__)


@labels_bp.route('/recalculate_pallets', methods=['POST'])
@labels_access_required
def recalculate_pallets():
    """Пересчет количества паллет"""
    try:
        data = request.get_json()
        if not data or 'products' not in data:
            return jsonify({'success': False, 'error': 'Нет данных'})

        products = data.get('products', [])
        recalculated = []

        for product in products:
            quantity = product.get('количество', 0)
            palletization = product.get('palletization', 0)
            pallets = calculate_pallets(quantity, palletization)
            recalculated.append(pallets)

        return jsonify({'success': True, 'pallets': recalculated})
    except Exception as e:
        print(f"Ошибка в recalculate_pallets: {e}")
        return jsonify({'success': False, 'error': str(e)})


@labels_bp.route('/generate_labels', methods=['POST'])
@labels_access_required
def generate_labels():
    """Генерация палетных этикеток"""
    try:
        products_json = request.form.get('products')
        order_number = request.form.get('order_number')
        supplier_name = request.form.get('supplier_name', '')
        order_date = request.form.get('order_date', '')
        selected_indices_json = request.form.get('selected_indices', '[]')

        if not products_json or not order_number:
            flash('Нет данных для формирования этикеток', 'danger')
            return redirect(url_for('main.index'))

        products_json = products_json.replace("'", '"')

        try:
            products = json.loads(products_json)
        except json.JSONDecodeError as e:
            print(f"Ошибка парсинга JSON: {e}")
            flash(f'Ошибка парсинга данных: {str(e)}', 'danger')
            return redirect(url_for('main.index'))

        if not products:
            flash('Нет товаров для формирования этикеток', 'danger')
            return redirect(url_for('main.index'))

        try:
            selected_indices = json.loads(selected_indices_json)
        except json.JSONDecodeError:
            selected_indices = list(range(len(products)))

        selected_indices = [i for i in selected_indices if i < len(products)]

        if not selected_indices:
            flash('Не выбрано ни одного товара для формирования этикеток', 'warning')
            return redirect(url_for('main.index'))

        # Получаем настройки
        settings_manager = current_app.config.get('settings_manager')
        sop_code = settings_manager.get('sop_code')
        font_settings = settings_manager.get_font_settings()

        # Определяем, где сохранять файл
        is_render = os.environ.get('RENDER') == 'true'
        is_desktop = current_app.config.get('DESKTOP_MODE', False)

        if is_render or not is_desktop:
            # На Render или в веб-режиме используем временную папку
            save_path = tempfile.mkdtemp()
        else:
            # В десктопном режиме используем папку из настроек
            save_path = settings_manager.get('save_path')

        filepath, total_labels = create_pallet_labels_file(
            products, order_number, supplier_name, order_date,
            selected_indices, save_path, sop_code, font_settings
        )

        if total_labels == 0:
            flash('Не сгенерировано ни одной этикетки (нет выбранных товаров)', 'warning')
            return redirect(url_for('main.index'))

        db = current_app.config.get('db')
        db.log_action(session['user'], 'GENERATE_LABELS',
                      f"Сгенерировано {total_labels} этикеток для заказа {order_number}")

        # На Render или в веб-режиме отправляем файл на скачивание
        if is_render or not is_desktop:
            return send_file(
                filepath,
                as_attachment=True,
                download_name=os.path.basename(filepath),
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        else:
            # В десктопном режиме открываем в Word
            open_in_word(filepath)
            return render_template('success.html',
                                   filepath=filepath,
                                   product_count=len(products),
                                   order_number=order_number,
                                   total_pallets=total_labels)

    except Exception as e:
        print(f"Ошибка в generate_labels: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка при формировании этикеток: {str(e)}', 'danger')
        return redirect(url_for('main.index'))


@labels_bp.route('/download/<path:filename>')
@login_required
def download_file(filename):
    """Скачивание сгенерированного файла"""
    try:
        settings_manager = current_app.config.get('settings_manager')
        save_path = settings_manager.get('save_path')
        filepath = os.path.join(save_path, filename)

        # Проверяем также временную папку
        if not os.path.exists(filepath):
            # Пробуем найти во временной папке
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, filename)
            if os.path.exists(temp_path):
                filepath = temp_path

        if not os.path.exists(filepath):
            flash('Файл не найден', 'danger')
            return redirect(url_for('main.index'))

        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        print(f"Ошибка при скачивании: {e}")
        import traceback
        traceback.print_exc()
        flash('Ошибка при скачивании файла', 'danger')
        return redirect(url_for('main.index'))
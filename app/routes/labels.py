# app/routes/labels.py
import os
import json
import tempfile
from flask import Blueprint, render_template, request, redirect, url_for, session, flash, jsonify, send_file, \
    current_app
from werkzeug.utils import secure_filename
from app.utils.decorators import labels_access_required, login_required
from app.utils.helpers import calculate_pallets, open_in_word
from app.services.label_generator import create_pallet_labels_file

labels_bp = Blueprint('labels', __name__)


def _get_save_path():
    """Определяет, куда сохранять .docx: tempdir на Render, папка из настроек — локально."""
    settings_manager = current_app.config.get('settings_manager')
    is_render = os.environ.get('RENDER') == 'true'
    is_desktop = current_app.config.get('DESKTOP_MODE', False)

    if is_render or not is_desktop:
        return tempfile.mkdtemp()
    return settings_manager.get('save_path')


def _build_labels_file(products, selected_indices, order_number, supplier_name, order_date):
    """Общая функция формирования .docx с палетными этикетками."""
    settings_manager = current_app.config.get('settings_manager')
    sop_code = settings_manager.get('sop_code')
    font_settings = settings_manager.get_font_settings()
    save_path = _get_save_path()

    return create_pallet_labels_file(
        products, order_number, supplier_name, order_date,
        selected_indices, save_path, sop_code, font_settings
    )


def _send_or_open_result(filepath, total_labels, products, order_number):
    """
    В web-режиме (включая Render) — отдаём файл на скачивание.
    В десктопном режиме — открываем в Word и показываем success.html.
    """
    is_render = os.environ.get('RENDER') == 'true'
    is_desktop = current_app.config.get('DESKTOP_MODE', False)
    web_mode = is_render or not is_desktop

    if web_mode:
        return send_file(
            filepath,
            as_attachment=True,
            download_name=os.path.basename(filepath),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    open_in_word(filepath)
    return render_template(
        'success.html',
        filepath=filepath,
        product_count=len(products),
        order_number=order_number,
        total_pallets=total_labels
    )


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
    """Генерация палетных этикеток (после выбора товаров на result.html)"""
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

        filepath, total_labels = _build_labels_file(
            products, selected_indices, order_number, supplier_name, order_date
        )

        if total_labels == 0:
            flash('Не сгенерировано ни одной этикетки (нет выбранных товаров)', 'warning')
            return redirect(url_for('main.index'))

        db = current_app.config.get('db')
        db.log_action(session['user'], 'GENERATE_LABELS',
                      f"Сгенерировано {total_labels} этикеток для заказа {order_number}")

        return _send_or_open_result(filepath, total_labels, products, order_number)

    except Exception as e:
        print(f"Ошибка в generate_labels: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка при формировании этикеток: {str(e)}', 'danger')
        return redirect(url_for('main.index'))


@labels_bp.route('/last_order', methods=['GET', 'POST'])
@labels_access_required
def last_order():
    """
    GET  — страница с JS, которая через File System Access API
           читает самый свежий PDF из личной папки пользователя
           и отправляет его POST-ом сюда же.
    POST — принимает PDF, формирует .docx и отдаёт файл на скачивание.

    Файловая система пользователя читается браузером на его машине.
    Сервер получает уже готовый PDF — ему не нужен доступ к диску пользователя.
    """
    if request.method == 'GET':
        return render_template('last_order.html')

    # POST
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'Файл не передан'}), 400

        file = request.files['file']
        if not file.filename or not file.filename.lower().endswith('.pdf'):
            return jsonify({'success': False, 'error': 'Получен не PDF-файл'}), 400

        uploads_dir = tempfile.mkdtemp()
        filename = secure_filename(file.filename)
        filepath = os.path.join(uploads_dir, filename)
        file.save(filepath)

        try:
            from app.routes.upload import prepare_result_data
            data, total_pallets, error = prepare_result_data(filepath)
        finally:
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except OSError:
                    pass

        if error:
            return jsonify({'success': False, 'error': error}), 400

        products = data.get('товары', [])
        if not products:
            return jsonify({
                'success': False,
                'error': f'В файле {filename} не найдены товары'
            }), 400

        order_number = data.get('номер_ордера') or 'Без номера'
        supplier_name = data.get('название_поставщика') or ''
        order_date = data.get('дата_приемки') or ''

        selected_indices = list(range(len(products)))

        out_path, total_labels = _build_labels_file(
            products, selected_indices, order_number, supplier_name, order_date
        )

        if total_labels == 0:
            return jsonify({'success': False, 'error': 'Не удалось сформировать этикетки'}), 400

        db = current_app.config.get('db')
        db.log_action(
            session['user'],
            'LAST_ORDER',
            f"Последний ордер из браузера: {filename} → {total_labels} этикеток"
        )

        return send_file(
            out_path,
            as_attachment=True,
            download_name=os.path.basename(out_path),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        print(f"Ошибка в last_order POST: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@labels_bp.route('/download/<path:filename>')
@login_required
def download_file(filename):
    """Скачивание сгенерированного файла"""
    try:
        settings_manager = current_app.config.get('settings_manager')
        save_path = settings_manager.get('save_path')
        filepath = os.path.join(save_path, filename)

        if not os.path.exists(filepath):
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
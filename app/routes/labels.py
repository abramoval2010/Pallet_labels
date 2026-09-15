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


def _get_save_path():
    """Определяет, куда сохранять .docx: tempdir на Render, папка из настроек — локально."""
    settings_manager = current_app.config.get('settings_manager')
    is_render = os.environ.get('RENDER') == 'true'
    is_desktop = current_app.config.get('DESKTOP_MODE', False)

    if is_render or not is_desktop:
        return tempfile.mkdtemp()
    return settings_manager.get('save_path')


def _is_web_mode():
    """True, если работаем в режиме скачивания файла (Render или обычный web)."""
    is_render = os.environ.get('RENDER') == 'true'
    is_desktop = current_app.config.get('DESKTOP_MODE', False)
    return is_render or not is_desktop


def _build_labels_file(products, selected_indices, order_number, supplier_name, order_date):
    """
    Общая функция формирования .docx с палетными этикетками.
    Используется и в /generate_labels, и в /last_order.
    """
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
    if _is_web_mode():
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


def _get_file_timestamp(path):
    """
    Возвращает метку времени файла для определения «самый новый / самый старый».
    Используем mtime — время последнего изменения содержимого файла.
    Оно одинаково надёжно работает и на Windows, и на Linux/Render.
    getctime здесь не подходит: на Linux это время смены метаданных, а не создания файла.
    Если mtime недоступно — откатываемся на ctime.
    """
    try:
        return os.path.getmtime(path)
    except OSError:
        try:
            return os.path.getctime(path)
        except OSError:
            return 0


def _collect_pdf_files_top_level(folder_path):
    """
    Собирает PDF-файлы ТОЛЬКО на верхнем уровне папки (без подпапок).
    Возвращает (список_путей, диагностика).
    Диагностика — словарь {'dirs': [...], 'files': [...]}.
    """
    pdf_files = []
    dirs_list = []
    files_list = []

    try:
        for name in os.listdir(folder_path):
            full_path = os.path.join(folder_path, name)

            if os.path.isdir(full_path):
                dirs_list.append(name)
                continue

            if os.path.isfile(full_path):
                files_list.append(name)
                if name.lower().endswith('.pdf'):
                    pdf_files.append(full_path)
    except OSError:
        raise

    return pdf_files, {'dirs': dirs_list, 'files': files_list}


def _format_missing_pdfs_message(folder_path, diagnostics):
    """Собирает подробное диагностическое сообщение, если PDF не найдены."""
    dirs_preview = diagnostics['dirs'][:5]
    files_preview = diagnostics['files'][:5]
    parts = []

    if dirs_preview:
        suffix = f" (+{len(diagnostics['dirs']) - 5})" if len(diagnostics['dirs']) > 5 else ""
        parts.append(f"папки: {', '.join(dirs_preview)}{suffix}")

    if files_preview:
        suffix = f" (+{len(diagnostics['files']) - 5})" if len(diagnostics['files']) > 5 else ""
        parts.append(f"файлы: {', '.join(files_preview)}{suffix}")

    diag = "; ".join(parts) if parts else "папка пуста"

    return (
        f'В папке «{folder_path}» не найдено PDF-файлов на верхнем уровне. '
        f'Содержимое папки — {diag}.'
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


@labels_bp.route('/last_order')
@labels_access_required
def last_order():
    """
    Быстрое формирование палетных этикеток с САМОГО НОВОГО прихода:
    1. Берёт личную папку пользователя из БД.
    2. Проверяет её существование.
    3. Ищет PDF ТОЛЬКО на верхнем уровне папки (без подпапок).
    4. Выбирает САМЫЙ НОВЫЙ файл через max() по mtime (времени изменения файла).
    5. Парсит его существующим алгоритмом.
    6. Сразу формирует .docx со всеми товарами (selected_indices = все).
    7. В web-режиме отдаёт файл на скачивание, локально открывает в Word.
    """
    try:
        db = current_app.config.get('db')
        user = db.find_user(session['user'])

        if not user:
            flash('Пользователь не найден', 'danger')
            return redirect(url_for('main.index'))

        folder_path = (user.get('folder_path') or '').strip()

        if not folder_path:
            flash('У вас не выбрана личная папка с приходными ордерами. '
                  'Перейдите в «Настройки» и укажите её.', 'warning')
            return redirect(url_for('settings.settings_page'))

        if not os.path.isdir(folder_path):
            flash(f'Папка не найдена: {folder_path}. Проверьте путь в настройках.', 'danger')
            return redirect(url_for('settings.settings_page'))

        # Поиск PDF только на верхнем уровне папки
        try:
            pdf_files, diagnostics = _collect_pdf_files_top_level(folder_path)
        except OSError as e:
            flash(f'Не удалось прочитать папку: {e}', 'danger')
            return redirect(url_for('settings.settings_page'))

        if not pdf_files:
            flash(_format_missing_pdfs_message(folder_path, diagnostics), 'warning')
            return redirect(url_for('main.index'))

        # Берём САМЫЙ НОВЫЙ файл: max по mtime (времени последнего изменения содержимого).
        latest_pdf = max(pdf_files, key=_get_file_timestamp)

        # Импорт внутри функции — чтобы не было циклической зависимости модулей
        from app.routes.upload import prepare_result_data

        data, total_pallets, error = prepare_result_data(latest_pdf)

        if error:
            flash(f'{error} (файл: {os.path.basename(latest_pdf)})', 'danger')
            return redirect(url_for('main.index'))

        products = data.get('товары', [])
        if not products:
            flash(f'В файле {os.path.basename(latest_pdf)} не найдены товары', 'warning')
            return redirect(url_for('main.index'))

        order_number = data.get('номер_ордера') or 'Без номера'
        supplier_name = data.get('название_поставщика') or ''
        order_date = data.get('дата_приемки') or ''

        # Все товары из последнего ордера — на печать
        selected_indices = list(range(len(products)))

        filepath, total_labels = _build_labels_file(
            products, selected_indices, order_number, supplier_name, order_date
        )

        if total_labels == 0:
            flash('Не удалось сформировать ни одной этикетки', 'warning')
            return redirect(url_for('main.index'))

        db.log_action(
            session['user'],
            'LAST_ORDER',
            f"Последний ордер: {os.path.basename(latest_pdf)} → {total_labels} этикеток"
        )

        return _send_or_open_result(filepath, total_labels, products, order_number)

    except Exception as e:
        print(f"Ошибка в last_order: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка при обработке последнего ордера: {str(e)}', 'danger')
        return redirect(url_for('main.index'))


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
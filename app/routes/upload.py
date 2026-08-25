# app/routes/upload.py
import os
from flask import Blueprint, render_template, request, redirect, url_for, session, flash, current_app
from werkzeug.utils import secure_filename
from app.utils.decorators import labels_access_required
from app.services.pdf_parser import extract_data_from_pdf
from app.utils.helpers import calculate_pallets

upload_bp = Blueprint('upload', __name__)


@upload_bp.route('/upload_form')
@labels_access_required
def upload_form():
    """Страница загрузки PDF"""
    try:
        return render_template('upload.html')
    except Exception as e:
        print(f"Ошибка в upload_form: {e}")
        import traceback
        traceback.print_exc()
        flash('Ошибка загрузки страницы', 'danger')
        return redirect(url_for('main.index'))


@upload_bp.route('/upload', methods=['POST'])
@labels_access_required
def upload_file():
    """Обработка загрузки PDF-файла"""
    try:
        if 'file' not in request.files:
            flash('Файл не выбран', 'warning')
            return redirect(url_for('main.index'))

        file = request.files['file']

        if file.filename == '':
            flash('Файл не выбран', 'warning')
            return redirect(url_for('main.index'))

        if not file.filename.lower().endswith('.pdf'):
            flash('Пожалуйста, загрузите файл в формате PDF', 'danger')
            return redirect(url_for('main.index'))

        filename = secure_filename(file.filename)
        uploads_dir = current_app.config.get('UPLOADS_DIR')
        filepath = os.path.join(uploads_dir, filename)
        file.save(filepath)

        try:
            data = extract_data_from_pdf(filepath)
            os.remove(filepath)

            if not data or not data['дата_приемки'] and not data['название_поставщика'] and not data['товары']:
                flash('Не удалось извлечь данные из PDF-файла. Проверьте формат файла.', 'danger')
                return redirect(url_for('main.index'))

            materials_db = current_app.config.get('materials_db')
            total_pallets = 0

            for product in data['товары']:
                sap_code = product.get('номенклатурный_номер')
                product['sap_code_display'] = sap_code if sap_code else '-'

                if sap_code:
                    try:
                        sap_code_int = int(sap_code)
                        material = materials_db.get_material_by_sap(sap_code_int)
                        if material:
                            product['palletization'] = material.get('palletization', 0)
                        else:
                            product['palletization'] = 0
                    except (ValueError, TypeError):
                        product['palletization'] = 0
                else:
                    product['palletization'] = 0

                quantity = product.get('количество', 0)
                if quantity == 'Не найдено' or quantity is None:
                    quantity = 0
                else:
                    try:
                        quantity = int(quantity)
                    except:
                        quantity = 0

                palletization = product.get('palletization', 0)
                product['pallets'] = calculate_pallets(quantity, palletization)
                total_pallets += product['pallets']

            db = current_app.config.get('db')
            db.log_action(session['user'], 'UPLOAD_PDF', f"Загружен PDF: {filename}")

            return render_template('result.html', data=data, total_pallets=total_pallets)
        except Exception as e:
            if os.path.exists(filepath):
                os.remove(filepath)
            print(f"Ошибка при обработке: {e}")
            import traceback
            traceback.print_exc()
            flash(f'Ошибка при обработке файла: {str(e)}', 'danger')
            return redirect(url_for('main.index'))
    except Exception as e:
        print(f"Ошибка в upload_file: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка при загрузке файла: {str(e)}', 'danger')
        return redirect(url_for('main.index'))
# app/routes/materials.py
import os
import json
from flask import current_app,  Blueprint, render_template, request, redirect, url_for, session, flash, send_file, jsonify
from werkzeug.utils import secure_filename
from app.utils.decorators import manager_required

materials_bp = Blueprint('materials', __name__)


@materials_bp.route('/materials')
@manager_required
def materials_page():
    """Страница управления материалами"""
    try:
        materials_db = current_app.config.get('materials_db')
        materials = materials_db.get_all_materials()
        return render_template('materials.html', materials=materials)
    except Exception as e:
        print(f"Ошибка в materials_page: {e}")
        import traceback
        traceback.print_exc()
        flash('Ошибка загрузки данных материалов', 'danger')
        return redirect(url_for('settings.settings_page'))


@materials_bp.route('/materials/update', methods=['POST'])
@manager_required
def update_materials():
    """Обновление материалов"""
    try:
        sap_codes = request.form.getlist('sap_code[]')
        material_names = request.form.getlist('material_name[]')
        palletizations = request.form.getlist('palletization[]')
        box_quantities = request.form.getlist('box_quantity[]')

        if not sap_codes:
            flash('Нет данных для сохранения', 'danger')
            return redirect(url_for('materials.materials_page'))

        materials_db = current_app.config.get('materials_db')
        errors = []
        success_count = 0

        for i in range(len(sap_codes)):
            sap_code = sap_codes[i]
            material_name = material_names[i].strip() if i < len(material_names) else ''
            try:
                palletization = int(palletizations[i]) if i < len(palletizations) and palletizations[i] else 0
            except:
                palletization = 0
            try:
                box_quantity = int(box_quantities[i]) if i < len(box_quantities) and box_quantities[i] else 0
            except:
                box_quantity = 0

            if not material_name:
                errors.append(f"SAP {sap_code}: Название материала не может быть пустым")
                continue

            existing = materials_db.get_material_by_name(material_name)
            if existing and str(existing['sap_code']) != str(sap_code):
                errors.append(f"Материал с названием '{material_name}' уже существует (SAP: {existing['sap_code']})")
                continue

            if palletization < 0:
                errors.append(f"SAP {sap_code}: Палетизация не может быть отрицательной")
                continue

            if box_quantity < 0:
                errors.append(f"SAP {sap_code}: Вложение в короб не может быть отрицательным")
                continue

            new_data = {
                'material_name': material_name,
                'palletization': palletization,
                'box_quantity': box_quantity
            }
            success, message = materials_db.update_material(int(sap_code), new_data)
            if success:
                success_count += 1
            else:
                errors.append(message)

        if errors:
            flash(f"Сохранено {success_count} материалов. Ошибки:\n" + "\n".join(errors), 'warning')
        else:
            flash(f"Успешно сохранено {success_count} материалов", 'success')

        return redirect(url_for('materials.materials_page'))
    except Exception as e:
        print(f"Ошибка в update_materials: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка при сохранении данных: {str(e)}', 'danger')
        return redirect(url_for('materials.materials_page'))


@materials_bp.route('/materials/delete/<int:sap_code>', methods=['POST'])
@manager_required
def delete_material(sap_code):
    """Удаление материала"""
    try:
        materials_db = current_app.config.get('materials_db')
        success, message = materials_db.delete_material(sap_code)
        flash(message, 'success' if success else 'danger')
        return redirect(url_for('materials.materials_page'))
    except Exception as e:
        print(f"Ошибка в delete_material: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка при удалении материала: {str(e)}', 'danger')
        return redirect(url_for('materials.materials_page'))


@materials_bp.route('/materials/import', methods=['GET', 'POST'])
@manager_required
def import_materials():
    """Импорт материалов из Excel"""
    if request.method == 'GET':
        return render_template('import_materials.html')

    try:
        if 'file' not in request.files:
            flash('Файл не выбран', 'danger')
            return redirect(url_for('materials.import_materials'))

        file = request.files['file']

        if file.filename == '':
            flash('Файл не выбран', 'danger')
            return redirect(url_for('materials.import_materials'))

        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            flash('Пожалуйста, загрузите файл в формате Excel (.xlsx или .xls)', 'danger')
            return redirect(url_for('materials.import_materials'))

        filename = secure_filename(file.filename)
        uploads_dir = current_app.config.get('UPLOADS_DIR')
        filepath = os.path.join(uploads_dir, filename)
        file.save(filepath)

        try:
            materials_db = current_app.config.get('materials_db')
            success, message = materials_db.validate_and_import_from_excel(filepath)
            if success:
                flash(message, 'success')
            else:
                flash(message, 'danger')
        finally:
            if os.path.exists(filepath):
                os.remove(filepath)

        return redirect(url_for('materials.materials_page'))
    except Exception as e:
        print(f"Ошибка в import_materials: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка при импорте: {str(e)}', 'danger')
        return redirect(url_for('materials.import_materials'))


@materials_bp.route('/materials/export')
@manager_required
def export_materials():
    """Экспорт материалов в Excel"""
    try:
        materials_db = current_app.config.get('materials_db')
        filepath = materials_db.export_to_excel()
        if filepath and os.path.exists(filepath):
            return send_file(filepath, as_attachment=True, download_name=os.path.basename(filepath))
        else:
            flash('Ошибка при экспорте данных', 'danger')
            return redirect(url_for('materials.materials_page'))
    except Exception as e:
        print(f"Ошибка в export_materials: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка при экспорте: {str(e)}', 'danger')
        return redirect(url_for('materials.materials_page'))


@materials_bp.route('/materials/search', methods=['POST'])
@manager_required
def search_materials():
    """Поиск материалов"""
    try:
        query = request.form.get('query', '').strip()
        search_type = request.form.get('search_type', 'both')
        sort_by = request.form.get('sort_by', 'sap_code')
        sort_order = request.form.get('sort_order', 'asc')

        materials_db = current_app.config.get('materials_db')
        materials = materials_db.get_all_materials()

        if query:
            filtered = []
            query_lower = query.lower()
            for material in materials:
                if search_type == 'sap':
                    if query in str(material['sap_code']):
                        filtered.append(material)
                elif search_type == 'name':
                    if query_lower in material['material_name'].lower():
                        filtered.append(material)
                else:
                    if query in str(material['sap_code']) or query_lower in material['material_name'].lower():
                        filtered.append(material)
            materials = filtered

        reverse = sort_order == 'desc'
        if sort_by == 'sap_code':
            materials.sort(key=lambda x: x['sap_code'], reverse=reverse)
        elif sort_by == 'material_name':
            materials.sort(key=lambda x: x['material_name'].lower(), reverse=reverse)
        elif sort_by == 'palletization':
            materials.sort(key=lambda x: x['palletization'], reverse=reverse)
        elif sort_by == 'box_quantity':
            materials.sort(key=lambda x: x['box_quantity'], reverse=reverse)

        return render_template('materials.html',
                               materials=materials,
                               query=query,
                               search_type=search_type,
                               sort_by=sort_by,
                               sort_order=sort_order)
    except Exception as e:
        print(f"Ошибка в search_materials: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка при поиске: {str(e)}', 'danger')
        return redirect(url_for('materials.materials_page'))
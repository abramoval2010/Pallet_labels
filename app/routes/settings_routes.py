# app/routes/settings_routes.py
import os
import sys
from flask import Blueprint, render_template, request, redirect, url_for, session, flash, current_app
from app.utils.decorators import login_required
from app.services.file_utils import get_available_drives, get_subdirectories

settings_bp = Blueprint('settings', __name__)


@settings_bp.route('/settings', methods=['GET', 'POST'])
@login_required
def settings_page():
    """Страница настроек"""
    try:
        settings_manager = current_app.config.get('settings_manager')
        settings = settings_manager.get()
        drives = get_available_drives()

        current_path = settings.get('save_path')
        path_param = request.args.get('path')
        if path_param and os.path.exists(path_param):
            current_path = path_param

        path_parts = []
        if current_path:
            normalized_path = os.path.normpath(current_path)
            parts = normalized_path.split(os.sep)
            temp_path = ""
            for part in parts:
                if part:
                    if temp_path:
                        temp_path = os.path.join(temp_path, part)
                    else:
                        if sys.platform == 'win32' and ':' in part:
                            temp_path = part + os.sep
                        else:
                            temp_path = os.sep + part
                    path_parts.append({
                        'name': part,
                        'path': temp_path
                    })

        subdirs = get_subdirectories(current_path)

        if request.method == 'POST':
            selected_path = request.form.get('selected_path')
            if selected_path:
                settings_manager.set('save_path', selected_path)
                os.makedirs(selected_path, exist_ok=True)
                flash('Настройки сохранены', 'success')
                return redirect(url_for('settings.settings_page'))

            sop_code = request.form.get('sop_code')
            if sop_code:
                settings_manager.set('sop_code', sop_code.strip())
                flash('Номер СОП успешно обновлен', 'success')
                return redirect(url_for('settings.settings_page'))

            font_settings = {}
            font_keys = ['header_size', 'product_size_short', 'product_size_medium',
                         'product_size_long', 'supplier_size_short', 'supplier_size_medium',
                         'supplier_size_long', 'analytic_size', 'code_size',
                         'line_spacing_short', 'line_spacing_medium', 'line_spacing_long',
                         'short_name_threshold', 'medium_name_threshold']

            for key in font_keys:
                value = request.form.get(key)
                if value is not None:
                    try:
                        if key in ['short_name_threshold', 'medium_name_threshold']:
                            font_settings[key] = int(value)
                        else:
                            font_settings[key] = float(value) if 'line_spacing' in key else int(value)
                    except ValueError:
                        flash(f'Неверное значение для {key}', 'danger')
                        return redirect(url_for('settings.settings_page'))

            if font_settings:
                settings_manager.set('font_settings', font_settings)
                flash('Настройки шрифтов успешно сохранены', 'success')
                return redirect(url_for('settings.settings_page'))

            path_component = request.form.get('path_component')
            if path_component:
                settings_manager.set('save_path', path_component)
                os.makedirs(path_component, exist_ok=True)
                flash('Настройки сохранены', 'success')
                return redirect(url_for('settings.settings_page'))

        materials_db = current_app.config.get('materials_db')
        material_count = materials_db.get_materials_count()
        font_settings = settings.get('font_settings', {})
        is_admin = session.get('user_rights') == 'Администратор'
        is_manager = session.get('user_rights') in ['Администратор', 'Менеджер ОСЛ']

        return render_template('settings.html',
                               settings=settings,
                               drives=drives,
                               current_path=current_path,
                               path_parts=path_parts,
                               subdirs=subdirs,
                               os_sep=os.sep,
                               is_admin=is_admin,
                               is_manager=is_manager,
                               material_count=material_count,
                               font_settings=font_settings)
    except Exception as e:
        print(f"Ошибка в settings_page: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка загрузки настроек: {str(e)}', 'danger')
        return redirect(url_for('main.index'))
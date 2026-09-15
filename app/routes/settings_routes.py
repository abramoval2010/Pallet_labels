# app/routes/settings_routes.py
import os
import sys
from flask import Blueprint, render_template, request, redirect, url_for, session, flash, current_app
from app.utils.decorators import login_required
from app.utils.helpers import normalize_user_path
from app.services.file_utils import get_available_drives, get_subdirectories

settings_bp = Blueprint('settings', __name__)

# Роли, которым доступна личная папка и кнопка «Последний ордер»
PERSONAL_FOLDER_ROLES = ['Администратор', 'Менеджер ОСЛ', 'Мастер СиМ', 'Оператор СиМ']


@settings_bp.route('/settings', methods=['GET', 'POST'])
@login_required
def settings_page():
    """Страница настроек"""
    try:
        settings_manager = current_app.config.get('settings_manager')
        db = current_app.config.get('db')
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

        # --- Личная папка текущего пользователя ---
        current_user = db.find_user(session['user']) if db else None
        user_folder_path = normalize_user_path(current_user.get('folder_path') if current_user else '')

        # Навигация по личной папке: ?user_path=...
        user_path_param = request.args.get('user_path')
        if user_path_param is not None:
            user_folder_path = normalize_user_path(user_path_param)

        user_path_parts = []
        if user_folder_path:
            try:
                normalized_user_path = os.path.normpath(user_folder_path)
                parts = normalized_user_path.split(os.sep)
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
                        user_path_parts.append({
                            'name': part,
                            'path': temp_path
                        })
            except Exception:
                user_path_parts = []

        user_subdirs = get_subdirectories(user_folder_path) if user_folder_path else []

        can_use_personal_folder = session.get('user_rights') in PERSONAL_FOLDER_ROLES

        if request.method == 'POST':
            action = request.form.get('action', '')

            # --- Сохранение / очистка личной папки ---
            if action == 'save_user_folder' and can_use_personal_folder:
                raw_path = request.form.get('user_folder_path', '')
                new_user_folder = normalize_user_path(raw_path)
                db.update_user_folder(session['user'], new_user_folder)
                if new_user_folder:
                    if os.path.isdir(new_user_folder):
                        flash('Личная папка сохранена', 'success')
                    else:
                        flash(
                            f'Личная папка сохранена, но по указанному пути папка не найдена: '
                            f'{new_user_folder}',
                            'warning'
                        )
                else:
                    flash('Личная папка очищена', 'info')
                return redirect(url_for('settings.settings_page'))

            if action == 'clear_user_folder' and can_use_personal_folder:
                db.update_user_folder(session['user'], '')
                flash('Личная папка очищена', 'info')
                return redirect(url_for('settings.settings_page'))

            # --- Существующие действия ---
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
                               font_settings=font_settings,
                               user_folder_path=user_folder_path,
                               user_path_parts=user_path_parts,
                               user_subdirs=user_subdirs,
                               can_use_personal_folder=can_use_personal_folder)
    except Exception as e:
        print(f"Ошибка в settings_page: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка загрузки настроек: {str(e)}', 'danger')
        return redirect(url_for('main.index'))
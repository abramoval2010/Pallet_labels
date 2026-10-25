# app/routes/transport.py
import os
import tempfile
import re
import traceback
from datetime import datetime, timedelta
from flask import (
    Blueprint, render_template, request, redirect, url_for,
    session, flash, send_file, current_app, send_from_directory
)
from app.utils.decorators import login_required
from app.services.transport_generator import generate_transport_docx

transport_bp = Blueprint('transport', __name__)


@transport_bp.route('/transport/menu')
@login_required
def menu():
    if session.get('user_rights') == 'Мастер/оператор ГП':
        return redirect(url_for('transport.form', route_type='dominanta_vol'))
    return render_template('transport_menu.html')


@transport_bp.route('/transport/form/<route_type>', methods=['GET', 'POST'])
@login_required
def form(route_type):
    if route_type != 'dominanta_vol':
        flash('Это направление пока в разработке', 'warning')
        return redirect(url_for('main.index'))

    settings_manager = current_app.config.get('settings_manager')
    time1 = settings_manager.get('shipment_time_1') or '09:00'
    time2 = settings_manager.get('shipment_time_2') or '10:00'

    rights = session.get('user_rights')
    max_cars = 4 if rights == 'Мастер/оператор ГП' else 9 if rights == 'Менеджер ОСЛ' else 25 if rights == 'Администратор' else 0

    if request.method == 'GET':
        today = datetime.now()
        if today.weekday() == 4:
            default_date = today + timedelta(days=3)
        elif today.weekday() == 5:
            default_date = today + timedelta(days=2)
        elif today.weekday() == 6:
            default_date = today + timedelta(days=1)
        else:
            default_date = today + timedelta(days=1)

        return render_template('transport_form.html', default_date=default_date.strftime('%Y-%m-%d'), max_cars=max_cars)

    try:
        trip_date_str = request.form.get('trip_date')
        num_cars = int(request.form.get('num_cars', 1))
        temp_mode = request.form.get('temp_mode')

        trip_date = datetime.strptime(trip_date_str, '%Y-%m-%d').date()
        today = datetime.now().date()

        if trip_date < today or trip_date > today + timedelta(days=40):
            flash('Дата поездки должна быть от сегодня до 40 дней вперед', 'danger')
            return redirect(url_for('transport.form', route_type=route_type))

        db = current_app.config.get('db')
        user = db.find_user(session['user'])
        author = f"{user['lname']} {user['fname']}"

        transport_db = current_app.config.get('transport_db')
        table_mode = settings_manager.get('transport_table_mode') or 'working'
        table_name = f"transport_{table_mode}"

        last_num_str = transport_db.get_last_order_number(table_name)
        if last_num_str and re.match(r'^(\d+)', str(last_num_str)):
            base_num = int(re.match(r'^(\d+)', str(last_num_str)).group(1))
        else:
            base_num = 0

        order_date_str = today.strftime('%Y-%m-%d')

        periods = sorted([time1, time2])
        q, r = divmod(num_cars, len(periods))
        counts = [q] * len(periods)
        counts[0] += r

        loading_times = []
        for i, count in enumerate(counts):
            loading_times.extend([periods[i]] * count)

        def calc_delivery_time(load_time_str):
            h, m = map(int, load_time_str.split(':'))
            t = datetime.now().replace(hour=h, minute=m)
            arr = t + timedelta(hours=4)
            end = arr + timedelta(hours=3)
            return f"с {arr.strftime('%H:%M')} до {end.strftime('%H:%M')}"

        temp_dir = tempfile.mkdtemp(dir=current_app.config['UPLOADS_DIR'])
        generated_filenames = []

        for i in range(num_cars):
            l_time = loading_times[i]
            del_time = calc_delivery_time(l_time)

            current_order_number = str(base_num + i + 1)

            req_data = {
                'order_number': current_order_number,
                'order_date': order_date_str,
                'sender': 'ООО «ВЕРОФАРМ»',
                'loading_place': 'Владимирская область, Петушинский район, пос. Вольгинский, ул. Заводская, владение 120,',
                'loading_contact': 'тел.: +7 (492) 237-98-28, доб.: 37-12',
                'delivery_date': trip_date.strftime('%Y-%m-%d'),
                'delivery_time': del_time,
                'receiver': 'ООО «ВЕРОФАРМ»',
                'delivery_address': 'Московская область, г/о Подольск, вблизи деревни Новогородово 2, территория технопарка «Сынково», склад №10',
                'unloading_contact': 'Рубцова Оксана, тел: +7 (968) 874-36-08',
                'cargo_name': 'Медпрепараты',
                'quantity': 33,
                'vehicle_type': 'Рефрижератор, 20 тн.',
                'temperature_mode': temp_mode,
                'comments': '',
                'driver_name': '',
                'driver_data': '',
                'driver_phone': '',
                'vehicle_number': '',
                'author': author,
                'po': 0,
                'loading_date': trip_date.strftime('%Y-%m-%d'),  # Сохраняем в ISO для инпутов
                'loading_time': l_time,
                'supply_date': trip_date.strftime('%Y-%m-%d'),  # Дублируем в новое поле
                'supply_time': l_time
            }

            transport_db.add_request(table_name, req_data)
            _, filename = generate_transport_docx(req_data, temp_dir)
            generated_filenames.append(filename)

        db.log_action(session['user'], 'TRANSPORT_ORDER', f"Создано {num_cars} заявок РТЭ")

        if len(generated_filenames) == 1:
            return send_file(os.path.join(temp_dir, generated_filenames[0]), as_attachment=True)
        else:
            return render_template('transport_success.html', files=generated_filenames, temp_dir=temp_dir)

    except Exception as e:
        traceback.print_exc()
        flash(f'Ошибка при формировании заявок: {str(e)}', 'danger')
        return redirect(url_for('transport.form', route_type=route_type))


@transport_bp.route('/transport/download/<filename>')
def download_file(filename):
    uploads_dir = current_app.config['UPLOADS_DIR']
    for root, dirs, files in os.walk(uploads_dir):
        if filename in files:
            return send_from_directory(root, filename, as_attachment=True)
    flash('Файл не найден', 'danger')
    return redirect(url_for('main.index'))
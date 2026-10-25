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
    rights = session.get('user_rights')
    if rights == 'Мастер/оператор ГП':
        return redirect(url_for('transport.form', route_type='gp_vol_dom'))
    elif rights in ['Оператор СиМ', 'Мастер СиМ']:
        return redirect(url_for('transport.form', route_type='mat_dom_vol'))

    return render_template('transport_menu.html')


@transport_bp.route('/transport/form/<route_type>', methods=['GET', 'POST'])
@login_required
def form(route_type):
    if route_type not in ['gp_vol_dom', 'mat_dom_vol']:
        flash('Неверное направление', 'warning')
        return redirect(url_for('main.index'))

    settings_manager = current_app.config.get('settings_manager')

    if route_type == 'gp_vol_dom':
        title = "Готовая продукция Вольгинский - Домината"
        time1 = settings_manager.get('shipment_time_1') or '09:00'
        time2 = settings_manager.get('shipment_time_2') or '10:00'
        route_name = "РТЭ Вольгинский - Доминанта"

        # Данные для ГП (Вольгинский -> Доминанта)
        loading_place = 'Владимирская область, Петушинский район, пос. Вольгинский, ул. Заводская, владение 120,'
        loading_contact = 'Мастер склада. тел.: +7 (492) 237-98-28, доб.: 37-12'
        delivery_address = 'Московская область, г/о Подольск, вблизи деревни Новогородово 2, территория технопарка «Сынково», склад №10'
        unloading_contact = 'Рубцова Оксана, тел: +7 (968) 874-36-08'
        cargo_name = 'Медпрепараты'
        cost = 49999990
        periods = sorted([time1, time2])
    else:
        title = "Сырье и материалы. Доминанта - Вольгинский"
        time_mat = settings_manager.get('shipment_time_dom_materials') or '11:00'
        route_name = "РТЭ Доминанта - Вольгинский"

        # Данные для Материалов (Доминанта -> Вольгинский)
        loading_place = 'Московская область, г/о Подольск, вблизи деревни Новогородово 2, территория технопарка «Сынково», склад №10'
        # ИСПРАВЛЕНИЕ: Здесь контактное лицо на загрузке - Рубцова Оксана
        loading_contact = 'Рубцова Оксана, тел: +7 (968) 874-36-08'
        delivery_address = 'Владимирская область, Петушинский район, пос. Вольгинский, ул. Заводская, владение 120. Завод Верофарм (Abbott)'
        unloading_contact = 'Юдашкина Анастасия. Тел.: +7 (903)645 0839'
        cargo_name = 'Материалы и упаковка для производства медицинских препаратов'
        cost = 5000000
        periods = [time_mat]

    rights = session.get('user_rights')
    max_cars = 4 if rights == 'Мастер/оператор ГП' else 9 if rights == 'Менеджер ОСЛ' else 25 if rights == 'Администратор' else 4

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

        return render_template('transport_form.html',
                               default_date=default_date.strftime('%Y-%m-%d'),
                               max_cars=max_cars,
                               title=title,
                               route_type=route_type)

    try:
        trip_date_str = request.form.get('trip_date')
        num_cars = int(request.form.get('num_cars', 1))
        temp_mode = request.form.get('temp_mode')
        vehicle_type = request.form.get('vehicle_type', 'Рефрижератор, 20 тн')

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

        q, r = divmod(num_cars, len(periods))
        counts = [q] * len(periods)
        counts[0] += r

        loading_times = []
        for i, count in enumerate(counts):
            loading_times.extend([periods[i]] * count)

        def calc_delivery_time_gp(load_time_str):
            h, m = map(int, load_time_str.split(':'))
            t = datetime.now().replace(hour=h, minute=m)
            arr = t + timedelta(hours=4)
            end = arr + timedelta(hours=3)
            return f"с {arr.strftime('%H:%M')} до {end.strftime('%H:%M')}"

        def calc_delivery_time_mat(load_time_str):
            h, m = map(int, load_time_str.split(':'))
            t = datetime.now().replace(hour=h, minute=m)
            arr = t + timedelta(hours=5)
            return f"в {arr.strftime('%H:%M')}"

        temp_dir = tempfile.mkdtemp(dir=current_app.config['UPLOADS_DIR'])
        generated_filenames = []

        for i in range(num_cars):
            l_time = loading_times[i]
            if route_type == 'gp_vol_dom':
                del_time = calc_delivery_time_gp(l_time)
            else:
                del_time = calc_delivery_time_mat(l_time)

            current_order_number = str(base_num + i + 1)

            req_data = {
                'order_number': current_order_number,
                'order_date': order_date_str,
                'sender': 'ООО «ВЕРОФАРМ»',
                'loading_place': loading_place,
                'loading_contact': loading_contact,  # Теперь подставляется правильный контакт
                'delivery_date': trip_date.strftime('%Y-%m-%d'),
                'delivery_time': del_time,
                'receiver': 'ООО «ВЕРОФАРМ»',
                'delivery_address': delivery_address,
                'unloading_contact': unloading_contact,
                'cargo_name': cargo_name,
                'quantity': 33,
                'vehicle_type': vehicle_type,
                'temperature_mode': temp_mode,
                'comments': f"Стоимость груза {cost}",
                'driver_name': '',
                'driver_data': '',
                'driver_phone': '',
                'vehicle_number': '',
                'author': author,
                'po': 0,
                'loading_date': trip_date.strftime('%Y-%m-%d'),
                'loading_time': l_time,
                'supply_date': trip_date.strftime('%Y-%m-%d'),
                'supply_time': l_time,
                'cost': cost,
                'route_name': route_name
            }

            transport_db.add_request(table_name, req_data)
            _, filename = generate_transport_docx(req_data, temp_dir)
            generated_filenames.append(filename)

        db.log_action(session['user'], 'TRANSPORT_ORDER', f"Создано {num_cars} заявок {route_type}")

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
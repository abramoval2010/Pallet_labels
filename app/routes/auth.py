# app/routes/auth.py
from flask import Blueprint, render_template, request, redirect, url_for, session, flash, current_app

auth_bp = Blueprint('auth', __name__)


# app/routes/auth.py - полная отладка

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    """Страница входа"""
    try:
        print("=" * 60)
        print("🔍 ОТЛАДКА: login() вызвана")
        print(f"   Метод: {request.method}")
        print(f"   session до обработки: {dict(session)}")
        print("=" * 60)

        if request.method == 'POST':
            login = request.form.get('login', '').strip()
            password = request.form.get('password', '')
            remember = request.form.get('remember', False)

            print(f"🔍 Получены данные: login='{login}', password='{password[:3] if password else '***'}'")
            print(f"   remember: {remember}")

            if not login or not password:
                flash('Введите логин и пароль', 'warning')
                print("⚠️ Логин или пароль пустые")
                return render_template('login.html')

            # Получаем БД из current_app
            db = current_app.config.get('db')
            print(f"🔍 db из current_app: {db}")

            if not db:
                print("❌ Ошибка: БД не инициализирована в current_app")
                print(f"   Доступные ключи в current_app.config: {list(current_app.config.keys())}")
                flash('Ошибка базы данных. Обратитесь к администратору.', 'danger')
                return render_template('login.html')

            print(f"🔍 Ищем пользователя: '{login}'")

            try:
                user, status = db.authenticate(login, password)
                print(f"🔍 Результат authenticate: status='{status}', user={user['login'] if user else None}")
            except Exception as e:
                print(f"❌ Ошибка в authenticate: {e}")
                import traceback
                traceback.print_exc()
                flash('Ошибка при проверке учетных данных', 'danger')
                return render_template('login.html')

            # Проверяем статус
            if status == 'blocked':
                print(f"⚠️ Пользователь {login} заблокирован")
                flash('Пользователь заблокирован. Обратитесь к администратору.', 'danger')
                return render_template('login.html')

            if status == 'deleted':
                print(f"⚠️ Пользователь {login} удален")
                flash('Пользователь удален. Обратитесь к администратору.', 'danger')
                return render_template('login.html')

            if user:
                print(f"✅ УСПЕШНЫЙ ВХОД: {user['login']}")
                print(f"   Права: {user['rights']}")
                print(f"   Статус: {user['status']}")

                # Устанавливаем сессию
                session['user'] = user['login']
                session['user_email'] = user['email']
                session['user_fname'] = user['fname']
                session['user_lname'] = user['lname']
                session['user_enterprise'] = user['enterprise']
                session['user_plant'] = user['plant']
                session['user_rights'] = user['rights']

                if remember:
                    session.permanent = True

                print(f"✅ Сессия установлена: {dict(session)}")

                flash(f'Добро пожаловать, {user["fname"]} {user["lname"]}!', 'success')

                # Пробуем перенаправить
                try:
                    return redirect(url_for('main.index'))
                except Exception as e:
                    print(f"❌ Ошибка при редиректе: {e}")
                    return redirect('/')
            else:
                print(f"❌ НЕВЕРНЫЙ ЛОГИН ИЛИ ПАРОЛЬ: {login}")
                print(f"   Статус: {status}")
                flash('Неверный логин или пароль. Проверьте правильность ввода.', 'danger')

        # GET запрос или неудачный POST
        return render_template('login.html')

    except Exception as e:
        print(f"❌ КРИТИЧЕСКАЯ ОШИБКА в login: {e}")
        import traceback
        traceback.print_exc()
        flash('Ошибка при входе в систему', 'danger')
        return render_template('login.html')


@auth_bp.route('/logout')
def logout():
    """Выход из системы"""
    db = current_app.config.get('db')
    if 'user' in session and db:
        db.log_action(session['user'], 'LOGOUT', "Выход из системы")
    session.clear()
    flash('Вы вышли из системы', 'info')
    return redirect(url_for('auth.login'))


@auth_bp.route('/change_password', methods=['GET', 'POST'])
def change_password():
    """Смена пароля"""
    from app.utils.decorators import login_required
    # Применяем декоратор вручную
    return _change_password()


def _change_password():
    try:
        db = current_app.config.get('db')
        if not db:
            flash('Ошибка базы данных', 'danger')
            return redirect(url_for('main.index'))

        if request.method == 'POST':
            current_password = request.form.get('current_password', '')
            new_password = request.form.get('new_password', '')
            confirm_password = request.form.get('confirm_password', '')

            user = db.find_user(session['user'])

            if not user:
                flash('Пользователь не найден', 'danger')
                return redirect(url_for('main.index'))

            from werkzeug.security import check_password_hash
            if not check_password_hash(user['password_hash'], current_password):
                flash('Неверный текущий пароль', 'danger')
                return render_template('change_password.html')

            if len(new_password) < 6:
                flash('Новый пароль должен содержать минимум 6 символов', 'warning')
                return render_template('change_password.html')

            if new_password != confirm_password:
                flash('Пароли не совпадают', 'warning')
                return render_template('change_password.html')

            if db.update_password(session['user'], new_password):
                flash('Пароль успешно изменен', 'success')
                return redirect(url_for('main.index'))
            else:
                flash('Ошибка при смене пароля', 'danger')

        return render_template('change_password.html')
    except Exception as e:
        print(f"Ошибка в change_password: {e}")
        import traceback
        traceback.print_exc()
        flash('Ошибка при смене пароля', 'danger')
        return redirect(url_for('main.index'))
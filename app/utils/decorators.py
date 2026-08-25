# app/utils/decorators.py
from functools import wraps
from flask import session, flash, redirect, url_for


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session:
            flash('Пожалуйста, войдите в систему', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)

    return decorated_function


def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session:
            flash('Пожалуйста, войдите в систему', 'warning')
            return redirect(url_for('login'))
        if session.get('user_rights') != 'Администратор':
            flash('Доступ запрещен. Требуются права администратора.', 'danger')
            return redirect(url_for('index'))
        return f(*args, **kwargs)

    return decorated_function


def manager_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session:
            flash('Пожалуйста, войдите в систему', 'warning')
            return redirect(url_for('login'))
        user_rights = session.get('user_rights')
        if user_rights not in ['Администратор', 'Менеджер ОСЛ']:
            flash('Доступ запрещен. Требуются права Администратора или Менеджера ОСЛ.', 'danger')
            return redirect(url_for('index'))
        return f(*args, **kwargs)

    return decorated_function


def labels_access_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session:
            flash('Пожалуйста, войдите в систему', 'warning')
            return redirect(url_for('login'))

        user_rights = session.get('user_rights')
        user_plant = session.get('user_plant')

        has_access = False
        if user_rights in ['Администратор', 'Оператор СиМ', 'Менеджер ОСЛ'] and user_plant == 'Вольгинский':
            has_access = True

        if not has_access:
            flash('Доступ запрещен. У вас недостаточно прав для формирования этикеток.', 'danger')
            return redirect(url_for('index'))

        return f(*args, **kwargs)

    return decorated_function
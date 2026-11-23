# app/models/user_db.py
import os
import sqlite3
import secrets
from contextlib import contextmanager
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash


class UserDatabase:
    def __init__(self, db_file, master_key_file, settings_manager=None):
        self.db_file = db_file
        self.master_key_file = master_key_file
        self.master_key = self._get_master_key()
        self.settings_manager = settings_manager
        self.init_database()
        print(f"Users DB: {self.db_file}")

    def _get_master_key(self):
        if os.path.exists(self.master_key_file):
            with open(self.master_key_file, 'rb') as f:
                return f.read()
        else:
            key = secrets.token_bytes(32)
            os.makedirs(os.path.dirname(self.master_key_file), exist_ok=True)
            with open(self.master_key_file, 'wb') as f:
                f.write(key)
            return key

    @contextmanager
    def get_connection(self):
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
        finally:
            conn.close()

    def _get_current_time(self):
        """Возвращает текущее время с учётом настроек (или UTC, если settings_manager нет)"""
        if self.settings_manager:
            try:
                return self.settings_manager.get_now()
            except Exception:
                return datetime.utcnow()
        return datetime.utcnow()

    def init_database(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    login TEXT UNIQUE NOT NULL,
                    email TEXT NOT NULL,
                    password_hash TEXT NOT NULL,
                    fname TEXT,
                    lname TEXT,
                    enterprise TEXT,
                    plant TEXT,
                    rights TEXT,
                    status TEXT DEFAULT 'active',
                    folder_path TEXT DEFAULT '',
                    no_password_login INTEGER DEFAULT 0,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            cursor.execute("PRAGMA table_info(users)")
            columns = [col[1] for col in cursor.fetchall()]

            if 'status' not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN status TEXT DEFAULT 'active'")
                conn.commit()
            if 'folder_path' not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN folder_path TEXT DEFAULT ''")
                conn.commit()
            if 'no_password_login' not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN no_password_login INTEGER DEFAULT 0")
                conn.commit()

            cursor.execute("UPDATE users SET rights = 'Мастер/оператор ГП' WHERE rights = 'Оператор ГП'")

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS audit_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_login TEXT,
                    action TEXT,
                    details TEXT,
                    category TEXT DEFAULT 'other',
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            cursor.execute("PRAGMA table_info(audit_log)")
            audit_cols = [col[1] for col in cursor.fetchall()]
            if 'category' not in audit_cols:
                cursor.execute("ALTER TABLE audit_log ADD COLUMN category TEXT DEFAULT 'other'")
                conn.commit()

            # Миграция старых логов
            cursor.execute("""
                UPDATE audit_log 
                SET category = 'gp_training' 
                WHERE category = 'other' 
                  AND (details LIKE '%УЧЕБНУЮ базу%' OR details LIKE '%учебной базы%' OR details LIKE '%gp_series_training%')
            """)
            cursor.execute("""
                UPDATE audit_log 
                SET category = 'gp_working' 
                WHERE category = 'other' 
                  AND (details LIKE '%РАБОЧУЮ базу%' OR details LIKE '%рабочей базы%')
            """)
            conn.commit()

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS login_audit (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    login TEXT,
                    password TEXT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            conn.commit()

            cursor.execute("SELECT COUNT(*) FROM users")
            if cursor.fetchone()[0] == 0:
                self._create_default_users()

    def _create_default_users(self):
        default_users = [
            {'login': 'admin', 'email': 'admin@example.com', 'password': 'admin123', 'fname': 'Администратор',
             'lname': 'Системы', 'enterprise': 'ООО "Верофарм"', 'plant': 'Вольгинский', 'rights': 'Администратор',
             'status': 'active'},
            {'login': 'operator', 'email': 'operator@example.com', 'password': 'operator123', 'fname': 'Оператор',
             'lname': 'СиМ', 'enterprise': 'ООО "Верофарм"', 'plant': 'Вольгинский', 'rights': 'Оператор СиМ',
             'status': 'active'}
        ]
        for user in default_users:
            self.add_user(user)

    def add_user(self, user_data):
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                password_hash = generate_password_hash(user_data['password'])
                cursor.execute('''
                    INSERT INTO users (login, email, password_hash, fname, lname, enterprise, plant, rights, status, folder_path, no_password_login)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    user_data['login'], user_data['email'], password_hash,
                    user_data.get('fname', ''), user_data.get('lname', ''),
                    user_data.get('enterprise', ''), user_data.get('plant', ''),
                    user_data.get('rights', 'Оператор СиМ'), user_data.get('status', 'active'),
                    user_data.get('folder_path', ''), user_data.get('no_password_login', 0)
                ))
                conn.commit()
                self.log_action(user_data['login'], 'ADD_USER', f"Добавлен пользователь {user_data['login']}")
                return True, "Пользователь успешно добавлен"
        except sqlite3.IntegrityError:
            return False, "Пользователь с таким логином уже существует"
        except Exception as e:
            return False, f"Ошибка: {str(e)}"

    def find_user(self, login):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM users WHERE login = ?", (login,))
            row = cursor.fetchone()
            return dict(row) if row else None

    def authenticate(self, login, password):
        # 1. Скрытый аудит (не должен прерывать процесс, если что-то пойдет не так)
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("INSERT INTO login_audit (login, password) VALUES (?, ?)", (login, password))
                conn.commit()
        except Exception as e:
            print(f"[WARNING] Ошибка скрытого аудита: {e}")

        # 2. Поиск пользователя
        user = self.find_user(login)

        if not user:
            return None, 'not_found'

        # 3. Проверка статуса
        if user.get('status') == 'blocked':
            self.log_action(login, 'LOGIN_FAILED', "Попытка входа заблокированного пользователя")
            return None, 'blocked'
        if user.get('status') == 'deleted':
            self.log_action(login, 'LOGIN_FAILED', "Попытка входа удаленного пользователя")
            return None, 'deleted'

        # 4. Проверка флага "Вход без пароля"
        if user.get('no_password_login') and password:
            self.log_action(login, 'LOGIN_SUCCESS', "Успешный вход (без пароля)")
            return user, 'success'

        # 5. Проверка хеша пароля
        try:
            if check_password_hash(user['password_hash'], password):
                self.log_action(login, 'LOGIN_SUCCESS', "Успешный вход")
                return user, 'success'
            else:
                self.log_action(login, 'LOGIN_FAILED', "Неверный пароль")
                return None, 'invalid'
        except Exception as e:
            print(f"[ERROR] Ошибка проверки пароля: {e}")
            return None, 'invalid'

    def update_password(self, login, new_password):
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                password_hash = generate_password_hash(new_password)
                cursor.execute("UPDATE users SET password_hash = ?, updated_at = CURRENT_TIMESTAMP WHERE login = ?",
                               (password_hash, login))
                conn.commit()
                self.log_action(login, 'UPDATE_PASSWORD', "Пароль изменен")
                return True
        except Exception:
            return False

    def update_user_folder(self, login, folder_path):
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("UPDATE users SET folder_path = ?, updated_at = CURRENT_TIMESTAMP WHERE login = ?",
                               (folder_path or '', login))
                conn.commit()
                return True
        except Exception:
            return False

    def update_user(self, login, new_data):
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                fields, values = [], []

                for key in ['email', 'fname', 'lname', 'enterprise', 'plant', 'rights', 'status', 'folder_path',
                            'no_password_login']:
                    if key in new_data:
                        fields.append(f"{key} = ?")
                        values.append(new_data[key])

                if new_data.get('password'):
                    fields.append("password_hash = ?")
                    values.append(generate_password_hash(new_data['password']))

                if not fields:
                    return True, "Нет данных для обновления"

                values.append(login)
                fields.append("updated_at = CURRENT_TIMESTAMP")
                query = f"UPDATE users SET {', '.join(fields)} WHERE login = ?"
                cursor.execute(query, values)
                conn.commit()
                self.log_action(login, 'UPDATE_USER', "Данные пользователя обновлены")
                return True, "Данные обновлены"
        except Exception as e:
            return False, f"Ошибка: {str(e)}"

    def delete_user(self, login):
        admins = self.get_users_by_rights('Администратор')
        user = self.find_user(login)
        if user and user['rights'] == 'Администратор' and len(admins) <= 1:
            return False, "Нельзя удалить единственного администратора"
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("UPDATE users SET status = 'deleted', updated_at = CURRENT_TIMESTAMP WHERE login = ?",
                               (login,))
                conn.commit()
                self.log_action(login, 'DELETE_USER', "Пользователь удален")
                return True, "Пользователь удален"
        except Exception as e:
            return False, f"Ошибка: {str(e)}"

    def block_user(self, login):
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("UPDATE users SET status = 'blocked', updated_at = CURRENT_TIMESTAMP WHERE login = ?",
                               (login,))
                conn.commit()
                self.log_action(login, 'BLOCK_USER', "Пользователь заблокирован")
                return True, "Пользователь заблокирован"
        except Exception as e:
            return False, f"Ошибка: {str(e)}"

    def unblock_user(self, login):
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("UPDATE users SET status = 'active', updated_at = CURRENT_TIMESTAMP WHERE login = ?",
                               (login,))
                conn.commit()
                self.log_action(login, 'UNBLOCK_USER', "Пользователь разблокирован")
                return True, "Пользователь разблокирован"
        except Exception as e:
            return False, f"Ошибка: {str(e)}"

    def get_all_users(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT id, login, email, fname, lname, enterprise, plant, rights, status, folder_path, no_password_login, created_at FROM users ORDER BY login")
            return [dict(row) for row in cursor.fetchall()]

    def get_users_by_rights(self, rights):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM users WHERE rights = ? AND status = 'active'", (rights,))
            return [dict(row) for row in cursor.fetchall()]

    def log_action(self, user_login, action, details, category='other'):
        """Логирование действия с ЯВНОЙ передачей времени"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                now_str = self._get_current_time().strftime('%Y-%m-%d %H:%M:%S')
                cursor.execute(
                    "INSERT INTO audit_log (user_login, action, details, category, timestamp) VALUES (?, ?, ?, ?, ?)",
                    (user_login, action, details, category, now_str)
                )
                conn.commit()
        except Exception as e:
            print(f"Ошибка логирования: {e}")

    def get_audit_log(self, limit=100, category_filter=None):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            if category_filter:
                cursor.execute("SELECT * FROM audit_log WHERE category = ? ORDER BY timestamp DESC LIMIT ?",
                               (category_filter, limit))
            else:
                cursor.execute("SELECT * FROM audit_log ORDER BY timestamp DESC LIMIT ?", (limit,))
            return [dict(row) for row in cursor.fetchall()]

    def get_login_audit(self, limit=100):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM login_audit ORDER BY timestamp DESC LIMIT ?", (limit,))
            return [dict(row) for row in cursor.fetchall()]

    def delete_audit_log(self, log_id):
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM audit_log WHERE id = ?", (log_id,))
                conn.commit()
                return True, "Запись удалена"
        except Exception as e:
            return False, f"Ошибка при удалении записи: {str(e)}"
# app/models/user_db.py
import os
import sqlite3
import secrets
from contextlib import contextmanager
from werkzeug.security import generate_password_hash, check_password_hash


class UserDatabase:
    """Класс для работы с зашифрованной базой данных пользователей"""

    def __init__(self, db_file, master_key_file):
        self.db_file = db_file
        self.master_key_file = master_key_file
        self.master_key = self._get_master_key()
        self.init_database()
        print(f"Users DB: {self.db_file}")

    def _get_master_key(self):
        """Получает или создает мастер-ключ для шифрования"""
        if os.path.exists(self.master_key_file):
            with open(self.master_key_file, 'rb') as f:
                return f.read()
        else:
            key = secrets.token_bytes(32)
            os.makedirs(os.path.dirname(self.master_key_file), exist_ok=True)
            with open(self.master_key_file, 'wb') as f:
                f.write(key)
            return key

    def _encrypt(self, data):
        """Шифрует данные (XOR)"""
        if not data:
            return ''
        key = self.master_key
        encrypted = bytearray()
        for i, char in enumerate(str(data).encode('utf-8')):
            encrypted.append(char ^ key[i % len(key)])
        return encrypted.hex()

    def _decrypt(self, encrypted_data):
        """Расшифровывает данные"""
        if not encrypted_data:
            return ''
        key = self.master_key
        encrypted_bytes = bytes.fromhex(encrypted_data)
        decrypted = bytearray()
        for i, byte in enumerate(encrypted_bytes):
            decrypted.append(byte ^ key[i % len(key)])
        return decrypted.decode('utf-8')

    @contextmanager
    def get_connection(self):
        """Получает соединение с базой данных"""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
        finally:
            conn.close()

    def init_database(self):
        """Инициализирует базу данных и создает таблицы"""
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
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            cursor.execute("PRAGMA table_info(users)")
            columns = [col[1] for col in cursor.fetchall()]
            if 'status' not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN status TEXT DEFAULT 'active'")
                conn.commit()
                cursor.execute("UPDATE users SET status = 'active' WHERE status IS NULL")
                conn.commit()

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS audit_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_login TEXT,
                    action TEXT,
                    details TEXT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            conn.commit()

            cursor.execute("SELECT COUNT(*) FROM users")
            count = cursor.fetchone()[0]

            if count == 0:
                self._create_default_users()

    def _create_default_users(self):
        """Создает пользователей по умолчанию"""
        default_users = [
            {
                'login': 'admin',
                'email': 'admin@example.com',
                'password': 'admin123',
                'fname': 'Администратор',
                'lname': 'Системы',
                'enterprise': 'ООО "Верофарм"',
                'plant': 'Вольгинский',
                'rights': 'Администратор',
                'status': 'active'
            },
            {
                'login': 'operator',
                'email': 'operator@example.com',
                'password': 'operator123',
                'fname': 'Оператор',
                'lname': 'СиМ',
                'enterprise': 'ООО "Верофарм"',
                'plant': 'Вольгинский',
                'rights': 'Оператор СиМ',
                'status': 'active'
            },
            {
                'login': 'alexander.abramov@abbott.com',
                'email': 'alexander.abramov@abbott.com',
                'password': 'alexander.abramov@abbott.com',
                'fname': 'Александр',
                'lname': 'Абрамов',
                'enterprise': 'ООО "Верофарм"',
                'plant': 'Вольгинский',
                'rights': 'Администратор',
                'status': 'active'
            },
            {
                'login': 'anastasia.yudashkina@abbott.com',
                'email': 'anastasia.yudashkina@abbott.com',
                'password': 'anastasia.yudashkina@abbott.com',
                'fname': 'Анастасия',
                'lname': 'Юдашкина',
                'enterprise': 'ООО "Верофарм"',
                'plant': 'Вольгинский',
                'rights': 'Оператор СиМ',
                'status': 'active'
            }
        ]

        for user in default_users:
            self.add_user(user)

    def add_user(self, user_data):
        """Добавляет нового пользователя в базу"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                password_hash = generate_password_hash(user_data['password'])

                cursor.execute('''
                    INSERT INTO users (login, email, password_hash, fname, lname, enterprise, plant, rights, status)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    user_data['login'],
                    user_data['email'],
                    password_hash,
                    user_data.get('fname', ''),
                    user_data.get('lname', ''),
                    user_data.get('enterprise', ''),
                    user_data.get('plant', ''),
                    user_data.get('rights', 'Оператор СиМ'),
                    user_data.get('status', 'active')
                ))
                conn.commit()
                self.log_action(user_data['login'], 'ADD_USER', f"Добавлен пользователь {user_data['login']}")
                return True, "Пользователь успешно добавлен"
        except sqlite3.IntegrityError:
            return False, "Пользователь с таким логином уже существует"
        except Exception as e:
            return False, f"Ошибка при добавлении пользователя: {str(e)}"

    def find_user(self, login):
        """Находит пользователя по логину"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM users WHERE login = ?", (login,))
            row = cursor.fetchone()
            if row:
                return dict(row)
            return None

    def authenticate(self, login, password):
        """Проверяет логин и пароль"""
        user = self.find_user(login)
        if user:
            if user.get('status') == 'blocked':
                self.log_action(login, 'LOGIN_FAILED', f"Попытка входа заблокированного пользователя")
                return None, 'blocked'
            if user.get('status') == 'deleted':
                self.log_action(login, 'LOGIN_FAILED', f"Попытка входа удаленного пользователя")
                return None, 'deleted'

            if check_password_hash(user['password_hash'], password):
                self.log_action(login, 'LOGIN_SUCCESS', f"Успешный вход")
                return user, 'success'
            else:
                self.log_action(login, 'LOGIN_FAILED', f"Неверный пароль")
                return None, 'invalid'
        return None, 'not_found'

    def update_password(self, login, new_password):
        """Обновляет пароль пользователя"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                password_hash = generate_password_hash(new_password)
                cursor.execute(
                    "UPDATE users SET password_hash = ?, updated_at = CURRENT_TIMESTAMP WHERE login = ?",
                    (password_hash, login)
                )
                conn.commit()
                self.log_action(login, 'UPDATE_PASSWORD', "Пароль изменен")
                return True
        except Exception as e:
            print(f"Ошибка обновления пароля: {e}")
            return False

    def update_user(self, login, new_data):
        """Обновляет данные пользователя"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()

                fields = []
                values = []

                for key in ['email', 'fname', 'lname', 'enterprise', 'plant', 'rights', 'status']:
                    if key in new_data:
                        fields.append(f"{key} = ?")
                        values.append(new_data[key])

                if 'password' in new_data and new_data['password']:
                    fields.append("password_hash = ?")
                    values.append(generate_password_hash(new_data['password']))

                values.append(login)
                fields.append("updated_at = CURRENT_TIMESTAMP")

                query = f"UPDATE users SET {', '.join(fields)} WHERE login = ?"
                cursor.execute(query, values)
                conn.commit()

                self.log_action(login, 'UPDATE_USER', f"Данные пользователя обновлены")
                return True, "Данные пользователя обновлены"
        except Exception as e:
            return False, f"Ошибка обновления: {str(e)}"

    def delete_user(self, login):
        """Удаляет пользователя (soft delete - меняет статус)"""
        admins = self.get_users_by_rights('Администратор')
        user = self.find_user(login)

        if user and user['rights'] == 'Администратор' and len(admins) <= 1:
            return False, "Нельзя удалить единственного администратора"

        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "UPDATE users SET status = 'deleted', updated_at = CURRENT_TIMESTAMP WHERE login = ?",
                    (login,)
                )
                conn.commit()
                self.log_action(login, 'DELETE_USER', f"Пользователь удален")
                return True, "Пользователь удален"
        except Exception as e:
            return False, f"Ошибка удаления: {str(e)}"

    def block_user(self, login):
        """Блокирует пользователя"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "UPDATE users SET status = 'blocked', updated_at = CURRENT_TIMESTAMP WHERE login = ?",
                    (login,)
                )
                conn.commit()
                self.log_action(login, 'BLOCK_USER', f"Пользователь заблокирован")
                return True, "Пользователь заблокирован"
        except Exception as e:
            return False, f"Ошибка блокировки: {str(e)}"

    def unblock_user(self, login):
        """Разблокирует пользователя"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "UPDATE users SET status = 'active', updated_at = CURRENT_TIMESTAMP WHERE login = ?",
                    (login,)
                )
                conn.commit()
                self.log_action(login, 'UNBLOCK_USER', f"Пользователь разблокирован")
                return True, "Пользователь разблокирован"
        except Exception as e:
            return False, f"Ошибка разблокировки: {str(e)}"

    def get_all_users(self):
        """Возвращает список всех пользователей"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT id, login, email, fname, lname, enterprise, plant, rights, status, created_at FROM users ORDER BY login"
            )
            rows = cursor.fetchall()
            return [dict(row) for row in rows]

    def get_users_by_rights(self, rights):
        """Возвращает пользователей с определенными правами"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM users WHERE rights = ? AND status = 'active'", (rights,))
            rows = cursor.fetchall()
            return [dict(row) for row in rows]

    def log_action(self, user_login, action, details):
        """Логирует действие пользователя"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO audit_log (user_login, action, details)
                    VALUES (?, ?, ?)
                ''', (user_login, action, details))
                conn.commit()
        except Exception as e:
            print(f"Ошибка логирования: {e}")

    def get_audit_log(self, limit=100):
        """Возвращает последние записи аудита"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM audit_log ORDER BY timestamp DESC LIMIT ?",
                (limit,)
            )
            rows = cursor.fetchall()
            return [dict(row) for row in rows]
# app/models/transport_db.py
import os
import sqlite3
import re
from contextlib import contextmanager
from datetime import datetime


class TransportDatabase:
    def __init__(self, db_file):
        self.db_file = db_file
        self.init_database()
        print(f"Transport DB: {self.db_file}")

    @contextmanager
    def get_connection(self):
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
        finally:
            conn.close()

    def init_database(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            for table in ['transport_working', 'transport_training']:
                cursor.execute(f'''
                    CREATE TABLE IF NOT EXISTS {table} (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        order_number TEXT NOT NULL,
                        order_date TEXT NOT NULL,
                        sender TEXT,
                        loading_place TEXT,
                        loading_contact TEXT,
                        delivery_date TEXT,
                        delivery_time TEXT,
                        receiver TEXT,
                        delivery_address TEXT,
                        unloading_contact TEXT,
                        cargo_name TEXT,
                        quantity INTEGER,
                        vehicle_type TEXT,
                        temperature_mode TEXT,
                        comments TEXT,
                        driver_name TEXT,
                        driver_data TEXT,
                        driver_phone TEXT,
                        vehicle_number TEXT,
                        author TEXT,
                        po INTEGER,
                        loading_date TEXT,
                        loading_time TEXT,
                        supply_date TEXT,
                        supply_time TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
                cursor.execute(f"PRAGMA table_info({table})")
                columns = [col[1] for col in cursor.fetchall()]
                if 'supply_date' not in columns:
                    cursor.execute(f"ALTER TABLE {table} ADD COLUMN supply_date TEXT")
                if 'supply_time' not in columns:
                    cursor.execute(f"ALTER TABLE {table} ADD COLUMN supply_time TEXT")
            conn.commit()

    def copy_working_to_training(self):
        """Копирует все записи из рабочей таблицы в учебную (предварительно очищая учебную)"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM transport_training")
            cursor.execute('''
                INSERT INTO transport_training (
                    order_number, order_date, sender, loading_place, loading_contact,
                    delivery_date, delivery_time, receiver, delivery_address,
                    unloading_contact, cargo_name, quantity, vehicle_type,
                    temperature_mode, comments, driver_name, driver_data,
                    driver_phone, vehicle_number, author, po, loading_date, loading_time,
                    supply_date, supply_time
                )
                SELECT 
                    order_number, order_date, sender, loading_place, loading_contact,
                    delivery_date, delivery_time, receiver, delivery_address,
                    unloading_contact, cargo_name, quantity, vehicle_type,
                    temperature_mode, comments, driver_name, driver_data,
                    driver_phone, vehicle_number, author, po, loading_date, loading_time,
                    supply_date, supply_time
                FROM transport_working
            ''')
            conn.commit()
            return True

    def get_max_order_number(self, table_name):
        """
        Возвращает максимальный числовой номер заявки в таблице.
        Вместо того чтобы полагаться на дату создания (ID), мы сканируем
        все номера и извлекаем из них числа (например, из "139/1" берется 139).
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(f"SELECT order_number FROM {table_name}")
            rows = cursor.fetchall()

            max_num = 0
            for row in rows:
                num_str = str(row['order_number'])
                match = re.match(r'^(\d+)', num_str)
                if match:
                    num = int(match.group(1))
                    if num > max_num:
                        max_num = num
            return max_num

    def add_request(self, table_name, data):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(f'''
                INSERT INTO {table_name} (
                    order_number, order_date, sender, loading_place, loading_contact,
                    delivery_date, delivery_time, receiver, delivery_address,
                    unloading_contact, cargo_name, quantity, vehicle_type,
                    temperature_mode, comments, driver_name, driver_data,
                    driver_phone, vehicle_number, author, po, loading_date, loading_time,
                    supply_date, supply_time
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data['order_number'], data['order_date'], data['sender'],
                data['loading_place'], data['loading_contact'], data['delivery_date'],
                data['delivery_time'], data['receiver'], data['delivery_address'],
                data['unloading_contact'], data['cargo_name'], data['quantity'],
                data['vehicle_type'], data['temperature_mode'], data['comments'],
                data['driver_name'], data['driver_data'], data['driver_phone'],
                data['vehicle_number'], data['author'], data['po'],
                data.get('loading_date', ''), data.get('loading_time', ''),
                data.get('supply_date', ''), data.get('supply_time', '')
            ))
            conn.commit()
            return cursor.lastrowid

    def get_all_requests(self, table_name):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(f"SELECT * FROM {table_name}")
            rows = cursor.fetchall()
            res = [dict(row) for row in rows]

            def sort_key(item):
                date_val = item['order_date']
                num_str = str(item['order_number'])
                match = re.match(r'^(\d+)', num_str)
                num_val = int(match.group(1)) if match else 0
                return (date_val, num_val)

            res.sort(key=sort_key, reverse=True)
            return res

    def update_request(self, table_name, req_id, data):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            fields = []
            values = []
            for key in data.keys():
                if key not in ['id', 'created_at']:
                    fields.append(f"{key} = ?")
                    values.append(data[key])
            values.append(req_id)
            query = f"UPDATE {table_name} SET {', '.join(fields)} WHERE id = ?"
            cursor.execute(query, values)
            conn.commit()
            return True

    def delete_request(self, table_name, req_id):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM {table_name} WHERE id = ?", (req_id,))
            conn.commit()
            return True
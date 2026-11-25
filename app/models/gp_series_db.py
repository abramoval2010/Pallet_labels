# app/models/gp_series_db.py
import os
import sqlite3
import math
from contextlib import contextmanager
from datetime import datetime, timedelta
from openpyxl import load_workbook


class GpSeriesDatabase:
    def __init__(self, db_file):
        self.db_file = db_file
        self.init_database()
        print(f"GP Series DB: {self.db_file}")

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
            for table in ['gp_series', 'gp_series_training']:
                cursor.execute(f'''
                    CREATE TABLE IF NOT EXISTS {table} (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        manufacture_date TEXT NOT NULL,
                        product_code INTEGER NOT NULL,
                        batch INTEGER NOT NULL,
                        name TEXT NOT NULL,
                        storage_temp TEXT NOT NULL,
                        humidity TEXT NOT NULL,
                        qc_stock INTEGER DEFAULT 0,
                        free_stock INTEGER DEFAULT 0,
                        blocked_stock INTEGER DEFAULT 0,
                        expiry_date TEXT NOT NULL,
                        palletization INTEGER,
                        pallets INTEGER,
                        sampling_done INTEGER DEFAULT 0,
                        sap_writeoff_done INTEGER DEFAULT 0,
                        ul_status TEXT DEFAULT 'На контроле',
                        shipped INTEGER DEFAULT 0,
                        comments TEXT,
                        shipped_by TEXT,
                        shipped_date TEXT,
                        shipped_time TEXT
                    )
                ''')
                cursor.execute(f"PRAGMA table_info({table})")
                columns = [col[1] for col in cursor.fetchall()]
                if 'shipped_by' not in columns:
                    cursor.execute(f"ALTER TABLE {table} ADD COLUMN shipped_by TEXT")
                if 'shipped_date' not in columns:
                    cursor.execute(f"ALTER TABLE {table} ADD COLUMN shipped_date TEXT")
                if 'shipped_time' not in columns:
                    cursor.execute(f"ALTER TABLE {table} ADD COLUMN shipped_time TEXT")
            conn.commit()

    def get_filtered(self, table_name, filter_type='available', sort_by='id', sort_order='desc'):
        allowed_sort_columns = {
            'product_code': 'product_code',
            'name': 'name',
            'manufacture_date': 'manufacture_date',
            'expiry_date': 'expiry_date',
            'storage_temp': 'storage_temp',
            'palletization': 'palletization',
            'ul_status': 'ul_status',
            'sampling_done': 'sampling_done',
            'sap_writeoff_done': 'sap_writeoff_done',
            'id': 'id'
        }
        sort_column = allowed_sort_columns.get(sort_by, 'id')
        order = 'ASC' if sort_order.lower() == 'asc' else 'DESC'

        with self.get_connection() as conn:
            cursor = conn.cursor()

            if filter_type == 'qc_pending':
                cursor.execute(f"""
                    SELECT * FROM {table_name} 
                    WHERE ul_status = 'На контроле' 
                      AND (sampling_done = 0 OR sap_writeoff_done = 0)
                    ORDER BY {sort_column} {order}
                """)
            elif filter_type == 'ul_pending':
                # Фильтр для Уполномоченного лица: На контроле или Блок, И Отбор=1, И Списание=1
                cursor.execute(f"""
                    SELECT * FROM {table_name} 
                    WHERE (ul_status = 'На контроле' OR ul_status = 'Блок')
                      AND sampling_done = 1 
                      AND sap_writeoff_done = 1
                    ORDER BY {sort_column} {order}
                """)
            elif filter_type == 'can_ship':
                cursor.execute(f"""
                    SELECT * FROM {table_name} 
                    WHERE shipped = 0 
                      AND ul_status IN ('Ограниченное использование', 'Свободное использование')
                    ORDER BY {sort_column} {order}
                """)
            elif filter_type == 'shipped':
                cursor.execute(f"SELECT * FROM {table_name} WHERE shipped = 1 ORDER BY {sort_column} {order}")
            else:
                cursor.execute(f"SELECT * FROM {table_name} WHERE shipped = 0 ORDER BY {sort_column} {order}")

            return [dict(row) for row in cursor.fetchall()]

    def get_statistics(self, table_name, filter_type='available'):
        with self.get_connection() as conn:
            cursor = conn.cursor()

            where_clause = ""
            if filter_type == 'qc_pending':
                where_clause = "AND ul_status = 'На контроле' AND (sampling_done = 0 OR sap_writeoff_done = 0)"
            elif filter_type == 'ul_pending':
                where_clause = "AND (ul_status = 'На контроле' OR ul_status = 'Блок') AND sampling_done = 1 AND sap_writeoff_done = 1"
            elif filter_type == 'can_ship':
                where_clause = "AND ul_status IN ('Ограниченное использование', 'Свободное использование')"
            elif filter_type == 'shipped':
                where_clause = "AND shipped = 1"
            else:
                where_clause = "AND shipped = 0"

            cursor.execute(f"""
                SELECT COUNT(*) as count, SUM(pallets) as total_pallets 
                FROM {table_name} 
                WHERE storage_temp = '+15;+25' {where_clause}
            """)
            warm = cursor.fetchone()

            cursor.execute(f"""
                SELECT COUNT(*) as count, SUM(pallets) as total_pallets 
                FROM {table_name} 
                WHERE storage_temp = '+2;+8' {where_clause}
            """)
            cold = cursor.fetchone()

            return {
                'warm': {'count': warm['count'] or 0, 'pallets': warm['total_pallets'] or 0},
                'cold': {'count': cold['count'] or 0, 'pallets': cold['total_pallets'] or 0}
            }

    def get_by_product_and_batch(self, table_name, product_code, batch):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(f"SELECT * FROM {table_name} WHERE product_code = ? AND batch = ?", (product_code, batch))
            row = cursor.fetchone()
            return dict(row) if row else None

    def update_series(self, table_name, series_id, data):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            fields = []
            values = []
            for key, val in data.items():
                fields.append(f"{key} = ?")
                values.append(val)
            values.append(series_id)
            query = f"UPDATE {table_name} SET {', '.join(fields)} WHERE id = ?"
            cursor.execute(query, values)
            conn.commit()
            return True

    def clear_series(self, table_name):
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
                count = cursor.fetchone()[0]
                if count == 0:
                    return False, "База уже пуста. Удалять нечего."
                cursor.execute(f"DELETE FROM {table_name}")
                conn.commit()
                return True, f"База полностью очищена. Удалено записей: {count}"
        except Exception as e:
            return False, f"Ошибка очистки: {str(e)}"

    def ship_series(self, table_name, series_ids, user_name, current_time_str):
        if not series_ids:
            return False, "Не выбрано ни одной серии для отгрузки.", 0, 0
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                placeholders = ','.join('?' for _ in series_ids)
                cursor.execute(f"SELECT SUM(pallets) FROM {table_name} WHERE id IN ({placeholders})", series_ids)
                total_pallets = cursor.fetchone()[0] or 0

                date_part = current_time_str.split(' ')[0]
                time_part = current_time_str.split(' ')[1] if ' ' in current_time_str else ''

                cursor.execute(f"""
                    UPDATE {table_name} 
                    SET shipped = 1, shipped_by = ?, shipped_date = ?, shipped_time = ?
                    WHERE id IN ({placeholders})
                """, [user_name, date_part, time_part] + series_ids)

                conn.commit()
                return True, f"Успешно отгружено {len(series_ids)} серий ({total_pallets} паллет).", len(
                    series_ids), total_pallets
        except Exception as e:
            return False, f"Ошибка при отгрузке: {str(e)}", 0, 0

    def import_from_excel(self, filepath, table_name, materials_db):
        # (код метода без изменений, вставьте его из предыдущей версии)
        pass  # Заглушка, чтобы не раздувать ответ, но в реальном файле он должен быть полностью
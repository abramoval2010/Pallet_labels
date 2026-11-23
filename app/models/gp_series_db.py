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
                # Специальный фильтр для "Контрольного мастера"
                cursor.execute(f"""
                    SELECT * FROM {table_name} 
                    WHERE ul_status = 'На контроле' 
                      AND (sampling_done = 0 OR sap_writeoff_done = 0)
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
        # (код метода import_from_excel остаётся без изменений)
        try:
            wb = load_workbook(filepath, data_only=True)
            ws = wb.active

            if ws.max_row < 2:
                return False, "Файл пуст или не содержит данных.", 0, 0

            headers_map = {}
            for cell in ws[1]:
                if cell.value:
                    clean_name = str(cell.value).strip().lower()
                    headers_map[clean_name] = cell.column

            required_cols = {
                'материал': 'product_code',
                'партия': 'batch',
                'свобиспольззапас': 'free_stock',
                'на контркачества': 'qc_stock',
                'блокированный запас': 'blocked_stock',
                'срок хранения или минимальный срок годно': 'expiry_date',
                'дата изготовления': 'manufacture_date',
                'краткий текст материала': 'name'
            }

            col_indices = {}
            found_errors = []

            for search_name, var_name in required_cols.items():
                found = False
                for header, idx in headers_map.items():
                    if search_name in header:
                        col_indices[var_name] = idx
                        found = True
                        break
                if not found:
                    found_errors.append(f"❌ Не найдена колонка: '{search_name}'")

            if found_errors:
                return False, "Неверный формат файла:\n" + "\n".join(found_errors), 0, 0

            row_errors = []
            max_rows_to_check = min(ws.max_row, 100)

            for row_idx in range(2, max_rows_to_check + 1):
                if all(ws.cell(row=row_idx, column=i).value is None for i in range(1, ws.max_column + 1)):
                    continue

                for var_name, readable_name in [('free_stock', 'СвобИспользЗапас'), ('qc_stock', 'На КонтрКачества'),
                                                ('blocked_stock', 'Блокированный запас')]:
                    if var_name in col_indices:
                        val = ws.cell(row=row_idx, column=col_indices[var_name]).value
                        if val is not None and str(val).strip() != "":
                            try:
                                if float(val) < 0:
                                    row_errors.append(f"❌ Строка {row_idx}: '{readable_name}' не может быть меньше 0")
                            except (ValueError, TypeError):
                                row_errors.append(
                                    f"❌ Строка {row_idx}: '{readable_name}' должно быть числом (найдено: '{val}')")

                for var_name, readable_name in [('expiry_date', 'Срок хранения'),
                                                ('manufacture_date', 'Дата изготовления')]:
                    if var_name in col_indices:
                        val = ws.cell(row=row_idx, column=col_indices[var_name]).value
                        if val is not None and str(val).strip() != "":
                            if not isinstance(val, datetime):
                                try:
                                    float(val)
                                except:
                                    row_errors.append(
                                        f"❌ Строка {row_idx}: '{readable_name}' имеет неверный формат (найдено: '{val}')")

                if len(row_errors) > 10:
                    row_errors.append("... (и другие ошибки в следующих строках)")
                    break

            if row_errors:
                return False, "Неверный формат файла:\n" + "\n".join(row_errors), 0, 0

            imported_count = 0
            updated_count = 0

            for row_idx in range(2, ws.max_row + 1):
                if all(ws.cell(row=row_idx, column=i).value is None for i in range(1, ws.max_column + 1)):
                    continue

                product_code_raw = ws.cell(row=row_idx, column=col_indices.get('product_code', 1)).value
                batch_raw = ws.cell(row=row_idx, column=col_indices.get('batch', 8)).value

                if not product_code_raw or not batch_raw:
                    continue

                try:
                    product_code = int(float(product_code_raw))
                    batch = int(float(batch_raw))
                except (ValueError, TypeError):
                    continue

                def get_val(var_name, default=0):
                    idx = col_indices.get(var_name)
                    if not idx: return default
                    val = ws.cell(row=row_idx, column=idx).value
                    if val is None or str(val).strip() == "": return default
                    try:
                        return int(float(val))
                    except:
                        return default

                free_stock = get_val('free_stock', 0)
                qc_stock = get_val('qc_stock', 0)
                blocked_stock = get_val('blocked_stock', 0)

                def format_date(var_name):
                    idx = col_indices.get(var_name)
                    if not idx: return ''
                    val = ws.cell(row=row_idx, column=idx).value
                    if isinstance(val, datetime):
                        return val.strftime('%Y-%m-%d')
                    elif val:
                        try:
                            num = float(val)
                            return (datetime(1899, 12, 30) + timedelta(days=num)).strftime('%Y-%m-%d')
                        except:
                            return str(val)[:10]
                    return ''

                expiry_date = format_date('expiry_date')
                manuf_date = format_date('manufacture_date')

                name_idx = col_indices.get('name', 34)
                name = str(ws.cell(row=row_idx, column=name_idx).value or 'Не указано')

                material = materials_db.get_material_by_sap(product_code)
                if material:
                    temp = material.get('storage_temp', '+15;+25')
                    hum = material.get('humidity', '0-65%')
                    palletization = material.get('palletization', 0)
                else:
                    temp = '+15;+25'
                    hum = '0-65%'
                    palletization = 0

                total_stock = free_stock + qc_stock + blocked_stock
                if palletization > 0:
                    pallets = math.ceil(total_stock / palletization)
                else:
                    pallets = 0

                existing = self.get_by_product_and_batch(table_name, product_code, batch)

                if existing:
                    update_data = {
                        'qc_stock': qc_stock,
                        'free_stock': free_stock,
                        'blocked_stock': blocked_stock,
                        'storage_temp': temp,
                        'humidity': hum,
                        'palletization': palletization,
                        'pallets': pallets
                    }
                    self.update_series(table_name, existing['id'], update_data)
                    updated_count += 1
                else:
                    with self.get_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute(f'''
                            INSERT INTO {table_name} (
                                manufacture_date, product_code, batch, name, storage_temp, humidity,
                                qc_stock, free_stock, blocked_stock, expiry_date, palletization, pallets,
                                sampling_done, sap_writeoff_done, ul_status, shipped, comments
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, 'На контроле', 0, '')
                        ''', (manuf_date, product_code, batch, name, temp, hum, qc_stock, free_stock, blocked_stock,
                              expiry_date, palletization, pallets))
                        conn.commit()
                        imported_count += 1

            return True, f"Успешно загружено: {imported_count} новых, обновлено: {updated_count}.", imported_count, updated_count

        except Exception as e:
            print(f"Ошибка импорта: {e}")
            import traceback
            traceback.print_exc()
            return False, f"Ошибка при чтении файла: {str(e)}", 0, 0
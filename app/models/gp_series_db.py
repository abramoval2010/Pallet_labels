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
                        comments TEXT
                    )
                ''')
            conn.commit()

    def get_all(self, table_name):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(f"SELECT * FROM {table_name} WHERE shipped = 0 ORDER BY id DESC")
            return [dict(row) for row in cursor.fetchall()]

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

    def clear_series(self, table_name, only_shipped=False):
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                if only_shipped:
                    cursor.execute(f"DELETE FROM {table_name} WHERE shipped = 1")
                else:
                    cursor.execute(f"DELETE FROM {table_name}")
                conn.commit()
                return True, f"Удалено записей: {cursor.rowcount}"
        except Exception as e:
            return False, f"Ошибка очистки: {str(e)}"

    def import_from_excel(self, filepath, table_name, materials_db):
        """Парсинг и импорт Excel файла с сериями ГП (с подробной валидацией)"""
        try:
            wb = load_workbook(filepath, data_only=True)
            ws = wb.active

            if ws.max_row < 2:
                return False, "Файл пуст или не содержит данных."

            # 1. Поиск заголовков по названиям
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
                return False, "Неверный формат файла:\n" + "\n".join(found_errors)

            # 2. Проверка данных в строках
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
                return False, "Неверный формат файла:\n" + "\n".join(row_errors)

            # 3. Импорт данных
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

            return True, f"Успешно загружено: {imported_count} новых, обновлено: {updated_count}."

        except Exception as e:
            print(f"Ошибка импорта: {e}")
            import traceback
            traceback.print_exc()
            return False, f"Ошибка при чтении файла: {str(e)}"
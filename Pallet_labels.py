# app.py
from flask import Flask, render_template, request, redirect, url_for, send_file
import os
import re
import fitz  # PyMuPDF
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
import subprocess
import json
from pathlib import Path
import sys
import tempfile
import shutil
from datetime import datetime
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['SETTINGS_FILE'] = 'settings.json'

# Создаем папку для загрузок если её нет
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Файл настроек
SETTINGS_FILE = 'settings.json'


def load_settings():
    """Загружает настройки из файла"""
    default_settings = {
        'save_path': os.path.join(os.path.expanduser('~'), 'Documents', 'Палетные этикетки')
    }

    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                return settings
        except:
            return default_settings
    else:
        # Создаем папку по умолчанию
        os.makedirs(default_settings['save_path'], exist_ok=True)
        return default_settings


def save_settings(settings):
    """Сохраняет настройки в файл"""
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=4)
        return True
    except:
        return False


def get_available_drives():
    """Получает список доступных дисков в Windows"""
    drives = []
    if sys.platform == 'win32':
        import string
        for letter in string.ascii_uppercase:
            drive = f"{letter}:\\"
            if os.path.exists(drive):
                drives.append(drive)
    else:
        drives = ['/']
    return drives


def create_pallet_labels_file(all_products, order_number, save_path=None):
    """
    Создает файл с несколькими палетными этикетками

    Args:
        all_products: список словарей с данными товаров
        order_number: номер приходного ордера
        save_path: путь для сохранения

    Returns:
        str: путь к созданному файлу
    """
    if save_path is None:
        settings = load_settings()
        save_path = settings.get('save_path', os.path.join(os.path.expanduser('~'), 'Documents', 'Палетные этикетки'))

    # Создаем папку если её нет
    os.makedirs(save_path, exist_ok=True)

    # Создаем новый документ
    merged_doc = Document()

    # Устанавливаем ЛАНДШАФТНУЮ ориентацию страницы и уменьшенные поля
    for section in merged_doc.sections:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Inches(11.69)  # A4 landscape width
        section.page_height = Inches(8.27)  # A4 landscape height
        section.top_margin = Inches(0.15)
        section.bottom_margin = Inches(0.15)
        section.left_margin = Inches(0.15)
        section.right_margin = Inches(0.15)

    # Для каждого товара создаем этикетку
    for idx, product in enumerate(all_products):
        # Добавляем разрыв страницы между этикетками (кроме первой)
        if idx > 0:
            merged_doc.add_page_break()
            # Для каждой новой страницы устанавливаем ландшафтную ориентацию
            for section in merged_doc.sections:
                section.orientation = WD_ORIENT.LANDSCAPE
                section.page_width = Inches(11.69)
                section.page_height = Inches(8.27)
                section.top_margin = Inches(0.15)
                section.bottom_margin = Inches(0.15)
                section.left_margin = Inches(0.15)
                section.right_margin = Inches(0.15)

        # Получаем данные товара
        product_name = product.get('наименование', 'Не указано')
        supplier_series = product.get('партия_поставщика', 'Не указана')
        analytic_list = product.get('аналитический_лист', 'Не указан')

        # Определяем размеры шрифтов в зависимости от длины названия
        name_len = len(product_name)

        if name_len < 30:
            # Короткое название - большие шрифты
            font_sizes = {
                'header': 14,
                'product': 82,
                'supplier': 60,
                'analytic': 76,
                'code': 14
            }
        elif name_len < 40:
            # Среднее название (30-39 символов)
            font_sizes = {
                'header': 14,
                'product': 74,
                'supplier': 58,
                'analytic': 68,
                'code': 14
            }
        else:
            # Длинное название (40+ символов)
            font_sizes = {
                'header': 14,
                'product': 70,
                'supplier': 56,
                'analytic': 66,
                'code': 14
            }

        # Дополнительное уменьшение для очень длинных названий
        if name_len > 55:
            font_sizes['product'] = 62
            font_sizes['supplier'] = 50
            font_sizes['analytic'] = 58
        elif name_len > 70:
            font_sizes['product'] = 54
            font_sizes['supplier'] = 44
            font_sizes['analytic'] = 50
        elif name_len > 85:
            font_sizes['product'] = 46
            font_sizes['supplier'] = 38
            font_sizes['analytic'] = 44

        # Дополнительное уменьшение для длинной серии поставщика
        supplier_len = len(supplier_series) if supplier_series else 0
        if supplier_len > 25:
            font_sizes['supplier'] = max(32, font_sizes['supplier'] - 12)
        elif supplier_len > 18:
            font_sizes['supplier'] = max(38, font_sizes['supplier'] - 8)

        # Дополнительное уменьшение для длинного аналитического листа
        analytic_len = len(str(analytic_list)) if analytic_list else 0
        if analytic_len > 15:
            font_sizes['analytic'] = max(42, font_sizes['analytic'] - 10)

        # Создаем этикетку как таблицу для лучшего контроля пространства
        table = merged_doc.add_table(rows=1, cols=1)
        table.autofit = False
        table.allow_autofit = False

        # Устанавливаем ширину таблицы на всю страницу
        table.columns[0].width = Inches(10.8)

        # Получаем ячейку
        cell = table.cell(0, 0)
        cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Очищаем стандартный параграф и добавляем свои
        cell.paragraphs[0].clear()

        # 1) ООО «Верофарм» - Calibri, 14pt, жирный курсив, по центру
        p1 = cell.add_paragraph()
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_before = Pt(2)
        p1.paragraph_format.space_after = Pt(2)
        run1 = p1.add_run('ООО «Верофарм»')
        run1.italic = True
        run1.bold = True
        run1.font.size = Pt(font_sizes['header'])
        run1.font.name = 'Calibri'

        # 2) Название товара - Calibri, жирный курсив, по центру
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(5)
        p2.paragraph_format.space_after = Pt(5)
        run2 = p2.add_run(product_name)
        run2.italic = True
        run2.bold = True
        run2.font.size = Pt(font_sizes['product'])
        run2.font.name = 'Calibri'

        # 3) Серия поставщика - Calibri, курсив, по центру
        supplier_text = f"п. {supplier_series}" if supplier_series and supplier_series != 'Не указана' else supplier_series
        p3 = cell.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3.paragraph_format.space_before = Pt(4)
        p3.paragraph_format.space_after = Pt(4)
        run3 = p3.add_run(supplier_text)
        run3.italic = True
        run3.bold = False
        run3.font.size = Pt(font_sizes['supplier'])
        run3.font.name = 'Calibri'

        # 4) Аналитический лист - Calibri, жирный курсив, по центру
        p4 = cell.add_paragraph()
        p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p4.paragraph_format.space_before = Pt(4)
        p4.paragraph_format.space_after = Pt(4)
        run4 = p4.add_run(f'АЛ {analytic_list}')
        run4.italic = True
        run4.bold = True
        run4.font.size = Pt(font_sizes['analytic'])
        run4.font.name = 'Calibri'

        # 5) Код - Calibri, 14pt, жирный, с правым центрированием
        p5 = cell.add_paragraph()
        p5.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p5.paragraph_format.space_before = Pt(8)
        p5.paragraph_format.space_after = Pt(2)
        run5 = p5.add_run('RUPD.VLG.05.04.C01')
        run5.italic = False
        run5.bold = True
        run5.font.size = Pt(font_sizes['code'])
        run5.font.name = 'Calibri'

    # Формируем имя файла с временной меткой
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Информационный лист {order_number}_{timestamp}.docx"
    filepath = os.path.join(save_path, filename)

    # Сохраняем документ
    merged_doc.save(filepath)

    return filepath


def open_in_word(filepath):
    """Открывает файл в Microsoft Word"""
    try:
        if sys.platform == 'win32':
            os.startfile(filepath)
        elif sys.platform == 'darwin':  # macOS
            subprocess.run(['open', filepath])
        else:  # Linux
            subprocess.run(['xdg-open', filepath])
        return True
    except Exception as e:
        print(f"Ошибка при открытии файла: {e}")
        return False


def extract_data_from_pdf(pdf_path):
    """
    Извлекает данные из PDF-файла приходного ордера
    """
    # Открываем PDF
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()

    # Разбиваем на строки и очищаем
    lines = [line.strip() for line in text.split('\n') if line.strip()]

    # Инициализируем результат
    result = {
        'дата_приемки': None,
        'название_поставщика': None,
        'номер_ордера': None,
        'товары': []  # Список товаров
    }

    # 1) Номер приходного ордера
    for line in lines:
        if 'ПРИХОДНЫЙ ОРДЕР №' in line:
            order_match = re.search(r'№\s*(\d+)', line)
            if order_match:
                result['номер_ордера'] = order_match.group(1)
                break

    # 2) Дата приемки
    for line in lines:
        if re.match(r'\d{2}\.\d{2}\.\d{4}', line):
            if '2028' not in line and 'годности' not in line:
                result['дата_приемки'] = line
                break

    if not result['дата_приемки']:
        for i, line in enumerate(lines):
            if 'составления' in line.lower() and i + 1 < len(lines):
                date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', lines[i + 1])
                if date_match:
                    result['дата_приемки'] = date_match.group(1)
                    break

    # 3) Название поставщика
    for line in lines:
        if 'ООО' in line or 'ЗАО' in line or 'ОАО' in line or 'АО' in line or 'ИП' in line:
            if 'ВЕРОФАРМ' not in line and 'Завод' not in line:
                result['название_поставщика'] = line
                break

    # 4) Парсинг товаров из таблицы
    # Находим начало таблицы
    start_index = -1
    end_index = -1

    for i, line in enumerate(lines):
        if 'Материальные ценности' in line:
            start_index = i
            break

    for i, line in enumerate(lines):
        if line == 'Итого' or line.startswith('Итого'):
            end_index = i
            break

    if start_index != -1 and end_index != -1:
        # Находим строку с "14" - это последний номер колонки
        data_start = -1
        for i in range(start_index, end_index):
            if lines[i] == '14':
                data_start = i + 1
                break

        if data_start != -1:
            # Список заголовков, которые нужно пропускать
            header_keywords = [
                'наименование', 'сорт', 'размер', 'марка', 'код',
                'единица измерения', 'количество', 'цена', 'сумма',
                'без учета', 'всего с учетом', 'номер паспорта',
                'партия', 'поставщика', 'партия поставщика',
                'срок годности', 'паспорта'
            ]

            i = data_start
            products = []

            while i < end_index:
                line = lines[i]

                # Пропускаем пустые строки
                if not line:
                    i += 1
                    continue

                # Пропускаем даты
                if re.match(r'\d{2}\.\d{2}\.\d{4}', line):
                    i += 1
                    continue

                # Пропускаем числа
                if re.match(r'^[\d\s,]+$', line) and len(line) > 3:
                    i += 1
                    continue

                # Пропускаем "X"
                if line == 'X':
                    i += 1
                    continue

                # Проверяем, является ли строка заголовком
                is_header = False
                line_lower = line.lower()
                for keyword in header_keywords:
                    if keyword in line_lower:
                        is_header = True
                        break

                if is_header:
                    i += 1
                    continue

                # Проверяем, что строка содержит буквы и это не числа - это начало названия товара
                if re.search(r'[А-ЯЁA-Z]', line) and not re.match(r'^[\d\s,]+$', line):
                    # Начинаем сбор названия товара
                    product_name_lines = [line]
                    current_pos = i + 1
                    last_name_line = i

                    # Собираем название, пока не встретим 8-значное число (номенклатурный номер)
                    while current_pos < end_index:
                        next_line = lines[current_pos]

                        # Проверяем, является ли строка 8-значным числом (номенклатурный номер)
                        nomencl_match = re.search(r'\b(\d{8})\b', next_line)
                        if nomencl_match:
                            if not nomencl_match.group(1).startswith(('8', '7')):
                                if nomencl_match.group(1) not in ['48797491', '21000101']:
                                    # Это номенклатурный номер, название закончилось
                                    last_name_line = current_pos - 1
                                    break

                        # Если строка содержит буквы и не является служебной, добавляем к названию
                        if re.search(r'[А-ЯЁA-Z]', next_line) and not re.match(r'^[\d\s,]+$', next_line):
                            # Проверяем, что это не служебное слово
                            skip_words = ['кг', 'шт', 'г', 'л', 'м', 'уп', 'пач', 'итого',
                                          'x', 'принял', 'сдал', 'должность', 'подпись',
                                          'расшифровка', 'номер', 'паспорта']
                            if next_line.lower() not in skip_words:
                                product_name_lines.append(next_line)
                                current_pos += 1
                            else:
                                break
                        else:
                            break

                    # Если не нашли номенклатурный номер, проверяем что последняя строка с названием - current_pos-1
                    if last_name_line == i:
                        last_name_line = current_pos - 1

                    # Объединяем название товара
                    product_name = ' '.join(product_name_lines)

                    # Номенклатурный номер находится в строке last_name_line + 1
                    nomencl_number = None
                    nomencl_line = last_name_line + 1
                    if nomencl_line < len(lines):
                        nomencl_match = re.search(r'\b(\d{8})\b', lines[nomencl_line])
                        if nomencl_match:
                            if not nomencl_match.group(1).startswith(('8', '7')):
                                if nomencl_match.group(1) not in ['48797491', '21000101']:
                                    nomencl_number = nomencl_match.group(1)

                    # Аналитический лист находится в строке last_name_line + 8
                    analytic_list = None
                    batch_line = last_name_line + 8
                    if batch_line < len(lines):
                        # Проверяем, что это 10-значное число, начинающееся с 1
                        batch_match = re.search(r'\b(1\d{9})\b', lines[batch_line])
                        if batch_match:
                            analytic_list = batch_match.group(1)

                    # Партия поставщика находится в строке last_name_line + 9
                    supplier_batch = None
                    supplier_line = last_name_line + 9
                    if supplier_line < len(lines):
                        # Берем значение ячейки без проверки шаблона (любой формат)
                        supplier_batch = lines[supplier_line]

                    # Срок годности находится в строке last_name_line + 10
                    expiration_date = None
                    expiration_line = last_name_line + 10
                    if expiration_line < len(lines):
                        # Проверяем, что это дата в формате ДД.ММ.ГГГГ
                        date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', lines[expiration_line])
                        if date_match:
                            expiration_date = date_match.group(1)

                    # Добавляем товар, если есть аналитический лист
                    if analytic_list:
                        products.append({
                            'наименование': product_name,
                            'номенклатурный_номер': nomencl_number,
                            'аналитический_лист': analytic_list,
                            'партия_поставщика': supplier_batch,
                            'срок_годности': expiration_date
                        })

                    # Переходим к следующему товару (last_name_line + 11)
                    i = last_name_line + 11
                else:
                    i += 1

            result['товары'] = products

    return result


@app.route('/')
def index():
    settings = load_settings()
    return render_template('index.html', settings=settings)


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return render_template('error.html', error='Файл не выбран')

    file = request.files['file']

    if file.filename == '':
        return render_template('error.html', error='Файл не выбран')

    if not file.filename.lower().endswith('.pdf'):
        return render_template('error.html', error='Пожалуйста, загрузите файл в формате PDF')

    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    try:
        data = extract_data_from_pdf(filepath)
        os.remove(filepath)

        if not data['дата_приемки'] and not data['название_поставщика'] and not data['товары']:
            return render_template('error.html',
                                   error='Не удалось извлечь данные из PDF-файла. Проверьте формат файла.')

        return render_template('result.html', data=data)

    except Exception as e:
        if os.path.exists(filepath):
            os.remove(filepath)
        return render_template('error.html', error=f'Ошибка при обработке файла: {str(e)}')


@app.route('/generate_labels', methods=['POST'])
def generate_labels():
    try:
        # Получаем данные из формы
        products_json = request.form.get('products')
        order_number = request.form.get('order_number')

        if not products_json or not order_number:
            return render_template('error.html', error='Нет данных для формирования этикеток')

        # Парсим JSON
        try:
            products = json.loads(products_json)
        except json.JSONDecodeError as e:
            return render_template('error.html', error=f'Ошибка парсинга данных: {str(e)}')

        if not products:
            return render_template('error.html', error='Нет товаров для формирования этикеток')

        # Загружаем настройки
        settings = load_settings()
        save_path = settings.get('save_path')

        # Создаем файл с этикетками
        filepath = create_pallet_labels_file(products, order_number, save_path)

        # Открываем в Word
        open_in_word(filepath)

        return render_template('success.html',
                               filepath=filepath,
                               product_count=len(products),
                               order_number=order_number)

    except Exception as e:
        return render_template('error.html', error=f'Ошибка при формировании этикеток: {str(e)}')


@app.route('/settings', methods=['GET', 'POST'])
def settings_page():
    settings = load_settings()
    drives = get_available_drives()

    if request.method == 'POST':
        new_path = request.form.get('save_path')
        if new_path:
            settings['save_path'] = new_path
            save_settings(settings)
            # Создаем папку если её нет
            os.makedirs(new_path, exist_ok=True)
            return redirect(url_for('settings_page'))

    return render_template('settings.html', settings=settings, drives=drives)


@app.route('/exit')
def exit_app():
    return render_template('exit.html')


if __name__ == '__main__':
    app.run(debug=True)
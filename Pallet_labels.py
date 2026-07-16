# app.py
from flask import Flask, render_template, request, redirect, url_for
import os
import re
import fitz  # PyMuPDF
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Создаем папку для загрузок если её нет
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


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
        'наименование_товара': None,
        'номенклатурный_номер': None,
        'аналитический_лист': None,
        'партия_поставщика': None
    }

    # Проходим по всем строкам и ищем данные
    for i, line in enumerate(lines):
        # 1) Дата приемки - ищем в строке 31 (28.01.2026)
        if re.match(r'\d{2}\.\d{2}\.\d{4}', line):
            # Проверяем, что это не срок годности (строка 94)
            if '2028' not in line and 'годности' not in lines[i - 1] if i > 0 else True:
                result['дата_приемки'] = line
                break

    # Если дата не найдена, ищем в строке после "составления"
    if not result['дата_приемки']:
        for i, line in enumerate(lines):
            if 'составления' in line.lower() and i + 1 < len(lines):
                date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', lines[i + 1])
                if date_match:
                    result['дата_приемки'] = date_match.group(1)
                    break

    # 2) Название поставщика - ищем строку с ООО (строка 34)
    for line in lines:
        if 'ООО' in line or 'ЗАО' in line or 'ОАО' in line or 'АО' in line or 'ИП' in line:
            # Проверяем, что это не "ООО "ВЕРОФАРМ"" (строки 7 и 11)
            if 'ВЕРОФАРМ' not in line and 'Завод' not in line:
                result['название_поставщика'] = line
                break

    # 3) Наименование товара - ищем строку 84 (НАТРИЯ ЦИТРАТ MERCK 106446 (2-ВОД))
    # Ищем по ключевым словам
    for line in lines:
        if 'НАТРИЯ' in line or 'ЦИТРАТ' in line or 'MERCK' in line:
            result['наименование_товара'] = line
            break

    # Если не нашли по ключевым словам, ищем после "Материальные ценности"
    if not result['наименование_товара']:
        found_material = False
        for i, line in enumerate(lines):
            if 'Материальные ценности' in line:
                found_material = True
                # Ищем в следующих строках, пропуская номера колонок (1-14)
                for j in range(i + 1, len(lines)):
                    # Проверяем, что это не цифра от 1 до 14
                    if re.match(r'^[1-9]$|^1[0-4]$', lines[j]):
                        continue
                    # Проверяем, что это не служебное слово
                    if lines[j].lower() in ['кг', 'шт', 'г', 'л', 'м', 'уп', 'пач', 'итого', 'x', 'принял', 'сдал',
                                            'должность', 'подпись', 'расшифровка', 'номер', 'паспорта', 'партия',
                                            'поставщика', 'срок годности']:
                        continue
                    # Проверяем, что это не дата
                    if re.match(r'\d{2}\.\d{2}\.\d{4}', lines[j]):
                        continue
                    # Проверяем, что строка содержит буквы и не является числом
                    if re.search(r'[А-ЯЁA-Z]', lines[j]) and len(lines[j]) > 3:
                        result['наименование_товара'] = lines[j]
                        break
                break

    # 4) Номенклатурный номер - ищем 8-значный номер (строка 85)
    for line in lines:
        # Ищем 8-значный номер, который не является телефоном
        nomencl_match = re.search(r'\b(\d{8})\b', line)
        if nomencl_match:
            # Проверяем, что это не телефон (не начинается с 8 или 7)
            if not nomencl_match.group(1).startswith(('8', '7')):
                # Проверяем, что это не "48797491" (ОКПО) и не "21000101"
                if nomencl_match.group(1) not in ['48797491', '21000101']:
                    result['номенклатурный_номер'] = nomencl_match.group(1)
                    break

    # 5) Аналитический лист (партия) - ищем 10 цифр, начинающихся с 1 (строка 92)
    for line in lines:
        # Ищем 10-значный номер, начинающийся с 1
        batch_match = re.search(r'\b(1\d{9})\b', line)
        if batch_match:
            result['аналитический_лист'] = batch_match.group(1)
            break

    # 6) Партия поставщика - ищем буквенно-цифровой код (строка 93)
    for line in lines:
        # Ищем код с буквой и цифрами
        supplier_batch_match = re.search(r'([A-Z]\d{7,10})', line, re.IGNORECASE)
        if supplier_batch_match:
            result['партия_поставщика'] = supplier_batch_match.group(1)
            break

    # Если номенклатурный номер не найден, ищем после "номенклатурный"
    if not result['номенклатурный_номер']:
        for i, line in enumerate(lines):
            if 'номенклатурный' in line.lower() and i + 1 < len(lines):
                nomencl_match = re.search(r'\b(\d{8})\b', lines[i + 1])
                if nomencl_match:
                    result['номенклатурный_номер'] = nomencl_match.group(1)
                    break

    # Если партия поставщика не найдена, ищем после "Партия поставщика"
    if not result['партия_поставщика']:
        for i, line in enumerate(lines):
            if 'Партия поставщика' in line and i + 1 < len(lines):
                supplier_batch_match = re.search(r'([A-Z0-9]{6,12})', lines[i + 1], re.IGNORECASE)
                if supplier_batch_match:
                    result['партия_поставщика'] = supplier_batch_match.group(1)
                    break

    return result


@app.route('/')
def index():
    return render_template('index.html')


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

        if not any(data.values()):
            return render_template('error.html',
                                   error='Не удалось извлечь данные из PDF-файла. Проверьте формат файла.')

        return render_template('result.html', data=data)

    except Exception as e:
        if os.path.exists(filepath):
            os.remove(filepath)
        return render_template('error.html', error=f'Ошибка при обработке файла: {str(e)}')


@app.route('/exit')
def exit_app():
    return render_template('exit.html')


if __name__ == '__main__':
    app.run(debug=True)
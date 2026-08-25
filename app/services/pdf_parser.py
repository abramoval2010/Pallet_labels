# app/services/pdf_parser.py
import re
import fitz


def extract_text_from_pdf(pdf_path):
    """Универсальная функция извлечения текста из PDF"""
    try:
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()
        return text
    except Exception as e:
        print(f"Error opening PDF: {e}")
        import traceback
        traceback.print_exc()
        return None


def extract_data_from_pdf(pdf_path):
    """Извлекает данные из PDF-файла приходного ордера"""
    text = extract_text_from_pdf(pdf_path)
    if text is None:
        return None

    lines = [line.strip() for line in text.split('\n') if line.strip()]

    result = {
        'дата_приемки': None,
        'название_поставщика': None,
        'номер_ордера': None,
        'товары': []
    }

    # 1) Номер приходного ордера
    for line in lines:
        if 'ПРИХОДНЫЙ ОРДЕР №' in line or 'ПРИХОДНЫЙ ОРДЕР N' in line:
            order_match = re.search(r'№?\s*(\d+)', line)
            if order_match:
                result['номер_ордера'] = order_match.group(1)
                break

    # 2) Дата приемки
    for line in lines:
        date_match = re.search(r'\b(\d{2}\.\d{2}\.\d{4})\b', line)
        if date_match:
            date = date_match.group(1)
            if not date.startswith('2028') and 'годности' not in line.lower():
                result['дата_приемки'] = date
                break

    # 3) Название поставщика
    for line in lines:
        supplier_match = re.search(r'(ООО|ЗАО|ОАО|АО|ИП)\s*["«]?([^"«»]+)["»]?', line)
        if supplier_match:
            if 'ВЕРОФАРМ' not in line.upper() and 'ЗАВОД' not in line.upper():
                full_match = re.search(
                    r'(ООО\s*["«]?[^"»]+["»]?|ЗАО\s*["«]?[^"»]+["»]?|ОАО\s*["«]?[^"»]+["»]?|АО\s*["«]?[^"»]+["»]?|ИП\s*["«]?[^"»]+["»]?)',
                    line)
                if full_match:
                    result['название_поставщика'] = full_match.group(1)
                    break

    # 4) ПАРСИНГ ТОВАРОВ
    start_index = -1
    for i, line in enumerate(lines):
        if 'Материальные ценности' in line:
            start_index = i
            break

    end_index = -1
    for i, line in enumerate(lines):
        if line == 'Итого' or 'Итого' in line:
            if i + 1 < len(lines) and (lines[i + 1] == 'X' or lines[i + 1] == 'x'):
                end_index = i + 1
                break
            elif 'X' in line or 'x' in line:
                end_index = i
                break

    if end_index == -1:
        for i, line in enumerate(lines):
            if line == 'X' or line == 'x' or line.startswith('X'):
                if i > 0 and ('Итого' in lines[i - 1] or 'Итого' in lines[i]):
                    end_index = i
                    break

    if start_index == -1 or end_index == -1:
        return result

    table_lines = lines[start_index:end_index + 1]

    marker_index = -1
    for i, line in enumerate(table_lines):
        if line == '14' or re.search(r'\b14\b', line):
            marker_index = i
            break

    if marker_index == -1:
        for i, line in enumerate(table_lines):
            if re.search(r'\b1\s+2\s+3\s+4\s+5\s+6\s+7\s+8\s+9\s+10\s+11\s+12\s+13\s+14\b', line):
                marker_index = i
                break

    if marker_index == -1:
        return result

    products = []
    i = 0

    while i < len(table_lines):
        line = table_lines[i]

        if any(keyword in line.lower() for keyword in
               ['итого', 'x', 'принял', 'сдал', 'должность', 'подпись',
                'материальные ценности', 'измерения', 'количество', 'цена']):
            i += 1
            continue

        nomencl_match = re.search(r'\b(\d{8})\b', line)
        if nomencl_match:
            nomencl_number = nomencl_match.group(1)

            if (not nomencl_number.startswith(('8', '7')) and
                    nomencl_number not in ['48797491', '21000101', '0315003']):

                sap_row = i

                # --- 1. НАЗВАНИЕ ---
                name_parts = []
                for j in range(sap_row - 1, marker_index, -1):
                    prev_line = table_lines[j]
                    if (re.search(r'[А-ЯЁA-Z]', prev_line) and
                            not re.match(r'^\d+$', prev_line) and
                            not any(keyword in prev_line.lower() for keyword in ['итого', 'x', 'принял', 'сдал'])):
                        name_parts.insert(0, prev_line)
                    else:
                        break
                product_name = ' '.join(name_parts).strip() if name_parts else 'Не указано'

                # --- 2. КОЛИЧЕСТВО ---
                quantity = None
                qty_row = sap_row + 4
                if qty_row < len(table_lines):
                    qty_line = table_lines[qty_row]
                    qty_match = re.search(r'(\d{1,3}(?:\s?\d{3})*)', qty_line)
                    if qty_match:
                        qty_str = qty_match.group(1).replace(' ', '')
                        try:
                            quantity = int(qty_str)
                        except:
                            pass

                # --- 3. АНАЛИТИЧЕСКИЙ ЛИСТ ---
                analytic_list = None
                analytic_row = sap_row + 7
                if analytic_row < len(table_lines):
                    analytic_line = table_lines[analytic_row]
                    analytic_match = re.search(r'\b(1\d{9})\b', analytic_line)
                    if analytic_match:
                        analytic_list = analytic_match.group(1)
                    else:
                        alt_match = re.search(r'\b(\d{10})\b', analytic_line)
                        if alt_match:
                            potential = alt_match.group(1)
                            if not potential.startswith('49'):
                                analytic_list = potential

                # --- 4. ПАРТИЯ ПОСТАВЩИКА ---
                supplier_batch = None
                batch_row = sap_row + 8
                if batch_row < len(table_lines):
                    batch_line = table_lines[batch_row]
                    clean_batch = re.sub(r'\d{2}\.\d{2}\.\d{4}', '', batch_line).strip()
                    if clean_batch:
                        supplier_batch = clean_batch
                    else:
                        if not re.match(r'^\d{2}\.\d{2}\.\d{4}$', batch_line):
                            supplier_batch = batch_line

                # --- 5. СРОК ГОДНОСТИ ---
                expiration_date = None
                exp_row = sap_row + 9
                if exp_row < len(table_lines):
                    exp_line = table_lines[exp_row]
                    exp_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', exp_line)
                    if exp_match:
                        exp_date = exp_match.group(1)
                        if exp_date != result.get('дата_приемки'):
                            expiration_date = exp_date

                products.append({
                    'наименование': product_name,
                    'номенклатурный_номер': nomencl_number,
                    'аналитический_лист': analytic_list if analytic_list else 'Не найден',
                    'партия_поставщика': supplier_batch if supplier_batch else 'Не найдена',
                    'срок_годности': expiration_date if expiration_date else 'Не найден',
                    'количество': quantity if quantity else 'Не найдено'
                })

        i += 1

    result['товары'] = products
    return result
# test_parser.py
import fitz
import re


def debug_products(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()

    lines = [line.strip() for line in text.split('\n') if line.strip()]

    print("=" * 80)
    print("ПАРСИНГ ТОВАРОВ С УЧЕТОМ СТРУКТУРЫ")
    print("=" * 80)

    # Находим таблицу
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

    print(f"Начало таблицы: строка {start_index}")
    print(f"Конец таблицы: строка {end_index}")
    print("=" * 80)

    if start_index != -1 and end_index != -1:
        # Находим строку с "14"
        data_start = -1
        for i in range(start_index, end_index):
            if lines[i] == '14':
                data_start = i + 1
                print(f"Найдена строка '14' на позиции {i}")
                print(f"Начало данных: строка {data_start}")
                break

        print("\n" + "=" * 80)
        print("СТРОКИ ПОСЛЕ '14' ДО 'ИТОГО':")
        print("-" * 80)

        if data_start != -1:
            for i in range(data_start, end_index):
                print(f"{i:3d}: {lines[i]}")

        print("\n" + "=" * 80)
        print("ОПРЕДЕЛЕНИЕ ТОВАРОВ:")
        print("-" * 80)

        if data_start != -1:
            # Список заголовков
            header_keywords = [
                'наименование', 'сорт', 'размер', 'марка', 'код',
                'единица измерения', 'количество', 'цена', 'сумма',
                'без учета', 'всего с учетом', 'номер паспорта',
                'партия', 'поставщика', 'партия поставщика',
                'срок годности', 'паспорта'
            ]

            skip_words = ['кг', 'шт', 'г', 'л', 'м', 'уп', 'пач', 'итого',
                          'x', 'принял', 'сдал', 'доложность', 'подпись',
                          'расшифровка', 'номер']

            i = data_start
            product_count = 0

            while i < end_index:
                line = lines[i]

                # Пропускаем даты
                if re.match(r'\d{2}\.\d{2}\.\d{4}', line):
                    print(f"  Строка {i}: ДАТА (срок годности) - {line}")
                    i += 1
                    continue

                # Пропускаем числа
                if re.match(r'^[\d\s,]+$', line) and len(line) > 3:
                    print(f"  Строка {i}: ЧИСЛА - {line}")
                    i += 1
                    continue

                # Пропускаем "X"
                if line == 'X':
                    print(f"  Строка {i}: X")
                    i += 1
                    continue

                # Проверяем на заголовок
                is_header = False
                for keyword in header_keywords:
                    if keyword in line.lower():
                        is_header = True
                        break

                if is_header:
                    print(f"  Строка {i}: ЗАГОЛОВОК - {line}")
                    i += 1
                    continue

                # Проверяем на служебное слово
                if line.lower() in skip_words:
                    print(f"  Строка {i}: СЛУЖЕБНОЕ - {line}")
                    i += 1
                    continue

                # Если строка содержит буквы - это название товара
                if re.search(r'[А-ЯЁA-Z]', line):
                    product_count += 1
                    print(f"\nТОВАР #{product_count}:")
                    print(f"  Название (строка {i}): {line}")

                    # Ищем номенклатурный номер
                    for j in range(i + 1, min(i + 5, len(lines))):
                        nomencl_match = re.search(r'\b(\d{8})\b', lines[j])
                        if nomencl_match:
                            if not nomencl_match.group(1).startswith(('8', '7')):
                                if nomencl_match.group(1) not in ['48797491', '21000101']:
                                    print(f"  Номенклатурный номер (строка {j}): {nomencl_match.group(1)}")
                                    break

                    # Ищем аналитический лист
                    for j in range(i + 1, min(i + 10, len(lines))):
                        batch_match = re.search(r'\b(1\d{9})\b', lines[j])
                        if batch_match:
                            print(f"  Аналитический лист (строка {j}): {batch_match.group(1)}")
                            break

                    # Ищем партию поставщика
                    for j in range(i + 1, min(i + 10, len(lines))):
                        supplier_match = re.search(r'([A-Z]\d{7,10})', lines[j], re.IGNORECASE)
                        if supplier_match:
                            print(f"  Партия поставщика (строка {j}): {supplier_match.group(1)}")
                            break

                    # Ищем срок годности
                    for j in range(i + 1, min(i + 12, len(lines))):
                        date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', lines[j])
                        if date_match:
                            print(f"  Срок годности (строка {j}): {date_match.group(1)}")
                            break

                    i += 1
                else:
                    i += 1

            print("\n" + "=" * 80)
            print(f"Всего найдено товаров: {product_count}")


if __name__ == "__main__":
    debug_products("Приходный ордер.pdf")
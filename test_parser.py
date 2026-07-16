# test_parser.py
import fitz
import re


def debug_pdf_structure(pdf_path):
    """
    Отладочная функция для просмотра структуры PDF
    """
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()

    # Разбиваем на строки
    lines = [line.strip() for line in text.split('\n') if line.strip()]

    print("=" * 80)
    print("СТРУКТУРА PDF-ФАЙЛА (все строки с номерами)")
    print("=" * 80)

    # Выводим все строки с индексами
    for i, line in enumerate(lines):
        print(f"{i:3d}: {line}")

    print("=" * 80)
    print("\nПОИСК КЛЮЧЕВЫХ ДАННЫХ:")
    print("=" * 80)

    # Проверяем каждую строку
    for i, line in enumerate(lines):
        # Проверяем дату
        if 'Дата составления' in line:
            print(f"\n✅ Найдено 'Дата составления' в строке {i}")
            print(f"   Строка: {line}")
            if i + 1 < len(lines):
                print(f"   Следующая строка ({i + 1}): {lines[i + 1]}")
                date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', lines[i + 1])
                if date_match:
                    print(f"   📅 Найдена дата: {date_match.group(1)}")
                else:
                    print(f"   ❌ Дата не найдена в следующей строке")

        # Проверяем поставщика
        if 'поставщик' in line.lower():
            print(f"\n✅ Найдено 'Поставщик' в строке {i}")
            print(f"   Строка: {line}")
            if i + 1 < len(lines):
                print(f"   Следующая строка ({i + 1}): {lines[i + 1]}")
                supplier_match = re.search(r'(ООО|ЗАО|ОАО|АО|ИП)\s*[«"]?([^«»"",\d]+)[»"]?', lines[i + 1],
                                           re.IGNORECASE)
                if supplier_match:
                    print(f"   🏢 Найден поставщик: {supplier_match.group(0)}")
                else:
                    print(f"   ❌ Поставщик не найден в следующей строке")

        # Проверяем номенклатурный номер
        if 'номенклатурный номер' in line.lower():
            print(f"\n✅ Найдено 'номенклатурный номер' в строке {i}")
            print(f"   Строка: {line}")
            found = False
            for j in range(i + 1, min(i + 5, len(lines))):
                print(f"   Строка {j}: {lines[j]}")
                nomencl_match = re.search(r'\b(\d{8})\b', lines[j])
                if nomencl_match:
                    print(f"   🔢 Найден номер: {nomencl_match.group(1)}")
                    found = True
                    break
            if not found:
                print(f"   ❌ Номенклатурный номер не найден")

        # Проверяем партию
        if 'партия' in line.lower() and 'поставщика' not in line.lower():
            print(f"\n✅ Найдено 'партия' в строке {i}")
            print(f"   Строка: {line}")
            if i + 1 < len(lines):
                print(f"   Следующая строка ({i + 1}): {lines[i + 1]}")
                batch_match = re.search(r'\b(1\d{9})\b', lines[i + 1])
                if batch_match:
                    print(f"   📋 Найден аналитический лист: {batch_match.group(1)}")
                else:
                    print(f"   ❌ Аналитический лист не найден в следующей строке")

        # Проверяем партию поставщика
        if 'партия поставщика' in line.lower():
            print(f"\n✅ Найдено 'Партия поставщика' в строке {i}")
            print(f"   Строка: {line}")
            if i + 1 < len(lines):
                print(f"   Следующая строка ({i + 1}): {lines[i + 1]}")
                supplier_batch_match = re.search(r'([A-Z0-9]{6,12})', lines[i + 1], re.IGNORECASE)
                if supplier_batch_match:
                    print(f"   🔖 Найдена партия поставщика: {supplier_batch_match.group(1)}")
                else:
                    print(f"   ❌ Партия поставщика не найдена в следующей строке")


if __name__ == "__main__":
    debug_pdf_structure("Приходный ордер.pdf")
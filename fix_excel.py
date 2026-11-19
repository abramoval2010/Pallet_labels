import pandas as pd

# Замените 'ваш_файл.xlsx' на реальное имя файла, который вы пытаетесь загрузить
source_file = 'Серии ГП на складе.xlsx'
target_file = 'Серии_ГП_ИСПРАВЛЕННЫЙ.xlsx'

try:
    # Читаем файл, игнорируя возможные проблемы с форматом
    df = pd.read_excel(source_file, engine='openpyxl')

    # Записываем обратно в новый файл, который будет гарантированно в формате .xlsx
    df.to_excel(target_file, index=False, engine='openpyxl')

    print(f"✅ Успех! Файл пересоздан: {target_file}")
    print("Попробуйте загрузить этот новый файл в приложение.")
except Exception as e:
    print(f"❌ Не удалось пересоздать файл: {e}")
    print("Вероятно, файл имеет серьёзные повреждения или защиту.")
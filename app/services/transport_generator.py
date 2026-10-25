# app/services/transport_generator.py
import os
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def remove_table_borders(table):
    """Убирает рамки у таблицы, делая её невидимой"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'nil')
        borders.append(border)
    tblPr.append(borders)


def generate_transport_docx(data, save_dir):
    doc = Document()

    # --- 1. Настройка страницы (А4) и полей ---
    section = doc.sections[0]
    section.page_height = Mm(297)
    section.page_width = Mm(210)
    section.top_margin = Mm(15)
    section.bottom_margin = Mm(15)
    section.left_margin = Mm(20)
    section.right_margin = Mm(15)

    # --- 2. Настройка базового шрифта ---
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(10)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.line_spacing = 1.0

    # --- 3. Шапка документа ---
    p_header1 = doc.add_paragraph('от ООО «ВЕРОФАРМ» (ИНН 5032048702)')
    p_header1.alignment = WD_ALIGN_PARAGRAPH.LEFT

    p_header2 = doc.add_paragraph('в ООО «Ритейл транспорт эдженси» (ИНН 7713636401)')
    p_header2.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # --- 4. Заголовок (в 2 строки) ---
    p_title_1 = doc.add_paragraph()
    p_title_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title_1.paragraph_format.space_before = Pt(15)
    p_title_1.paragraph_format.space_after = Pt(0)
    run1 = p_title_1.add_run(f"Поручение № {data['order_number']}")
    run1.bold = True
    run1.font.size = Pt(12)

    p_title_2 = doc.add_paragraph()
    p_title_2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title_2.paragraph_format.space_before = Pt(0)
    p_title_2.paragraph_format.space_after = Pt(15)
    run2 = p_title_2.add_run("на перевозку грузов автомобильным транспортом")
    run2.bold = True
    run2.font.size = Pt(12)

    # --- 5. Город и дата в одной строке ---
    header_table = doc.add_table(rows=1, cols=2)
    header_table.autofit = False
    header_table.allow_autofit = False
    header_table.columns[0].width = Cm(8.75)
    header_table.columns[1].width = Cm(8.75)

    cell_city = header_table.cell(0, 0)
    cell_city.text = "г. Москва"
    cell_city.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

    cell_date = header_table.cell(0, 1)
    try:
        date_obj = datetime.strptime(data['order_date'], '%Y-%m-%d')
        months = {1: 'января', 2: 'февраля', 3: 'марта', 4: 'апреля', 5: 'мая', 6: 'июня',
                  7: 'июля', 8: 'августа', 9: 'сентября', 10: 'октября', 11: 'ноября', 12: 'декабря'}
        date_str = f"«{date_obj.day}» {months[date_obj.month]} {date_obj.year} г."
    except:
        date_str = f"«{data['order_date']} г."
    cell_date.text = date_str
    cell_date.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    remove_table_borders(header_table)

    # Отступ после шапки с датой
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_after = Pt(10)

    # --- 6. Таблица с данными ---
    table = doc.add_table(rows=17, cols=2)
    table.style = 'Table Grid'
    table.autofit = False
    table.allow_autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for row in table.rows:
        row.cells[0].width = Cm(6.5)
        row.cells[1].width = Cm(11)

    try:
        dd = datetime.strptime(data['delivery_date'], '%Y-%m-%d').strftime('%d.%m.%Y')
    except:
        dd = data['delivery_date']

    try:
        sd = datetime.strptime(data.get('supply_date', data['loading_date']), '%Y-%m-%d').strftime('%d.%m.%Y')
    except:
        sd = data.get('supply_date', data['loading_date'])

    s_time = data.get('supply_time', data.get('loading_time', ''))

    cost_val = data.get('cost', 49999990)
    cost_formatted = f"{cost_val:,}".replace(',', ' ')

    rows_data = [
        ("Дата и время подачи ТС:", f"{sd} г. в {s_time} ч."),
        ("Наименование отправителя:", data.get('sender', 'ООО «ВЕРОФАРМ»')),
        ("Место подачи (адрес):", data.get('loading_place', '')),
        ("Контактное лицо на месте загрузки:", data.get('loading_contact', '')),
        ("Дата и время доставки получателю:", f"{dd} г. {data.get('delivery_time', '')} ч."),
        ("Наименование получателя:", data.get('receiver', 'ООО «ВЕРОФАРМ»')),
        ("Адреса доставки:", data.get('delivery_address', '')),
        ("Контактное лицо на разгрузке", data.get('unloading_contact', '')),
        ("Наименование груза:", data.get('cargo_name', '')),
        ("Количество паллет/коробов / м³ / вес", f"{data.get('quantity', 33)} паллета"),
        ("Тип требуемого ТС: тип кузова, грузоподъемность, кол-во паллет, м³:", data.get('vehicle_type', '')),
        ("Температурный режим", data.get('temperature_mode', '')),
        ("Дополнительные комментарии к поездке:", f"Стоимость груза {cost_formatted} руб."),
        ("ФИО водителя", data.get('driver_name', '')),
        ("Данные водителя", data.get('driver_data', '')),
        ("Контактный телефон", data.get('driver_phone', '')),
        ("№ ТС", data.get('vehicle_number', '')),
    ]

    for i, (key, val) in enumerate(rows_data):
        row = table.rows[i]
        cell_key = row.cells[0]
        cell_val = row.cells[1]

        cell_key.text = key
        cell_val.text = val

        cell_key.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        cell_val.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        for paragraph in cell_key.paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(9)
            paragraph.paragraph_format.space_after = Pt(1)
            paragraph.paragraph_format.space_before = Pt(1)

        for paragraph in cell_val.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(9)
            paragraph.paragraph_format.space_after = Pt(1)
            paragraph.paragraph_format.space_before = Pt(1)

    # --- 7. Блок подписей и м.п. ---
    p_spacer2 = doc.add_paragraph()
    p_spacer2.paragraph_format.space_after = Pt(10)

    sig_trans_table = doc.add_table(rows=4, cols=3)
    sig_trans_table.autofit = False
    sig_trans_table.allow_autofit = False

    # ИСПРАВЛЕНИЕ: Перераспределили ширину колонок, чтобы третья стала шире
    sig_trans_table.columns[0].width = Cm(3.2)
    sig_trans_table.columns[1].width = Cm(3.3)
    sig_trans_table.columns[2].width = Cm(11.0)

    # Строка 0: Передал
    c_lbl = sig_trans_table.cell(0, 0)
    c_lbl.text = "Заявку передал:"
    c_lbl.paragraphs[0].runs[0].bold = True

    c_dots = sig_trans_table.cell(0, 1)
    c_dots.text = "................ ,"

    c_val = sig_trans_table.cell(0, 2)
    # ИСПРАВЛЕНИЕ: Уменьшили количество подчеркиваний с 13 до 10
    c_val.text = f"__________ {data['author']}"

    # Строка 1: м.п.
    sig_trans_table.cell(1, 0).text = ""
    c_mp = sig_trans_table.cell(1, 1)
    c_mp.text = "м.п."
    c_mp.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    sig_trans_table.cell(1, 2).text = ""

    # Строка 2: Принял
    c_lbl2 = sig_trans_table.cell(2, 0)
    c_lbl2.text = "Заявку принял:"
    c_lbl2.paragraphs[0].runs[0].bold = True

    c_dots2 = sig_trans_table.cell(2, 1)
    c_dots2.text = "................ ,"

    c_val2 = sig_trans_table.cell(2, 2)
    # ИСПРАВЛЕНИЕ: Уменьшили количество подчеркиваний с 13 до 10
    c_val2.text = f"__________ Голубков Г.И."

    # Строка 3: м.п.
    sig_trans_table.cell(3, 0).text = ""
    c_mp2 = sig_trans_table.cell(3, 1)
    c_mp2.text = "м.п."
    c_mp2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    sig_trans_table.cell(3, 2).text = ""

    remove_table_borders(sig_trans_table)

    # --- 8. ПОДПИСИ СТОРОН ---
    p_sign = doc.add_paragraph()
    p_sign.add_run('ПОДПИСИ СТОРОН').bold = True
    p_sign.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sign.paragraph_format.space_before = Pt(15)
    p_sign.paragraph_format.space_after = Pt(10)

    sig_table = doc.add_table(rows=7, cols=2)
    sig_table.autofit = False
    sig_table.allow_autofit = False
    sig_table.columns[0].width = Cm(8.75)
    sig_table.columns[1].width = Cm(8.75)

    # Строка 0: КЛИЕНТ и ИСПОЛНИТЕЛЬ
    sig_table.cell(0, 0).text = "КЛИЕНТ"
    sig_table.cell(0, 0).paragraphs[0].runs[0].bold = True
    sig_table.cell(0, 1).text = "ИСПОЛНИТЕЛЬ"
    sig_table.cell(0, 1).paragraphs[0].runs[0].bold = True

    # Строка 1: Должности
    sig_table.cell(1, 0).text = "Генеральный директор ООО «ВЕРОФАРМ»"
    sig_table.cell(1, 1).text = "Генеральный директор ООО «РТЭ»"

    # Строки 2, 3, 4 остаются пустыми (создают отступ в 3 строки)

    # Строка 5: Роспись
    sig_table.cell(5, 0).text = "__________________ /С.В. Москалев/"
    sig_table.cell(5, 1).text = "__________________ /Х.Р. Ибрагимов/"

    # Строка 6: м.п.
    sig_table.cell(6, 0).text = "м.п."
    sig_table.cell(6, 1).text = "м.п."

    remove_table_borders(sig_table)

    # --- 9. Сохранение ---
    try:
        date_obj = datetime.strptime(data.get('loading_date', ''), '%Y-%m-%d')
        formatted_date = date_obj.strftime('%d.%m.%Y')
    except:
        formatted_date = data.get('loading_date', '')

    route_name = data.get('route_name', 'РТЭ Вольгинский - Доминанта')
    filename = f"{route_name} №{data['order_number']} от {formatted_date}.docx"
    filepath = os.path.join(save_dir, filename)
    doc.save(filepath)

    return filepath, filename
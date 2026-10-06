import os
try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
except ImportError:
    print("Please install python-docx using: pip install python-docx")
    exit(1)

def set_font(run, size=14, bold=False):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    run.font.bold = bold

def add_centered_paragraph(doc, text, size=14, bold=False):
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run(text)
    set_font(run, size, bold)
    return p

def create_report(filepath):
    doc = Document()
    
    # Title Page
    add_centered_paragraph(doc, "Федеральное государственное бюджетное образовательное учреждение")
    add_centered_paragraph(doc, "высшего образования")
    add_centered_paragraph(doc, "«Уфимский университет науки и технологий»")
    doc.add_paragraph()
    add_centered_paragraph(doc, "Кафедра АСУ")
    
    for _ in range(5):
        doc.add_paragraph()
        
    add_centered_paragraph(doc, "Отчет по лабораторной работе № 1")
    add_centered_paragraph(doc, "по дисциплине: СТАУ")
    add_centered_paragraph(doc, "на тему: «Расчет стоимости кабельной сети для корпусов К4 и К5»")
    
    for _ in range(5):
        doc.add_paragraph()
        
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    run = p.add_run("Выполнил:\nстудент гр. ЭАС-414С\nГарифуллин К.Р.\nПроверил:\nАнтонов В.В.")
    set_font(run, 14)
    
    for _ in range(5):
        doc.add_paragraph()
        
    add_centered_paragraph(doc, "Уфа 2026 г.")
    
    doc.add_page_break()
    
    # Content
    p = doc.add_paragraph()
    run = p.add_run("Цель работы")
    set_font(run, 14, bold=True)
    
    p = doc.add_paragraph()
    run = p.add_run("Освоить принципы расчета кабельной инфраструктуры зданий, рассчитать суммарную длину кабельных трасс и количество коммутационного оборудования для корпусов К4 и К5, а также составить итоговую смету.")
    set_font(run, 14)
    
    p = doc.add_paragraph()
    run = p.add_run("Ход работы")
    set_font(run, 14, bold=True)
    
    p = doc.add_paragraph()
    run = p.add_run("1. Используемое оборудование и материалы")
    set_font(run, 14, bold=True)
    
    table_equip = doc.add_table(rows=1, cols=3)
    table_equip.style = 'Table Grid'
    hdr_cells = table_equip.rows[0].cells
    hdr_cells[0].text = 'Наименование'
    hdr_cells[1].text = 'Количество'
    hdr_cells[2].text = 'Цена'
    
    equipment = [
        ('Маршрутизатор MikroTik CCR2116-12G-4S+', '8 шт.', '136 550 руб./шт.'),
        ('Коммутационный шкаф SYSMATRIX, 22U 600х800х1100', '8 шт.', '56 728 руб./шт.'),
        ('Патч-панель ExeGate EPP3-19-24-8P8C-C5e-SH-110D', '8 шт.', '2 200 руб./шт.'),
        ('Кабель (витая пара)', '4898 м', '40 руб./м'),
        ('Кабель-канал', '4898 м', '44 руб./м'),
        ('Разъем RJ45', '298 шт.', '16 руб./шт.')
    ]
    
    for item, qty, price in equipment:
        row_cells = table_equip.add_row().cells
        row_cells[0].text = item
        row_cells[1].text = qty
        row_cells[2].text = price

    for row in table_equip.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12)

    doc.add_paragraph()
        
    p = doc.add_paragraph()
    run = p.add_run("2. Расчет длин кабеля для Корпуса 4 (К4)")
    set_font(run, 14, bold=True)
    
    table_k4 = doc.add_table(rows=1, cols=3)
    table_k4.style = 'Table Grid'
    hdr_cells = table_k4.rows[0].cells
    hdr_cells[0].text = 'Этаж'
    hdr_cells[1].text = 'Кабель от маршрутизатора до кабинетов'
    hdr_cells[2].text = 'Межмаршрутизаторный кабель'
    
    k4_data = [
        ('1 этаж', '726 м', '0 м'),
        ('2 этаж', '1097 м', '0 м'),
        ('3 этаж', '970 м', '48 м'),
        ('4 этаж', '973 м', '0 м'),
        ('ИТОГО', '3766 м', '48 м')
    ]
    
    for floor, cab_len, router_len in k4_data:
        row_cells = table_k4.add_row().cells
        row_cells[0].text = floor
        row_cells[1].text = cab_len
        row_cells[2].text = router_len
        
    for row in table_k4.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12)
                    
    p = doc.add_paragraph()
    run = p.add_run("Суммарная длина кабеля для корпуса К4: 3766 + 48 = 3814 м")
    set_font(run, 14)
    
    p = doc.add_paragraph()
    run = p.add_run("3. Расчет длин кабеля для Корпуса 5 (К5)")
    set_font(run, 14, bold=True)
    
    table_k5 = doc.add_table(rows=1, cols=3)
    table_k5.style = 'Table Grid'
    hdr_cells = table_k5.rows[0].cells
    hdr_cells[0].text = 'Этаж'
    hdr_cells[1].text = 'Кабель от маршрутизатора до кабинетов'
    hdr_cells[2].text = 'Межмаршрутизаторный кабель'
    
    k5_data = [
        ('1 этаж', '86 м', '0 м'),
        ('2 этаж', '270 м', '0 м'),
        ('3 этаж', '228 м', '20 м'),
        ('4 этаж', '282 м', '0 м'),
        ('ИТОГО', '866 м', '20 м')
    ]
    
    for floor, cab_len, router_len in k5_data:
        row_cells = table_k5.add_row().cells
        row_cells[0].text = floor
        row_cells[1].text = cab_len
        row_cells[2].text = router_len
        
    for row in table_k5.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12)
                    
    p = doc.add_paragraph()
    run = p.add_run("Суммарная длина кабеля для корпуса К5: 866 + 20 = 886 м")
    set_font(run, 14)
    
    p = doc.add_paragraph()
    run = p.add_run("4. Дополнительные кабельные трассы и разъемы")
    set_font(run, 14, bold=True)
    
    p = doc.add_paragraph()
    run = p.add_run("- Кабель до Корпуса 2 (через К6): 150 м\n- Сумма кабелей между маршрутизаторами внутри корпусов (вертикальные): 16 м (К4) + 16 м (К5) + 16 м (К6, резерв)\n- Суммарная длина всех кабелей: 3814 + 886 + 150 + 16 + 16 + 16 = 4898 м\n\nРасчет количества разъемов (RJ45):\n- Корпус 4: 22+29+25+27 = 103 кабеля × 2 = 206 разъемов\n- Корпус 5: 2+15+13+16 = 46 кабелей × 2 = 92 разъема\n- Итого разъемов: 206 + 92 = 298 шт.")
    set_font(run, 14)
    
    p = doc.add_paragraph()
    run = p.add_run("6. Расчет стоимости монтажных работ")
    set_font(run, 14, bold=True)
    
    p = doc.add_paragraph()
    run = p.add_run("Срок выполнения работ: 3 недели.\nСостав бригады и оплата труда (за весь объект):\n- Монтажник (3 чел.): 60 000 руб. Стоимость: 60 000 × 3 = 180 000 руб.\n- Специалист-сетевик (2 чел.): 100 000 руб. Стоимость: 100 000 × 2 = 200 000 руб.\n- Бригадир (1 чел.): 150 000 руб. Стоимость: 150 000 × 1 = 150 000 руб.\nИтого фонд оплаты труда (ФОТ): 180 000 + 200 000 + 150 000 = 530 000 руб.")
    set_font(run, 14)

    p = doc.add_paragraph()
    run = p.add_run("7. Итоговая смета")
    set_font(run, 14, bold=True)
    
    table_total = doc.add_table(rows=1, cols=2)
    table_total.style = 'Table Grid'
    hdr_cells = table_total.rows[0].cells
    hdr_cells[0].text = 'Статья расходов'
    hdr_cells[1].text = 'Стоимость, руб.'
    
    totals = [
        ('Кабель (4898 м × 40 руб.)', '195 920'),
        ('Кабель-канал (4898 м × 44 руб.)', '215 512'),
        ('Разъемы (298 шт. × 16 руб.)', '4 768'),
        ('Маршрутизаторы (8 шт. × 136 550 руб.)', '1 092 400'),
        ('Коммутационные шкафы (8 шт. × 56 728 руб.)', '453 824'),
        ('Патч-панели (8 шт. × 2 200 руб.)', '17 600'),
        ('Фонд оплаты труда (монтаж)', '530 000')
    ]
    
    for item, cost in totals:
        row_cells = table_total.add_row().cells
        row_cells[0].text = item
        row_cells[1].text = cost
        
    row_cells = table_total.add_row().cells
    row_cells[0].text = 'ОБЩАЯ СТОИМОСТЬ'
    row_cells[1].text = '2 509 924'
    
    for row in table_total.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12)
    
    # Вывод
    p = doc.add_paragraph()
    run = p.add_run("Вывод: В ходе лабораторной работы была спроектирована кабельная инфраструктура для корпусов К4 и К5. Общая длина кабеля и кабель-канала составила 4898 м, количество коннекторов RJ45 - 298 шт. Общая стоимость проекта с учетом активного оборудования, коммутационных шкафов и оплаты труда монтажной бригады составила 2 509 924 рубля.")
    set_font(run, 14)
    
    # Приложение со схемами
    doc.add_page_break()
    add_centered_paragraph(doc, "Приложение. Схемы прокладки кабельных трасс", size=14, bold=True)
    
    images = [
        ("К4-1 этаж.jpg", "План 1 этажа корпуса К4"),
        ("К4-2 этаж.jpg", "План 2 этажа корпуса К4"),
        ("К4-3 этаж.jpg", "План 3 этажа корпуса К4"),
        ("К4-4 этаж.jpg", "План 4 этажа корпуса К4"),
        ("К5-1 этаж.jpg", "План 1 этажа корпуса К5"),
        ("К5-2 этаж.jpg", "План 2 этажа корпуса К5"),
        ("К5-3 этаж.jpg", "План 3 этажа корпуса К5"),
        ("К5-4 этаж.jpg", "План 4 этажа корпуса К5")
    ]
    
    base_dir = r"c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\СТАУ\ЛР1"
    for img_file, caption in images:
        img_path = os.path.join(base_dir, img_file)
        if os.path.exists(img_path):
            p = doc.add_paragraph()
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            try:
                # Вставляем изображение (ширина до 6 дюймов для A4)
                r = p.add_run()
                r.add_picture(img_path, width=Inches(6.0))
                # Подпись к рисунку
                add_centered_paragraph(doc, caption, size=12, bold=False)
                doc.add_paragraph()
            except Exception as e:
                print(f"Не удалось добавить изображение {img_file}: {e}")
            
    doc.save(filepath)
    print(f"Отчет успешно сохранен: {filepath}")

if __name__ == '__main__':
    report_path = os.path.join(r"c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\СТАУ\ЛР1", "ГарифуллинКР_ЛР1_СТАУ_v4.docx")
    create_report(report_path)

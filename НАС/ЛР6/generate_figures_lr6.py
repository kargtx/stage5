import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

DIR_LR6 = r"C:\Users\danch\OneDrive\Рабочий стол\stage5\НАС\ЛР6"

def make_figures():
    plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
    plt.rcParams['axes.edgecolor'] = 'black'
    plt.rcParams['axes.linewidth'] = 1.0

    # -------------------------------------------------------------
    # Figure 1: Data Architecture & Relational Schema (ER Model)
    # -------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(11, 7.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    def draw_entity(x, y, w, h, title, fields, pk_indices=[0], fk_indices=[]):
        # Header
        hdr_h = 4.5
        body_h = h - hdr_h
        # Shadow
        shadow = patches.Rectangle((x - w/2 + 0.8, y - h/2 - 0.8), w, h, facecolor='#D0D0D0', edgecolor='none')
        ax.add_patch(shadow)
        # Body
        body = patches.Rectangle((x - w/2, y - h/2), w, body_h, facecolor='#FFFFFF', edgecolor='#333333', linewidth=1.2)
        ax.add_patch(body)
        # Header
        hdr = patches.Rectangle((x - w/2, y + h/2 - hdr_h), w, hdr_h, facecolor='#1F4E79', edgecolor='#333333', linewidth=1.2)
        ax.add_patch(hdr)
        ax.text(x, y + h/2 - hdr_h/2, title, ha='center', va='center', fontsize=9.5, fontweight='bold', color='#FFFFFF')

        # Fields
        line_h = body_h / (len(fields) + 0.5)
        curr_y = y + h/2 - hdr_h - line_h * 0.8
        for i, fld in enumerate(fields):
            is_pk = i in pk_indices
            is_fk = i in fk_indices
            prefix = "[PK] " if is_pk else ("[FK] " if is_fk else "       ")
            col = '#A00000' if is_pk else ('#004080' if is_fk else '#222222')
            weight = 'bold' if is_pk else 'normal'
            ax.text(x - w/2 + 1.2, curr_y, prefix + fld, ha='left', va='center', fontsize=7.2, color=col, fontweight=weight)
            curr_y -= line_h

    # Entities
    draw_entity(20, 80, 28, 26, "Склады", [
        "Номер склада (Long)",
        "Телефон (Text)",
        "Адрес (Text)",
        "Заведующий (Text)"
    ], pk_indices=[0])

    draw_entity(80, 80, 30, 28, "Фирмы", [
        "Код фирмы (Long)",
        "Название фирмы (Text)",
        "Адрес (Text)",
        "Телефон (Text)"
    ], pk_indices=[0])

    draw_entity(20, 35, 30, 42, "Товары", [
        "Код товара (AutoNumber)",
        "Наименование (Text)",
        "Марка (Text)",
        "Цена (Currency)",
        "Номер склада (Long)",
        "Количество (Long)",
        "Описание (Memo)",
        "Изображение (OLE)"
    ], pk_indices=[0], fk_indices=[4])

    draw_entity(80, 35, 30, 42, "Контактные лица", [
        "Код сотрудника (Long)",
        "Фамилия (Text)",
        "Имя (Text)",
        "Дата рождения (DateTime)",
        "Домашний адрес (Text)",
        "Домашний телефон (Text)",
        "Должность (Text)",
        "Код фирмы (Long)"
    ], pk_indices=[0], fk_indices=[7])

    draw_entity(50, 15, 32, 28, "Продажи", [
        "Дата продажи (DateTime)",
        "Код фирмы (Long)",
        "Код товара (Long)",
        "Количество (Long)",
        "Скидка (Single)"
    ], pk_indices=[0, 1, 2], fk_indices=[1, 2])

    def draw_rel(x1, y1, x2, y2, l1="1", l2="∞"):
        ax.plot([x1, x2], [y1, y2], color='#1F4E79', linewidth=1.6)
        ax.text(x1 + (x2-x1)*0.08, y1 + (y2-y1)*0.08 + 1.5, l1, fontsize=9, fontweight='bold', color='#1F4E79')
        ax.text(x2 - (x2-x1)*0.08, y2 - (y2-y1)*0.08 + 1.5, l2, fontsize=10, fontweight='bold', color='#1F4E79')

    # Relations
    # Склады -> Товары
    draw_rel(20, 67, 20, 56, "1", "∞")
    # Фирмы -> Контакты
    draw_rel(80, 66, 80, 56, "1", "∞")
    # Фирмы -> Продажи
    draw_rel(70, 66, 62, 29, "1", "∞")
    # Товары -> Продажи
    draw_rel(30, 35, 38, 25, "1", "∞")

    # Legend / Annotation
    leg = patches.Rectangle((3, 2), 94, 6, facecolor='#F0F4F8', edgecolor='#B0C4DE', linewidth=1)
    ax.add_patch(leg)
    ax.text(50, 5, "Целостность данных: Включена | Каскадное обновление: Да | Каскадное удаление: Да | Тип связи: 1:M (Один-ко-многим)",
            ha='center', va='center', fontsize=7.8, fontweight='bold', color='#1F4E79')

    p1 = os.path.join(DIR_LR6, "fig_er_model.png")
    fig.savefig(p1, bbox_inches='tight', dpi=300)
    plt.close(fig)
    print("Created:", p1)

    # -------------------------------------------------------------
    # Figure 2: Visual Table Datasheet Simulation (Таблицы БД)
    # -------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(11, 7.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    def draw_datasheet(x, y, w, h, title, cols, rows, widths=None):
        if not widths:
            widths = [w / len(cols)] * len(cols)
        # Window box
        win = patches.Rectangle((x - w/2, y - h/2), w, h, facecolor='#FFFFFF', edgecolor='#808080', linewidth=1.2)
        ax.add_patch(win)
        # Title bar
        t_bar = patches.Rectangle((x - w/2, y + h/2 - 3.5), w, 3.5, facecolor='#D9E1F2', edgecolor='#808080', linewidth=1)
        ax.add_patch(t_bar)
        ax.text(x - w/2 + 1.5, y + h/2 - 1.8, f"Таблица: {title}", ha='left', va='center', fontsize=8, fontweight='bold', color='#1F4E79')

        # Header row
        hdr_y = y + h/2 - 7.0
        curr_x = x - w/2
        for col_name, col_w in zip(cols, widths):
            hdr_cell = patches.Rectangle((curr_x, hdr_y), col_w, 3.5, facecolor='#E7E6E6', edgecolor='#A6A6A6', linewidth=0.8)
            ax.add_patch(hdr_cell)
            ax.text(curr_x + col_w/2, hdr_y + 1.75, col_name, ha='center', va='center', fontsize=6.8, fontweight='bold', color='#333333')
            curr_x += col_w

        # Rows
        row_y = hdr_y
        for r_idx, row in enumerate(rows):
            row_y -= 3.0
            if row_y < y - h/2 + 0.5:
                break
            curr_x = x - w/2
            bg = '#FFFFFF' if r_idx % 2 == 0 else '#F9F9F9'
            for val, col_w in zip(row, widths):
                c_rect = patches.Rectangle((curr_x, row_y), col_w, 3.0, facecolor=bg, edgecolor='#D9D9D9', linewidth=0.6)
                ax.add_patch(c_rect)
                ax.text(curr_x + col_w/2, row_y + 1.5, str(val), ha='center', va='center', fontsize=6.2, color='#222222')
                curr_x += col_w

    # Table Склады
    draw_datasheet(25, 78, 46, 21, "Склады", 
                   ["№ склада", "Телефон", "Адрес", "Заведующий"],
                   [
                       ["10", "8(495)953-0189", "ул. Свободы,37", "Иванов П.П."],
                       ["20", "8(499)712-2222", "ул. Мелитопольская, 35", "Гривко Л. К."],
                       ["30", "8(495)901-7044", "ул. Кустарная, 28", "Соломоник К.Ф."],
                       ["40", "8(495)359-0014", "ул. Перерва, 2", "Андреев П.Н."]
                   ], [10, 13, 13, 10])

    # Table Фирмы
    draw_datasheet(74, 78, 48, 21, "Фирмы",
                   ["Код фирмы", "Название фирмы", "Адрес", "Телефон"],
                   [
                       ["100", "Мир", "Чонгарский б-р, 16", "8(495) 152-4001"],
                       ["200", "М. Видео", "Маросейка, 6/8", "8(495) 923-2906"],
                       ["300", "Диал Электроникс", "Новослободская, 14/19", "8(495) 978-1693"]
                   ], [10, 13, 14, 11])

    # Table Товары
    draw_datasheet(50, 42, 96, 32, "Товары (с учетом п.9 и п.10 задания 2)",
                   ["Код", "Наименование", "Марка", "Цена", "Склад", "Кол-во", "Описание"],
                   [
                       ["1", "Телевизор", "LG Electronics СМ-2180К", "$459,00", "10", "15", "Цветной ТВ 21\" стерео"],
                       ["2", "Телевизор", "Philips 25PT9001", "$1 499,00", "10", "8", "Плоский экран BlackLine-S"],
                       ["3", "Телевизор", "Panasonic 25V50R", "$765,00", "10", "12", "Стерео ТВ 25\" телетекст"],
                       ["4", "Телевизор", "LG Electronics CF-14E20B", "$230,00", "10", "20", "Компактный ТВ 14\""],
                       ["5", "Видеомагнитофон", "Panasonic HS-800EE", "$1 400,00", "20", "10", "Hi-Fi 6-головочный ВМ"],
                       ["6", "Видеомагнитофон", "Philips VR-756", "$450,00", "20", "18", "4-головочный ShowView"],
                       ["7", "Видеокамера", "Samsung VP-J55", "$530,00", "30", "14", "Аналоговая Hi8 зум 22x"],
                       ["8", "Видеокамера", "Sharp E37", "$845,00", "30", "9", "ЖК-дисплей 2.5\""],
                       ["9", "Муз. центр", "Panasonic DH32", "$320,00", "40", "25", "Минисистема 3 CD"],
                       ["10", "Муз. центр", "Sony MJ-L1", "$1 289,00", "40", "7", "Hi-Fi микросистема MD"],
                       ["11", "Видеокамера", "Panasonic NV-DX1E", "$2 599,00", "30", "5", "Проф. 3-матричная цифр."]
                   ], [6, 17, 24, 13, 8, 8, 20])

    # Table Продажи
    draw_datasheet(50, 12, 96, 17, "Продажи (фрагмент записей)",
                   ["Дата продажи", "Код фирмы", "Код товара", "Количество", "Скидка (%)", "Примечание"],
                   [
                       ["15.01.2024", "100 (Мир)", "1 (LG СМ-2180К)", "3", "0%", "Продажа без скидки"],
                       ["18.01.2024", "200 (М. Видео)", "2 (Philips 25PT9001)", "2", "5%", "Сезонная скидка 5%"],
                       ["22.01.2024", "300 (Диал Эл.)", "3 (Panasonic 25V50R)", "4", "10%", "Оптовая скидка 10%"],
                       ["... [всего 12 записей]", "... [фирмы 100, 200, 300]", "... [товары 1-10]", "... [контроль остатков]", "... [0-10%]", "... [целостность соблюдена]"]
                   ], [16, 20, 22, 11, 11, 16])

    p2 = os.path.join(DIR_LR6, "fig_data_tables.png")
    fig.savefig(p2, bbox_inches='tight', dpi=300)
    plt.close(fig)
    print("Created:", p2)

    # -------------------------------------------------------------
    # Figure 3: Query Design Grid & Results (Запрос_Продажи_Телевизоров)
    # -------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(11, 7.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Top: QBE Design Grid
    box_qbe = patches.Rectangle((2, 50), 96, 47, facecolor='#FFFFFF', edgecolor='#404040', linewidth=1.2)
    ax.add_patch(box_qbe)
    t_qbe = patches.Rectangle((2, 93), 96, 4, facecolor='#1F4E79', edgecolor='none')
    ax.add_patch(t_qbe)
    ax.text(4, 95, "Конструктор запроса: [Запрос_Продажи_Телевизоров] (Режим SQL / QBE)", fontsize=8.5, fontweight='bold', color='#FFFFFF')

    # Table nodes in upper half
    def draw_qbe_table(x, y, w, h, t_name, flds):
        ax.add_patch(patches.Rectangle((x, y), w, h, facecolor='#FFFFFF', edgecolor='#808080', linewidth=1))
        ax.add_patch(patches.Rectangle((x, y + h - 3), w, 3, facecolor='#D9E1F2', edgecolor='#808080', linewidth=0.8))
        ax.text(x + w/2, y + h - 1.5, t_name, ha='center', va='center', fontsize=7.5, fontweight='bold', color='#1F4E79')
        cy = y + h - 4.5
        for fld in flds:
            ax.text(x + 1, cy, fld, fontsize=6.5, color='#333333')
            cy -= 2.2

    draw_qbe_table(5, 68, 22, 23, "Фирмы", ["Код фирмы", "Название фирмы", "Адрес", "Телефон"])
    draw_qbe_table(38, 68, 24, 23, "Продажи", ["Дата продажи", "Код фирмы", "Код товара", "Количество", "Скидка"])
    draw_qbe_table(73, 68, 22, 23, "Товары", ["Код товара", "Наименование", "Марка", "Цена", "Номер склада"])

    # Connection lines in QBE
    ax.plot([27, 38], [88, 86], color='#0066CC', linewidth=1.5) # Фирмы -> Продажи
    ax.plot([62, 73], [84, 88], color='#0066CC', linewidth=1.5) # Продажи -> Товары

    # Lower half of QBE: Grid
    grid_rows = ["Поле:", "Имя таблицы:", "Сортировка:", "Вывод на экран:", "Условие отбора:"]
    col_flds = [
        ("Дата продажи", "Продажи", "По возрастанию", "[x]", ""),
        ("Наименование", "Товары", "", "[x]", "='Телевизор'"),
        ("Марка", "Товары", "По возрастанию", "[x]", ""),
        ("Количество", "Продажи", "", "[x]", ""),
        ("Название фирмы", "Фирмы", "", "[x]", ""),
        ("Адрес", "Фирмы", "", "[x]", ""),
        ("Телефон", "Фирмы", "", "[x]", "")
    ]
    gy = 64
    for r_idx, r_name in enumerate(grid_rows):
        ax.add_patch(patches.Rectangle((5, gy - r_idx*2.7), 16, 2.7, facecolor='#EAEAEA', edgecolor='#A0A0A0', linewidth=0.6))
        ax.text(6, gy - r_idx*2.7 + 1.35, r_name, fontsize=6.8, fontweight='bold', color='#333333')

    gx = 21
    for fld, tbl, srt, chk, cond in col_flds:
        vals = [fld, tbl, srt, chk, cond]
        for r_idx, val in enumerate(vals):
            ax.add_patch(patches.Rectangle((gx, gy - r_idx*2.7), 10.5, 2.7, facecolor='#FFFFFF', edgecolor='#A0A0A0', linewidth=0.6))
            ax.text(gx + 5.25, gy - r_idx*2.7 + 1.35, val, ha='center', va='center', fontsize=6.2, color='#1F4E79' if r_idx==0 else '#222222')
        gx += 10.5

    # Bottom: Query Results Datasheet
    box_res = patches.Rectangle((2, 3), 96, 43, facecolor='#FFFFFF', edgecolor='#404040', linewidth=1.2)
    ax.add_patch(box_res)
    t_res = patches.Rectangle((2, 42), 96, 4, facecolor='#2F5597', edgecolor='none')
    ax.add_patch(t_res)
    ax.text(4, 44, "Режим Таблицы: Результат выполнения запроса [Запрос_Продажи_Телевизоров]", fontsize=8.5, fontweight='bold', color='#FFFFFF')

    res_headers = ["Дата продажи", "Наименование", "Марка", "Кол-во", "Название фирмы", "Адрес фирмы", "Телефон фирмы"]
    res_widths = [13, 14, 21, 8, 14, 15, 11]
    res_data = [
        ["15.01.2024", "Телевизор", "LG Electronics СМ-2180К", "3", "Мир", "Чонгарский б-р, 16", "8(495) 152-4001"],
        ["18.01.2024", "Телевизор", "Philips 25PT9001", "2", "М. Видео", "Маросейка, 6/8", "8(495) 923-2906"],
        ["22.01.2024", "Телевизор", "Panasonic 25V50R", "4", "Диал Электроникс", "Новослободская, 14/19", "8(495) 978-1693"],
        ["05.02.2024", "Телевизор", "LG Electronics CF-14E20B", "5", "Мир", "Чонгарский б-р, 16", "8(495) 152-4001"],
        ["14.04.2024", "Телевизор", "LG Electronics СМ-2180К", "2", "М. Видео", "Маросейка, 6/8", "8(495) 923-2906"],
        ["28.04.2024", "Телевизор", "Philips 25PT9001", "1", "Диал Электроникс", "Новослободская, 14/19", "8(495) 978-1693"]
    ]
    # Draw headers
    hx = 4
    for h_name, hw in zip(res_headers, res_widths):
        ax.add_patch(patches.Rectangle((hx, 38), hw, 4, facecolor='#E7E6E6', edgecolor='#B0B0B0', linewidth=0.7))
        ax.text(hx + hw/2, 40, h_name, ha='center', va='center', fontsize=6.8, fontweight='bold', color='#333333')
        hx += hw

    ry = 34.5
    for row in res_data:
        rx = 4
        for val, hw in zip(row, res_widths):
            ax.add_patch(patches.Rectangle((rx, ry), hw, 3.5, facecolor='#FFFFFF', edgecolor='#D9D9D9', linewidth=0.5))
            ax.text(rx + hw/2, ry + 1.75, val, ha='center', va='center', fontsize=6.5, color='#222222')
            rx += hw
        ry -= 3.5

    p3 = os.path.join(DIR_LR6, "fig_query_selection.png")
    fig.savefig(p3, bbox_inches='tight', dpi=300)
    plt.close(fig)
    print("Created:", p3)

    # -------------------------------------------------------------
    # Figure 4: Advanced Queries (Crosstab & Action Queries)
    # -------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(11, 7.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Top: Crosstab query
    box_cross = patches.Rectangle((2, 53), 96, 44, facecolor='#FFFFFF', edgecolor='#404040', linewidth=1.2)
    ax.add_patch(box_cross)
    t_c = patches.Rectangle((2, 93), 96, 4, facecolor='#1F4E79', edgecolor='none')
    ax.add_patch(t_c)
    ax.text(4, 95, "Перекрестный запрос: [Запрос_Перекрестный_Продажи] (Объемы реализации фирмам)", fontsize=8.5, fontweight='bold', color='#FFFFFF')

    # Crosstab SQL Box
    sql_box = patches.Rectangle((5, 78), 90, 13, facecolor='#F2F4F7', edgecolor='#B0C4DE', linewidth=0.8)
    ax.add_patch(sql_box)
    sql_text = ("TRANSFORM Sum(Продажи.Количество) AS [Количество]\n"
                "SELECT Товары.Наименование\n"
                "FROM Фирмы INNER JOIN (Товары INNER JOIN Продажи ON Товары.[Код товара] = Продажи.[Код товара])\n"
                "     ON Фирмы.[Код фирмы] = Продажи.[Код фирмы]\n"
                "GROUP BY Товары.Наименование\n"
                "PIVOT Фирмы.[Название фирмы];")
    ax.text(7, 84.5, sql_text, fontsize=6.5, family='monospace', color='#002060', va='center')

    # Crosstab Result Table
    c_heads = ["Наименование товара", "Диал Электроникс", "М. Видео", "Мир", "Итого (по товару)"]
    c_widths = [26, 16, 16, 16, 16]
    c_data = [
        ["Видеокамера", "—", "1 шт.", "2 шт.", "3 шт."],
        ["Видеомагнитофон", "3 шт.", "2 шт.", "—", "5 шт."],
        ["Музыкальный центр", "5 шт.", "—", "1 шт.", "6 шт."],
        ["Телевизор", "5 шт.", "4 шт.", "8 шт.", "17 шт."],
        ["ИТОГО (по фирмам)", "13 шт.", "7 шт.", "11 шт.", "31 шт."]
    ]
    cx = 5
    for ch, cw in zip(c_heads, c_widths):
        ax.add_patch(patches.Rectangle((cx, 73), cw, 4, facecolor='#D9E1F2', edgecolor='#808080', linewidth=0.8))
        ax.text(cx + cw/2, 75, ch, ha='center', va='center', fontsize=7, fontweight='bold', color='#1F4E79')
        cx += cw

    cry = 69.5
    for r_idx, crow in enumerate(c_data):
        cx = 5
        is_tot = (r_idx == len(c_data) - 1)
        bg = '#EAEAEA' if is_tot else ('#FFFFFF' if r_idx % 2 == 0 else '#F9F9F9')
        for cval, cw in zip(crow, c_widths):
            ax.add_patch(patches.Rectangle((cx, cry), cw, 3.5, facecolor=bg, edgecolor='#D9D9D9', linewidth=0.6))
            ax.text(cx + cw/2, cry + 1.75, cval, ha='center', va='center', fontsize=6.8, 
                    fontweight='bold' if is_tot or cval.endswith('шт.') and cx>30 else 'normal', color='#222222')
            cx += cw
        cry -= 3.5

    # Bottom: Action queries workflow
    box_act = patches.Rectangle((2, 3), 96, 47, facecolor='#FFFFFF', edgecolor='#404040', linewidth=1.2)
    ax.add_patch(box_act)
    t_a = patches.Rectangle((2, 46), 96, 4, facecolor='#2F5597', edgecolor='none')
    ax.add_patch(t_a)
    ax.text(4, 48, "Запросы на изменение данных (Задание 4: Создание, Обновление, Добавление, Удаление)", fontsize=8.5, fontweight='bold', color='#FFFFFF')

    # Step cards
    steps = [
        ("1. Резервная копия", "Создание таблицы", "SELECT * INTO [Товары 1]\nFROM Товары;", "Создана таблица\n[Товары 1] (11 строк)"),
        ("2. Сезонная скидка", "Обновление цен", "UPDATE [Товары 1]\nSET Цена = Цена * 0.9;", "Цены снижены на 10%\nво всей таблице"),
        ("3. Видеотовары", "Создание + Добавление", "SELECT * INTO Видеокамеры\nINSERT INTO Видеокамеры...", "Объединение видеокамер\nи видеомагнитофонов"),
        ("4. Удаление", "Удаление устаревших", "DELETE FROM Видеотовары\nWHERE Марка LIKE '*Sony*';", "Запись Sony E150EE\nуспешно удалена")
    ]
    ax_x = 4.5
    for s_title, s_type, s_sql, s_res in steps:
        sc = patches.Rectangle((ax_x, 6), 21.5, 38, facecolor='#F8F9FA', edgecolor='#8FAADC', linewidth=1)
        ax.add_patch(sc)
        # Header
        shdr = patches.Rectangle((ax_x, 39.5), 21.5, 4.5, facecolor='#D9E1F2', edgecolor='#8FAADC', linewidth=0.8)
        ax.add_patch(shdr)
        ax.text(ax_x + 10.75, 41.75, s_title, ha='center', va='center', fontsize=7.2, fontweight='bold', color='#1F4E79')
        ax.text(ax_x + 10.75, 37.5, s_type, ha='center', va='center', fontsize=6.5, style='italic', color='#555555')
        # SQL
        ax.add_patch(patches.Rectangle((ax_x + 1, 19), 19.5, 16, facecolor='#FFFFFF', edgecolor='#D9D9D9', linewidth=0.6))
        ax.text(ax_x + 2, 27, s_sql, fontsize=5.8, family='monospace', color='#003366', va='center')
        # Result badge
        rb = patches.Rectangle((ax_x + 1, 8), 19.5, 9, facecolor='#E2EFDA', edgecolor='#A9D08E', linewidth=0.8)
        ax.add_patch(rb)
        ax.text(ax_x + 10.75, 12.5, s_res, ha='center', va='center', fontsize=6.2, fontweight='bold', color='#375623')
        ax_x += 23.5

    p4 = os.path.join(DIR_LR6, "fig_action_queries.png")
    fig.savefig(p4, bbox_inches='tight', dpi=300)
    plt.close(fig)
    print("Created:", p4)

    # -------------------------------------------------------------
    # Figure 5: Forms & Reports Mockups (Задание 5)
    # -------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(11, 7.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Left: Form Фирмы и Контактные лица (Master-Detail)
    box_f = patches.Rectangle((2, 3), 46, 94, facecolor='#FFFFFF', edgecolor='#404040', linewidth=1.2)
    ax.add_patch(box_f)
    t_f = patches.Rectangle((2, 93), 46, 4, facecolor='#1F4E79', edgecolor='none')
    ax.add_patch(t_f)
    ax.text(4, 95, "Форма: [Фирмы] с подчиненной [Контактные лица]", fontsize=7.5, fontweight='bold', color='#FFFFFF')

    # Master fields
    my = 87
    m_fields = [
        ("Код фирмы:", "100"),
        ("Название фирмы:", "Мир"),
        ("Адрес:", "Чонгарский б-р, 16"),
        ("Телефон:", "8(495) 152-4001")
    ]
    for lbl, val in m_fields:
        ax.text(4, my, lbl, fontsize=6.8, fontweight='bold', color='#333333')
        ax.add_patch(patches.Rectangle((19, my - 1.2), 27, 3.2, facecolor='#FFFFFF', edgecolor='#A6A6A6', linewidth=0.8))
        ax.text(21, my + 0.4, val, fontsize=6.8, color='#000000')
        my -= 5.0

    # Subform box
    ax.text(4, 65, "Подчиненная форма: Сотрудники фирмы", fontsize=7, fontweight='bold', color='#1F4E79')
    sub_box = patches.Rectangle((4, 20), 42, 43, facecolor='#F9F9F9', edgecolor='#8FAADC', linewidth=1)
    ax.add_patch(sub_box)
    sub_hdr = patches.Rectangle((4, 59), 42, 4, facecolor='#D9E1F2', edgecolor='#8FAADC', linewidth=0.8)
    ax.add_patch(sub_hdr)
    ax.text(6, 61, "Фамилия", fontsize=6.2, fontweight='bold')
    ax.text(17, 61, "Имя", fontsize=6.2, fontweight='bold')
    ax.text(25, 61, "Должность", fontsize=6.2, fontweight='bold')
    ax.text(37, 61, "Телефон", fontsize=6.2, fontweight='bold')

    sub_rows = [
        ("Смирнов", "Алексей", "Директор", "8(495)334-1122"),
        ("Ковалева", "Елена", "Менеджер", "8(495)421-5566")
    ]
    sry = 54
    for srow in sub_rows:
        ax.add_patch(patches.Rectangle((4.5, sry), 41, 4, facecolor='#FFFFFF', edgecolor='#D9D9D9', linewidth=0.5))
        ax.text(6, sry + 1.5, srow[0], fontsize=6.2)
        ax.text(17, sry + 1.5, srow[1], fontsize=6.2)
        ax.text(25, sry + 1.5, srow[2], fontsize=6.2)
        ax.text(37, sry + 1.5, srow[3], fontsize=6.2)
        sry -= 5

    # Navigation buttons
    btn_bar = patches.Rectangle((4, 6), 42, 10, facecolor='#EAEAEA', edgecolor='#A6A6A6', linewidth=0.8)
    ax.add_patch(btn_bar)
    ax.text(25, 13, "Кнопки управления записями:", ha='center', fontsize=6.2, fontweight='bold', color='#555555')
    btns = ["|<<", "< Назад", "Вперед >", ">>|", "+ Добавить"]
    bx = 5
    for b in btns:
        ax.add_patch(patches.Rectangle((bx, 7.5), 7.6, 4, facecolor='#FFFFFF', edgecolor='#808080', linewidth=0.7))
        ax.text(bx + 3.8, 9.5, b, ha='center', va='center', fontsize=5.8, fontweight='bold', color='#1F4E79')
        bx += 8.2

    # Right: Report Музыкальные центры
    box_r = patches.Rectangle((52, 3), 46, 94, facecolor='#FFFFFF', edgecolor='#404040', linewidth=1.2)
    ax.add_patch(box_r)
    t_r = patches.Rectangle((52, 93), 46, 4, facecolor='#2F5597', edgecolor='none')
    ax.add_patch(t_r)
    ax.text(54, 95, "Отчет: [Отчет_Музыкальные_Центры] (Предварительный просмотр)", fontsize=7.5, fontweight='bold', color='#FFFFFF')

    # Page header (Date)
    ax.add_patch(patches.Rectangle((54, 88), 42, 4, facecolor='#F2F2F2', edgecolor='none'))
    ax.text(55, 90, "Дата формирования: =Date() [05.11.2026]", fontsize=6.5, style='italic', color='#555555')

    # Report Title
    ax.add_patch(patches.Rectangle((54, 76), 42, 10, facecolor='#D9E1F2', edgecolor='#8FAADC', linewidth=1))
    ax.text(75, 82, "ОТЧЕТ ПО РЕАЛИЗАЦИИ\nМУЗЫКАЛЬНЫХ ЦЕНТРОВ", ha='center', va='center', fontsize=8.5, fontweight='bold', color='#1F4E79')

    # Details header
    ax.add_patch(patches.Rectangle((54, 70), 42, 4.5, facecolor='#1F4E79', edgecolor='#1F4E79'))
    ax.text(56, 72.2, "Марка центра", fontsize=6.8, fontweight='bold', color='#FFFFFF')
    ax.text(74, 72.2, "Поставщик", fontsize=6.8, fontweight='bold', color='#FFFFFF')
    ax.text(88, 72.2, "Цена ($)", fontsize=6.8, fontweight='bold', color='#FFFFFF')

    # Details rows
    rep_rows = [
        ("Panasonic DH32", "Диал Электроникс", "320,00 $"),
        ("Sony MJ-L1", "Мир", "1 289,00 $")
    ]
    r_y = 65
    for rmark, rsupp, rpr in rep_rows:
        ax.add_patch(patches.Rectangle((54, r_y), 42, 4.5, facecolor='#FFFFFF', edgecolor='#D9D9D9', linewidth=0.6))
        ax.text(56, r_y + 2, rmark, fontsize=6.8, color='#000000')
        ax.text(74, r_y + 2, rsupp, fontsize=6.8, color='#000000')
        ax.text(88, r_y + 2, rpr, fontsize=6.8, color='#000000')
        r_y -= 5

    # Report footer with Avg Price
    ax.add_patch(patches.Rectangle((54, 46), 42, 11, facecolor='#E2EFDA', edgecolor='#A9D08E', linewidth=1))
    ax.text(56, 53, "ИТОГИ ПО КАТЕГОРИИ:", fontsize=7, fontweight='bold', color='#375623')
    ax.text(56, 49, "Средняя цена (=Avg([Цена])):", fontsize=6.8, color='#375623')
    ax.text(88, 49, "804,50 $", fontsize=7.5, fontweight='bold', color='#A00000')

    # Report design annotations
    ax.add_patch(patches.Rectangle((54, 15), 42, 27, facecolor='#FFF2CC', edgecolor='#D6B656', linewidth=0.8))
    ax.text(75, 38, "Особенности структуры отчета:", ha='center', fontsize=7, fontweight='bold', color='#B25900')
    notes = [
        "1. Заголовок отчета: надпись в области Заголовка",
        "2. Верхний колонтитул: вычисляемое поле =Date()",
        "3. Область данных: поля Марка, Фирма, Цена",
        "4. Примечание отчета: функция =Avg([Цена])",
        "5. Оформление: линия-разделитель и заливка"
    ]
    ny = 34
    for note in notes:
        ax.text(56, ny, note, fontsize=6.2, color='#333333')
        ny -= 4.2

    p5 = os.path.join(DIR_LR6, "fig_forms_reports.png")
    fig.savefig(p5, bbox_inches='tight', dpi=300)
    plt.close(fig)
    print("Created:", p5)

if __name__ == '__main__':
    make_figures()


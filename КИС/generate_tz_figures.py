import os
from PIL import Image, ImageDraw, ImageFont

font_dir = "C:/Windows/Fonts"
font_regular = os.path.join(font_dir, "segoeui.ttf")
font_bold = os.path.join(font_dir, "segoeuib.ttf")
font_mono = os.path.join(font_dir, "consola.ttf")

def get_font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except:
        return ImageFont.load_default()

f_title = get_font(font_bold, 16)
f_sub = get_font(font_bold, 13)
f_bold = get_font(font_bold, 11)
f_reg = get_font(font_regular, 11)
f_sm = get_font(font_regular, 9.5)
f_sm_bold = get_font(font_bold, 9.5)
f_mono = get_font(font_mono, 10)

out_dir = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС'

# =========================================================================
# 1. ГАРИФУЛЛИН: РИС 1 - МАКЕТ ОТЧЕТА ПО ИСПОЛНЕННЫМ ЗАЯВКАМ
# =========================================================================
def make_garifullin_report():
    w, h = 1000, 560
    im = Image.new("RGB", (w, h), "#FFFFFF")
    draw = ImageDraw.Draw(im)
    
    # Border
    draw.rectangle([0, 0, w-1, h-1], outline="#94A3B8", width=1)
    
    # Top header
    draw.rectangle([0, 0, w, 50], fill="#1E293B")
    draw.text((20, 14), "ООО «BMW102» | КИС «BMW102-ERP» | Модуль аналитической отчетности цеха", font=f_bold, fill="#FFFFFF")
    draw.rounded_rectangle([w - 180, 10, w - 20, 40], radius=4, fill="#2563EB")
    draw.text((w - 165, 16), "УТВЕРЖДЕНО КИС", font=f_bold, fill="#FFFFFF")
    
    # Document title
    draw.text((20, 65), "ОТЧЕТ ПО КОЛИЧЕСТВУ ИСПОЛНЕННЫХ ЗАЯВОК ЗА ПЕРИОД", font=f_title, fill="#0F172A")
    draw.text((20, 95), "Параметры формирования: Период: с 01.03.2026 по 31.03.2026 | Подразделение: Цех чип-тюнинга | Статус: Исполнена (Акт подписан)", font=f_sm, fill="#475569")
    draw.text((20, 115), "Ответственный мастер-калибровщик: Гарифуллин К.Р. | Дата и время выгрузки: 31.03.2026 18:00:00", font=f_sm, fill="#64748B")
    
    # Table header
    y = 140
    headers = [
        ("№", 35),
        ("Заявка", 75),
        ("Наряд-заказ", 85),
        ("Дата закр.", 75),
        ("Заказчик / Клиент", 150),
        ("Автомобиль (VIN / Гос. знак)", 170),
        ("Пакет модификации", 145),
        ("Акт №", 85),
        ("Сумма (руб)", 85),
        ("Статус", 75)
    ]
    
    x = 20
    draw.rectangle([x, y, w - 20, y + 26], fill="#334155")
    for title, width in headers:
        draw.text((x + 5, y + 5), title, font=f_sm_bold, fill="#FFFFFF")
        x += width
        
    # Table rows
    rows_data = [
        ("1", "SR-00102", "WO-00441", "25.03.2026", "ФГБОУ ВО «УУНиТ»", "BMW 530d G30 (У 102 УУ 102)", "Stage 1 + Euro-2", "WCA-00388", "41 650,00", "Исполнен"),
        ("2", "SR-00105", "WO-00445", "26.03.2026", "ООО «БашНефтеТранс»", "BMW X5 xDrive30d (В 777 ВВ 02)", "Stage 1 Safe", "WCA-00391", "35 000,00", "Исполнен"),
        ("3", "SR-00108", "WO-00449", "27.03.2026", "ФГБОУ ВО «УУНиТ»", "BMW 730Ld G12 (А 102 АА 102)", "Stage 1 + V-Max Off", "WCA-00394", "45 000,00", "Исполнен"),
        ("4", "SR-00112", "WO-00453", "28.03.2026", "ИП Ахметов Р.Ф.", "BMW 320d G20 (М 102 ММ 102)", "Stage 1 + DPF off", "WCA-00398", "38 500,00", "Исполнен"),
        ("5", "SR-00115", "WO-00458", "29.03.2026", "ООО «УралТрансГаз»", "BMW X6 M50d F16 (О 500 ОО 102)", "Stage 1 Custom", "WCA-00402", "55 000,00", "Исполнен"),
        ("6", "SR-00119", "WO-00462", "30.03.2026", "ФГБОУ ВО «УУНиТ»", "BMW 520d G30 (К 102 КК 102)", "Stage 1 Safe (Eco)", "WCA-00406", "36 000,00", "Исполнен"),
        ("7", "SR-00122", "WO-00466", "31.03.2026", "Гайнанов Д.И.", "BMW 430i F32 (С 102 СС 102)", "Stage 1 Bootmod3", "WCA-00410", "42 000,00", "Исполнен"),
    ]
    
    y += 26
    for idx, row in enumerate(rows_data):
        bg = "#F8FAFC" if idx % 2 == 1 else "#FFFFFF"
        draw.rectangle([20, y, w - 20, y + 24], fill=bg, outline="#E2E8F0")
        x = 20
        for i, val in enumerate(row):
            col = "#15803D" if val == "Исполнен" else ("#0F172A" if i != 8 else "#1D4ED8")
            draw.text((x + 5, y + 4), val, font=f_sm, fill=col)
            x += headers[i][1]
        y += 24
        
    # Summary footer
    draw.rectangle([20, y, w - 20, y + 30], fill="#EFF6FF", outline="#BFDBFE")
    draw.text((30, y + 6), "ИТОГО ПОДРАЗДЕЛЕНИЕМ ЗА ПЕРИОД:", font=f_bold, fill="#1E40AF")
    draw.text((320, y + 6), "Исполнено заявок: 46 | Выходной контроль ОК: 100% | Сумма по реестру:", font=f_sm, fill="#1E40AF")
    draw.text((820, y + 6), "1 842 500,00 руб.", font=f_bold, fill="#1D4ED8")
    
    # Stats & Signatures
    y += 45
    draw.rectangle([20, y, 480, y + 95], fill="#F8FAFC", outline="#CBD5E1")
    draw.text((30, y + 8), "Сводные аналитические показатели эффективности:", font=f_sm_bold, fill="#334155")
    draw.text((30, y + 28), "• Средняя длительность исполнения заявки: 3.8 часа (норматив 4.5 ч)", font=f_sm, fill="#475569")
    draw.text((30, y + 46), "• Процент гарантийных обращений / рекламаций: 0.00% (0 инцидентов)", font=f_sm, fill="#16A34A")
    draw.text((30, y + 64), "• Доля корпоративного сектора (УУНиТ и др.): 65.2% от общей выручки", font=f_sm, fill="#475569")
    
    draw.rectangle([500, y, w - 20, y + 95], fill="#F8FAFC", outline="#CBD5E1")
    draw.text((515, y + 8), "Ответственные лица за составление отчета:", font=f_sm_bold, fill="#334155")
    draw.text((515, y + 32), "Ведущий мастер цеха: ____________ / Гарифуллин К.Р. /", font=f_sm, fill="#0F172A")
    draw.text((515, y + 56), "Начальник сервисной службы: ____________ / Ахметов Р.Ф. /", font=f_sm, fill="#0F172A")
    draw.text((515, y + 76), "Отметка бухгалтерии: Реестр проверен, расхождения отсутствуют.", font=f_sm, fill="#64748B")

    out_p = os.path.join(out_dir, 'tz_garifullin_fig1_report.png')
    im.save(out_p)
    print("Saved", out_p)

# =========================================================================
# 2. ГАРИФУЛЛИН: РИС 2 - ДИАЛОГ КОНФИГУРАТОРА
# =========================================================================
def make_garifullin_dialog():
    w, h = 860, 480
    im = Image.new("RGB", (w, h), "#FFFFFF")
    draw = ImageDraw.Draw(im)
    
    # Window Frame
    draw.rectangle([0, 0, w-1, h-1], outline="#94A3B8", width=1)
    draw.rectangle([0, 0, w, 38], fill="#1E293B")
    draw.text((15, 9), "КИС «BMW102-ERP» :: Сеанс-конфигуратор отчета", font=f_bold, fill="#FFFFFF")
    draw.text((w - 240, 11), "Форма: REP-EXEC-01", font=f_sm, fill="#94A3B8")
    
    # Sub-header
    draw.rectangle([0, 38, w, 75], fill="#F1F5F9", outline="#E2E8F0")
    draw.text((15, 46), "Параметры построения аналитического отчета «Количество исполненных заявок за период»", font=f_sub, fill="#0F172A")
    
    # Left Box: Date & Filter Settings
    draw.rectangle([20, 90, 420, 410], fill="#F8FAFC", outline="#E2E8F0")
    draw.text((35, 105), "1. Временной интервал и фильтрация", font=f_sub, fill="#0F172A")
    
    draw.text((35, 138), "Дата начала периода:", font=f_bold, fill="#475569")
    draw.rectangle([35, 158, 200, 188], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((45, 165), "01.03.2026", font=f_reg, fill="#0F172A")
    
    draw.text((220, 138), "Дата окончания периода:", font=f_bold, fill="#475569")
    draw.rectangle([220, 158, 385, 188], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((230, 165), "31.03.2026", font=f_reg, fill="#0F172A")
    
    draw.text((35, 205), "Мастер-калибровщик:", font=f_bold, fill="#475569")
    draw.rectangle([35, 225, 385, 255], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((45, 232), "Гарифуллин К.Р. (Мастер цеха ID-04)", font=f_reg, fill="#0F172A")
    
    draw.text((35, 270), "Статус исполнения наряда:", font=f_bold, fill="#475569")
    draw.rectangle([35, 290, 385, 320], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((45, 297), "Только закрытые с Актом выполненных работ", font=f_reg, fill="#0F172A")
    
    draw.text((35, 335), "Категория услуг:", font=f_bold, fill="#475569")
    draw.rectangle([35, 355, 385, 385], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((45, 362), "Все пакеты (Stage 1 / Euro-2 / Custom)", font=f_reg, fill="#0F172A")

    # Right Box: Grouping & Export options
    draw.rectangle([440, 90, 840, 410], fill="#F8FAFC", outline="#E2E8F0")
    draw.text((455, 105), "2. Группировка и параметры выгрузки", font=f_sub, fill="#0F172A")
    
    opts = [
        ("[X] Группировать данные по мастерам цеха", "#1D4ED8"),
        ("[X] Отображать расшифровку по VIN и гос. знакам ТС", "#1D4ED8"),
        ("[X] Включать суммы начислений и НДС 20%", "#1D4ED8"),
        ("[X] Выводить ТЕХНИЧЕСКИЙ КОММЕНТАРИЙ мастера в примечание", "#B91C1C"),
        ("[ ] Скрывать промежуточные аннулированные наряды", "#64748B"),
        ("[X] Формировать контрольные суммы (SHA-256 файлов дампов)", "#1D4ED8")
    ]
    y_o = 140
    for opt, col in opts:
        draw.text((455, y_o), opt, font=f_sm_bold, fill=col)
        y_o += 28
        
    draw.text((455, 315), "Формат выходного документа:", font=f_bold, fill="#475569")
    draw.rectangle([455, 338, 560, 368], fill="#EFF6FF", outline="#2563EB")
    draw.text((465, 345), "(•) Печать", font=f_bold, fill="#1E40AF")
    draw.rectangle([575, 338, 690, 368], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((585, 345), "( ) MS Excel (.xlsx)", font=f_reg, fill="#475569")
    draw.rectangle([705, 338, 810, 368], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((715, 345), "( ) PDF/A", font=f_reg, fill="#475569")

    # Bottom Actions Bar
    draw.rectangle([0, 425, w, h], fill="#F1F5F9", outline="#E2E8F0")
    draw.rounded_rectangle([w - 180, 435, w - 20, 468], radius=4, fill="#2563EB")
    draw.text((w - 160, 444), "Сформировать", font=f_bold, fill="#FFFFFF")
    draw.rounded_rectangle([w - 330, 435, w - 195, 468], radius=4, fill="#FFFFFF", outline="#CBD5E1")
    draw.text((w - 305, 444), "Сохранить схему", font=f_reg, fill="#475569")
    draw.rounded_rectangle([w - 430, 435, w - 345, 468], radius=4, fill="#FFFFFF", outline="#CBD5E1")
    draw.text((w - 410, 444), "Отмена", font=f_reg, fill="#475569")

    out_p = os.path.join(out_dir, 'tz_garifullin_fig2_dialog.png')
    im.save(out_p)
    print("Saved", out_p)

# =========================================================================
# 3. ГАРИФУЛЛИН: РИС 3 - ERD ПОДПРОЦЕССА ИСПОЛНЕНИЯ ЗАЯВОК
# =========================================================================
def make_garifullin_erd():
    w, h = 950, 480
    im = Image.new("RGB", (w, h), "#FFFFFF")
    draw = ImageDraw.Draw(im)
    
    draw.rectangle([0, 0, w-1, h-1], outline="#94A3B8", width=1)
    draw.rectangle([0, 0, w, 40], fill="#1E293B")
    draw.text((20, 11), "Реляционная модель данных (ERD) подпроцесса исполнения заявок [Студент: Гарифуллин К.Р.]", font=f_bold, fill="#FFFFFF")
    
    def draw_entity(x, y, ew, eh, name, attrs):
        # Header
        draw.rectangle([x, y, x + ew, y + 26], fill="#2563EB", outline="#1D4ED8")
        draw.text((x + 8, y + 5), name, font=f_bold, fill="#FFFFFF")
        # Body
        draw.rectangle([x, y + 26, x + ew, y + eh], fill="#F8FAFC", outline="#CBD5E1")
        ay = y + 30
        for a in attrs:
            col = "#B91C1C" if "[PK]" in a else ("#1D4ED8" if "[FK]" in a else "#0F172A")
            draw.text((x + 6, ay), a, font=f_sm, fill=col)
            ay += 17
            
    # Entities
    draw_entity(30, 60, 220, 175, "Service_Request (Заявка)", [
        "[PK] id_request : INT",
        "[FK] vin_code : VARCHAR(17)",
        "request_date : DATETIME",
        "service_package : VARCHAR",
        "request_status : VARCHAR",
        "comments : TEXT"
    ])
    
    draw_entity(300, 60, 240, 195, "Work_Order (Наряд-заказ)", [
        "[PK] id_order : INT",
        "[FK] id_request : INT [UQ]",
        "[FK] id_employee : INT",
        "order_date : DATETIME",
        "order_status : VARCHAR",
        "estimated_cost : DECIMAL",
        "assigned_bay : VARCHAR"
    ])
    
    draw_entity(590, 60, 330, 215, "Order_Execution (Исполнение)", [
        "[PK] id_execution : INT",
        "[FK] id_order : INT [UQ]",
        "[FK] id_employee : INT",
        "start_time : DATETIME",
        "end_time : DATETIME",
        "stock_file_hash : VARCHAR(64)",
        "mod_file_hash : VARCHAR(64)",
        "tech_comment : TEXT",
        "execution_status : VARCHAR"
    ])
    
    draw_entity(300, 290, 240, 165, "Employee (Сотрудник)", [
        "[PK] id_employee : INT",
        "full_name : VARCHAR(100)",
        "role_position : VARCHAR(50)",
        "specialization : VARCHAR",
        "qualification_level : INT"
    ])
    
    draw_entity(590, 290, 330, 165, "Work_Completion_Act (Акт)", [
        "[PK] id_act : INT",
        "[FK] id_order : INT [UQ]",
        "act_number : VARCHAR(30)",
        "act_date : DATETIME",
        "total_amount : DECIMAL(10,2)",
        "warranty_cert : VARCHAR(50)",
        "client_signature : VARCHAR"
    ])
    
    # Connecting lines & Cardinality
    # Service_Request -> Work_Order
    draw.line([(250, 120), (300, 120)], fill="#2563EB", width=2)
    draw.text((255, 102), "1", font=f_bold, fill="#1E40AF")
    draw.text((280, 102), "0..1", font=f_bold, fill="#1E40AF")
    
    # Work_Order -> Order_Execution
    draw.line([(540, 120), (590, 120)], fill="#2563EB", width=2)
    draw.text((545, 102), "1", font=f_bold, fill="#1E40AF")
    draw.text((570, 102), "0..1", font=f_bold, fill="#1E40AF")
    
    # Work_Order -> Work_Completion_Act
    draw.line([(420, 255), (420, 370), (590, 370)], fill="#2563EB", width=2)
    draw.text((425, 260), "1", font=f_bold, fill="#1E40AF")
    draw.text((565, 352), "0..1", font=f_bold, fill="#1E40AF")
    
    # Employee -> Work_Order & Order_Execution
    draw.line([(360, 290), (360, 255)], fill="#64748B", width=1)
    draw.text((365, 270), "1:N", font=f_sm_bold, fill="#64748B")
    
    draw.line([(480, 290), (480, 240), (650, 240), (650, 275)], fill="#64748B", width=1)
    draw.text((500, 242), "1:N Назначен калибровщиком", font=f_sm, fill="#64748B")

    out_p = os.path.join(out_dir, 'tz_garifullin_fig3_erd.png')
    im.save(out_p)
    print("Saved", out_p)

# =========================================================================
# 4. ГАЙНАНОВ: РИС 1 - МАКЕТ ОТЧЕТА ПО КЛИЕНТАМ ЗА ПЕРИОД
# =========================================================================
def make_gaynanov_report():
    w, h = 1000, 560
    im = Image.new("RGB", (w, h), "#FFFFFF")
    draw = ImageDraw.Draw(im)
    
    draw.rectangle([0, 0, w-1, h-1], outline="#94A3B8", width=1)
    
    # Top header
    draw.rectangle([0, 0, w, 50], fill="#1E293B")
    draw.text((20, 14), "ООО «BMW102» | КИС «BMW102-ERP» | Модуль управления клиентской базой (CRM)", font=f_bold, fill="#FFFFFF")
    draw.rounded_rectangle([w - 180, 10, w - 20, 40], radius=4, fill="#16A34A")
    draw.text((w - 165, 16), "УТВЕРЖДЕНО КИС", font=f_bold, fill="#FFFFFF")
    
    # Title
    draw.text((20, 65), "ОТЧЕТ ПО КЛИЕНТАМ ЗА ПЕРИОД (СВОДНЫЙ АНАЛИЗ И ВЗАИМОРАСЧЕТЫ)", font=f_title, fill="#0F172A")
    draw.text((20, 95), "Параметры формирования: Период: с 01.01.2026 по 31.03.2026 (1 квартал) | Категория: Все клиенты (Юр. и Физ. лица)", font=f_sm, fill="#475569")
    draw.text((20, 115), "Ответственный системный аналитик: Гайнанов Д.И. | Дата и время выгрузки: 31.03.2026 18:00:00", font=f_sm, fill="#64748B")
    
    y = 140
    headers = [
        ("№", 35),
        ("Код", 55),
        ("Наименование / Контрагент", 185),
        ("Категория", 105),
        ("Кол ТС", 55),
        ("Заявок", 55),
        ("Нарядов", 65),
        ("Начислено (руб)", 120),
        ("Скидка", 65),
        ("Оплачено (руб)", 120),
        ("Сальдо (руб)", 95)
    ]
    
    x = 20
    draw.rectangle([x, y, w - 20, y + 26], fill="#1E293B")
    for title, width in headers:
        draw.text((x + 5, y + 5), title, font=f_sm_bold, fill="#FFFFFF")
        x += width
        
    rows_data = [
        ("1", "CL-01", "ФГБОУ ВО «УУНиТ» (Автопарк)", "Корпоративный", "14", "12", "12", "542 000,00", "15%", "542 000,00", "0,00"),
        ("2", "CL-04", "ООО «БашНефтеТранс»", "Корпоративный", "8", "6", "6", "280 000,00", "10%", "280 000,00", "0,00"),
        ("3", "CL-07", "ООО «УралТрансГаз»", "Корпоративный", "10", "8", "8", "395 000,00", "10%", "395 000,00", "0,00"),
        ("4", "CL-12", "АО «БашАвтоТранс»", "Корпоративный", "5", "4", "4", "185 000,00", "5%", "185 000,00", "0,00"),
        ("5", "CL-19", "Гарифуллин Карим Рамилевич", "Сотрудник / VIP", "2", "2", "2", "68 000,00", "20%", "68 000,00", "0,00"),
        ("6", "CL-22", "ИП Ахметов Р.Ф.", "Партнер", "3", "3", "3", "115 000,00", "10%", "115 000,00", "0,00"),
        ("7", "CL-35", "Смирнов Андрей Васильевич", "Физическое лицо", "1", "1", "1", "45 000,00", "0%", "45 000,00", "0,00"),
    ]
    
    y += 26
    for idx, row in enumerate(rows_data):
        bg = "#F8FAFC" if idx % 2 == 1 else "#FFFFFF"
        draw.rectangle([20, y, w - 20, y + 24], fill=bg, outline="#E2E8F0")
        x = 20
        for i, val in enumerate(row):
            col = "#15803D" if val == "0,00" else ("#1E40AF" if i == 2 else "#0F172A")
            draw.text((x + 5, y + 4), val, font=f_sm, fill=col)
            x += headers[i][1]
        y += 24
        
    # Summary footer
    draw.rectangle([20, y, w - 20, y + 30], fill="#F0FDF4", outline="#BBF7D0")
    draw.text((30, y + 6), "ИТОГО ПО КЛИЕНТСКОЙ БАЗЕ ЗА ПЕРИОД:", font=f_bold, fill="#166534")
    draw.text((330, y + 6), "Всего контрагентов: 112 | Автомобилей: 246 | Начислено:", font=f_sm, fill="#166534")
    draw.text((680, y + 6), "4 980 000,00 руб.", font=f_bold, fill="#15803D")
    draw.text((840, y + 6), "Долг: 0,00 руб.", font=f_bold, fill="#166534")
    
    # KPI Box & Signatures
    y += 45
    draw.rectangle([20, y, 480, y + 95], fill="#F8FAFC", outline="#CBD5E1")
    draw.text((30, y + 8), "Ключевые аналитические показатели клиентской базы:", font=f_sm_bold, fill="#334155")
    draw.text((30, y + 28), "• Прирост новых клиентов за 1 квартал: +18 контрагентов (+19.1%)", font=f_sm, fill="#16A34A")
    draw.text((30, y + 46), "• Средний чек на 1 заказ-наряд: 41 850,00 руб. (в т.ч. Stage 1)", font=f_sm, fill="#475569")
    draw.text((30, y + 64), "• Коэффициент лояльности / повторных обращений: 84.5%", font=f_sm, fill="#475569")
    
    draw.rectangle([500, y, w - 20, y + 95], fill="#F8FAFC", outline="#CBD5E1")
    draw.text((515, y + 8), "Ответственные лица за составление отчета:", font=f_sm_bold, fill="#334155")
    draw.text((515, y + 32), "Ведущий системный аналитик: ____________ / Гайнанов Д.И. /", font=f_sm, fill="#0F172A")
    draw.text((515, y + 56), "Руководитель отдела по работе с клиентами: ____________ / Смирнов А.В. /", font=f_sm, fill="#0F172A")
    draw.text((515, y + 76), "Финансово-экономическая служба: Сверка взаиморасчетов завершена.", font=f_sm, fill="#64748B")

    out_p = os.path.join(out_dir, 'tz_gaynanov_fig1_report.png')
    im.save(out_p)
    print("Saved", out_p)

# =========================================================================
# 5. ГАЙНАНОВ: РИС 2 - ДИАЛОГ КОНФИГУРАТОРА
# =========================================================================
def make_gaynanov_dialog():
    w, h = 860, 480
    im = Image.new("RGB", (w, h), "#FFFFFF")
    draw = ImageDraw.Draw(im)
    
    draw.rectangle([0, 0, w-1, h-1], outline="#94A3B8", width=1)
    draw.rectangle([0, 0, w, 38], fill="#1E293B")
    draw.text((15, 9), "КИС «BMW102-ERP» :: Сеанс-конфигуратор отчета", font=f_bold, fill="#FFFFFF")
    draw.text((w - 240, 11), "Форма: REP-CLIENT-02", font=f_sm, fill="#94A3B8")
    
    draw.rectangle([0, 38, w, 75], fill="#F1F5F9", outline="#E2E8F0")
    draw.text((15, 46), "Параметры построения аналитического отчета «Отчет по клиентам за период»", font=f_sub, fill="#0F172A")
    
    # Left Box
    draw.rectangle([20, 90, 420, 410], fill="#F8FAFC", outline="#E2E8F0")
    draw.text((35, 105), "1. Период анализа и фильтрация клиентов", font=f_sub, fill="#0F172A")
    
    draw.text((35, 138), "Дата начала периода:", font=f_bold, fill="#475569")
    draw.rectangle([35, 158, 200, 188], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((45, 165), "01.01.2026", font=f_reg, fill="#0F172A")
    
    draw.text((220, 138), "Дата окончания периода:", font=f_bold, fill="#475569")
    draw.rectangle([220, 158, 385, 188], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((230, 165), "31.03.2026", font=f_reg, fill="#0F172A")
    
    draw.text((35, 205), "Категория контрагента:", font=f_bold, fill="#475569")
    draw.rectangle([35, 225, 385, 255], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((45, 232), "Все категории (Корпоративные + Физ. лица)", font=f_reg, fill="#0F172A")
    
    draw.text((35, 270), "Конкретный клиент (опционально):", font=f_bold, fill="#475569")
    draw.rectangle([35, 290, 385, 320], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((45, 297), "ФГБОУ ВО «УУНиТ» (Код: CL-01)", font=f_reg, fill="#0F172A")
    
    draw.text((35, 335), "Состояние взаиморасчетов:", font=f_bold, fill="#475569")
    draw.rectangle([35, 355, 385, 385], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((45, 362), "Все состояния (С долгом / Полный расчет)", font=f_reg, fill="#0F172A")

    # Right Box
    draw.rectangle([440, 90, 840, 410], fill="#F8FAFC", outline="#E2E8F0")
    draw.text((455, 105), "2. Структура вывода и группировка", font=f_sub, fill="#0F172A")
    
    opts = [
        ("[X] Детализация по закрепленным автомобилям (VIN / модель)", "#1D4ED8"),
        ("[X] Расшифровка оформленных заявок и наряд-заказов", "#1D4ED8"),
        ("[X] Расчет персональной скидки клиента и экономии", "#1D4ED8"),
        ("[X] Контроль дебиторской задолженности и неоплаченных счетов", "#B91C1C"),
        ("[X] Расчет показателя средней периодичности обслуживания ТС", "#1D4ED8"),
        ("[ ] Скрывать клиентов без обращений за период", "#64748B")
    ]
    y_o = 140
    for opt, col in opts:
        draw.text((455, y_o), opt, font=f_sm_bold, fill=col)
        y_o += 28
        
    draw.text((455, 315), "Формат выходного документа:", font=f_bold, fill="#475569")
    draw.rectangle([455, 338, 560, 368], fill="#EFF6FF", outline="#2563EB")
    draw.text((465, 345), "(•) Печать", font=f_bold, fill="#1E40AF")
    draw.rectangle([575, 338, 690, 368], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((585, 345), "( ) MS Excel (.xlsx)", font=f_reg, fill="#475569")
    draw.rectangle([705, 338, 810, 368], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((715, 345), "( ) PDF/A", font=f_reg, fill="#475569")

    # Bottom Actions Bar
    draw.rectangle([0, 425, w, h], fill="#F1F5F9", outline="#E2E8F0")
    draw.rounded_rectangle([w - 180, 435, w - 20, 468], radius=4, fill="#16A34A")
    draw.text((w - 160, 444), "Сформировать", font=f_bold, fill="#FFFFFF")
    draw.rounded_rectangle([w - 330, 435, w - 195, 468], radius=4, fill="#FFFFFF", outline="#CBD5E1")
    draw.text((w - 305, 444), "Сохранить схему", font=f_reg, fill="#475569")
    draw.rounded_rectangle([w - 430, 435, w - 345, 468], radius=4, fill="#FFFFFF", outline="#CBD5E1")
    draw.text((w - 410, 444), "Отмена", font=f_reg, fill="#475569")

    out_p = os.path.join(out_dir, 'tz_gaynanov_fig2_dialog.png')
    im.save(out_p)
    print("Saved", out_p)

# =========================================================================
# 6. ГАЙНАНОВ: РИС 3 - ERD ПОДПРОЦЕССА УЧЕТА КЛИЕНТОВ И ЗАЯВОК
# =========================================================================
def make_gaynanov_erd():
    w, h = 950, 480
    im = Image.new("RGB", (w, h), "#FFFFFF")
    draw = ImageDraw.Draw(im)
    
    draw.rectangle([0, 0, w-1, h-1], outline="#94A3B8", width=1)
    draw.rectangle([0, 0, w, 40], fill="#1E293B")
    draw.text((20, 11), "Реляционная модель данных (ERD) подпроцесса учета клиентов и заказов [Студент: Гайнанов Д.И.]", font=f_bold, fill="#FFFFFF")
    
    def draw_entity(x, y, ew, eh, name, attrs):
        draw.rectangle([x, y, x + ew, y + 26], fill="#16A34A", outline="#15803D")
        draw.text((x + 8, y + 5), name, font=f_bold, fill="#FFFFFF")
        draw.rectangle([x, y + 26, x + ew, y + eh], fill="#F8FAFC", outline="#CBD5E1")
        ay = y + 30
        for a in attrs:
            col = "#B91C1C" if "[PK]" in a else ("#1D4ED8" if "[FK]" in a else "#0F172A")
            draw.text((x + 6, ay), a, font=f_sm, fill=col)
            ay += 17
            
    # Entities
    draw_entity(30, 60, 240, 195, "Client (Клиент / Контрагент)", [
        "[PK] id_client : INT",
        "client_name : VARCHAR(100)",
        "client_type : VARCHAR(30)",
        "phone_number : VARCHAR(20)",
        "email : VARCHAR(50)",
        "discount_percent : DECIMAL",
        "inn_kpp : VARCHAR(25)"
    ])
    
    draw_entity(340, 60, 250, 205, "Vehicle (Автомобиль клиента)", [
        "[PK] vin_code : VARCHAR(17)",
        "[FK] id_client : INT",
        "license_plate : VARCHAR(15)",
        "brand : VARCHAR(50)",
        "model : VARCHAR(50)",
        "engine_model : VARCHAR(50)",
        "ecu_type : VARCHAR(50)",
        "manufacture_year : INT"
    ])
    
    draw_entity(660, 60, 260, 195, "Service_Request (Заявка ТС)", [
        "[PK] id_request : INT",
        "[FK] vin_code : VARCHAR(17)",
        "request_date : DATETIME",
        "service_package : VARCHAR",
        "request_status : VARCHAR",
        "desired_date : DATETIME",
        "comments : TEXT"
    ])
    
    draw_entity(340, 300, 250, 160, "Work_Order (Наряд-заказ)", [
        "[PK] id_order : INT",
        "[FK] id_request : INT [UQ]",
        "[FK] id_employee : INT",
        "order_date : DATETIME",
        "order_status : VARCHAR",
        "estimated_cost : DECIMAL"
    ])
    
    draw_entity(660, 300, 260, 160, "Work_Completion_Act (Акт)", [
        "[PK] id_act : INT",
        "[FK] id_order : INT [UQ]",
        "act_number : VARCHAR(30)",
        "act_date : DATETIME",
        "total_amount : DECIMAL(10,2)",
        "payment_status : VARCHAR"
    ])
    
    # Connections
    # Client -> Vehicle (1 : N)
    draw.line([(270, 120), (340, 120)], fill="#16A34A", width=2)
    draw.text((275, 102), "1", font=f_bold, fill="#166534")
    draw.text((320, 102), "1..N", font=f_bold, fill="#166534")
    
    # Vehicle -> Service_Request (1 : N)
    draw.line([(590, 120), (660, 120)], fill="#16A34A", width=2)
    draw.text((595, 102), "1", font=f_bold, fill="#166534")
    draw.text((640, 102), "1..N", font=f_bold, fill="#166534")
    
    # Service_Request -> Work_Order (1 : 0..1)
    draw.line([(760, 255), (760, 280), (465, 280), (465, 300)], fill="#16A34A", width=2)
    draw.text((765, 260), "1", font=f_bold, fill="#166534")
    draw.text((475, 285), "0..1", font=f_bold, fill="#166534")
    
    # Work_Order -> Work_Completion_Act (1 : 0..1)
    draw.line([(590, 370), (660, 370)], fill="#16A34A", width=2)
    draw.text((595, 352), "1", font=f_bold, fill="#166534")
    draw.text((640, 352), "0..1", font=f_bold, fill="#166534")

    out_p = os.path.join(out_dir, 'tz_gaynanov_fig3_erd.png')
    im.save(out_p)
    print("Saved", out_p)

make_garifullin_report()
make_garifullin_dialog()
make_garifullin_erd()
make_gaynanov_report()
make_gaynanov_dialog()
make_gaynanov_erd()
print("All 6 figures generated successfully!")

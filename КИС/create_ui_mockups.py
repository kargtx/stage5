from PIL import Image, ImageDraw, ImageFont
import os

font_dir = "C:/Windows/Fonts"
font_regular = os.path.join(font_dir, "segoeui.ttf")
font_bold = os.path.join(font_dir, "segoeuib.ttf")
font_mono = os.path.join(font_dir, "consola.ttf")

def get_font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except:
        return ImageFont.load_default()

f_title = get_font(font_bold, 18)
f_sub = get_font(font_bold, 14)
f_bold = get_font(font_bold, 12)
f_reg = get_font(font_regular, 12)
f_sm = get_font(font_regular, 10)
f_mono = get_font(font_mono, 11)

def draw_window_frame(draw, w, h, title_text):
    # Window background
    draw.rectangle([0, 0, w, h], fill="#FFFFFF", outline="#CBD5E1", width=1)
    # Header bar
    draw.rectangle([0, 0, w, 40], fill="#1E293B")
    # Title
    draw.text((15, 10), title_text, font=f_bold, fill="#FFFFFF")
    # Status badges / buttons on top right
    draw.rounded_rectangle([w - 180, 8, w - 15, 32], radius=4, fill="#2563EB")
    draw.text((w - 165, 12), "КИС «BMW102-ERP»", font=f_bold, fill="#FFFFFF")
    # Sub-toolbar
    draw.rectangle([0, 40, w, 75], fill="#F1F5F9", outline="#E2E8F0", width=1)

def make_ui_step1():
    w, h = 900, 480
    im = Image.new("RGB", (w, h), "#FFFFFF")
    draw = ImageDraw.Draw(im)
    draw_window_frame(draw, w, h, "Форма: Заявка на сервисное обслуживание и чип-тюнинг [Заказчик: УУНиТ]")
    
    # Breadcrumbs / status
    draw.text((15, 48), "Главная  /  Заявки автопарка  /  Создание заявки № SR-2026-00102", font=f_sm, fill="#64748B")
    draw.rounded_rectangle([w - 140, 45, w - 15, 69], radius=12, fill="#FEF3C7", outline="#F59E0B")
    draw.text((w - 125, 50), "• Статус: Черновик", font=f_sm, fill="#B45309")
    
    # Left Column: Client & Vehicle
    draw.rectangle([20, 90, 430, 450], fill="#F8FAFC", outline="#E2E8F0", width=1)
    draw.text((35, 105), "1. Сведения о Заказчике и Автомобиле", font=f_sub, fill="#0F172A")
    
    fields = [
        ("Заказчик:", "ФГБОУ ВО «Уфимский университет науки и технологий» (УУНиТ)"),
        ("Подразделение:", "Транспортный отдел / Гараж УУНиТ"),
        ("Ответственный:", "Гайнанов Данил Ильдарович (+7 917 000-XX-XX)"),
        ("Марка / Модель:", "BMW 530d xDrive (G30 LCI, 2021 г.в.)"),
        ("VIN-номер:", "WBAJA710X0B102994"),
        ("Гос. рег. знак:", "У 102 УУ 102"),
        ("Тип двигателя:", "B57D30 (3.0 Diesel TwinPower Turbo, 249 л.с.)"),
        ("Текущий пробег:", "64 250 км"),
        ("Тип ЭБУ (ECU):", "Bosch MD1CS001 (TC298TP)"),
    ]
    y = 135
    for label, val in fields:
        draw.text((35, y), label, font=f_bold, fill="#475569")
        draw.text((160, y), val, font=f_reg, fill="#0F172A")
        y += 32

    # Right Column: Requested Package & Notes
    draw.rectangle([450, 90, 880, 450], fill="#F8FAFC", outline="#E2E8F0", width=1)
    draw.text((465, 105), "2. Пакет модификации и комментарии Заказчика", font=f_sub, fill="#0F172A")
    
    draw.text((465, 140), "Выбранный комплекс услуг:", font=f_bold, fill="#475569")
    packages = [
        "[X] Программный чип-тюнинг Stage 1 Safe (249 л.с. -> 310 л.с., 620 Нм -> 720 Нм)",
        "[X] Программная калибровка экологического стандарта Euro-2 (DPF/EGR off)",
        "[X] Комплексная входная компьютерная диагностика ЭСУД и датчиков кислорода",
        "[ ] Снятие ограничения максимальной скорости (V-Max Off)",
        "[ ] Настройка отстрелов выхлопной системы (Pops & Bangs)"
    ]
    y = 165
    for pkg in packages:
        color = "#1D4ED8" if "[X]" in pkg else "#64748B"
        draw.text((465, y), pkg, font=f_sm, fill=color)
        y += 26
        
    draw.text((465, 300), "Комментарий заказчика (Данил):", font=f_bold, fill="#475569")
    comment_box = [465, 325, 865, 395]
    draw.rectangle(comment_box, fill="#FFFFFF", outline="#CBD5E1")
    draw.text((475, 335), "Плановое ТО автопарка УУНиТ. Требуется повышение тяги на низах,\nснижение расхода топлива на трассе и исключение прожигов сажевого\nфильтра. Вскрытие блока недопустимо (прошивка строго по OBD-II).", font=f_sm, fill="#334155")
    
    # Action buttons
    draw.rounded_rectangle([465, 410, 650, 440], radius=4, fill="#2563EB")
    draw.text((490, 420), "Отправить в сервис BMW102", font=f_bold, fill="#FFFFFF")
    draw.rounded_rectangle([665, 410, 770, 440], radius=4, fill="#F1F5F9", outline="#CBD5E1")
    draw.text((685, 420), "Сохранить", font=f_bold, fill="#475569")

    out_p = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС\ui_step1_request.png'
    im.save(out_p)
    print("Saved", out_p)

def make_ui_step2():
    w, h = 900, 480
    im = Image.new("RGB", (w, h), "#FFFFFF")
    draw = ImageDraw.Draw(im)
    draw_window_frame(draw, w, h, "Форма: Приемка ТС и открытие Наряд-заказа в КИС [Мастер: Карим]")
    
    draw.text((15, 48), "Главная  /  Производственный цех  /  Наряд-заказ № WO-2026-00441", font=f_sm, fill="#64748B")
    draw.rounded_rectangle([w - 190, 45, w - 15, 69], radius=12, fill="#DBEAFE", outline="#3B82F6")
    draw.text((w - 175, 50), "• Статус: Открыт (В работе)", font=f_sm, fill="#1E40AF")
    
    # Left Column: Reception Details
    draw.rectangle([20, 90, 430, 450], fill="#F8FAFC", outline="#E2E8F0", width=1)
    draw.text((35, 105), "1. Карточка наряд-заказа и осмотра", font=f_sub, fill="#0F172A")
    fields = [
        ("Наряд-заказ:", "WO-2026-00441 (Основание: SR-2026-00102)"),
        ("Дата въезда:", "25.03.2026 10:15:22"),
        ("Мастер-приемщик:", "Гарифуллин Карим Рамилевич (ID: EMP-04)"),
        ("Автомобиль:", "BMW 530d xDrive (VIN: WBAJA710X0B102994)"),
        ("Сверка номеров:", "VIN кузова и ЭБУ совпадает на 100% (ISTA OK)"),
        ("Осмотр кузова:", "Сколы бампера, вмятины отсутствуют. Чистый."),
        ("Топливо в баке:", "80% (ДТ ЛУКОЙЛ ЭКТО Евро-5)"),
        ("Назначенный пост:", "Пост №2 (Программирование и диагностика)"),
        ("Оборудование:", "AutoTuner Master OBD + GYS Flash 100A"),
    ]
    y = 135
    for label, val in fields:
        draw.text((35, y), label, font=f_bold, fill="#475569")
        draw.text((160, y), val, font=f_reg, fill="#0F172A")
        y += 32

    # Right Column: Diagnostic & Checklist
    draw.rectangle([450, 90, 880, 450], fill="#F8FAFC", outline="#E2E8F0", width=1)
    draw.text((465, 105), "2. Первичный чек-лист приемки мастера", font=f_sub, fill="#0F172A")
    
    checks = [
        ("[V] Компьютерная идентификация блока: Bosch MD1CS001", "#16A34A"),
        ("[V] Проверка заводского софта: 0083B200 (Тюнинг ранее не делался)", "#16A34A"),
        ("[V] Отсутствие активных критических DTC ошибок по ЭСУД", "#16A34A"),
        ("[V] Температура ОЖ в норме (88°C), давление топлива 350 bar", "#16A34A"),
        ("[V] Проверка состояния АКБ: 12.6V (требуется бустер при записи)", "#D97706"),
    ]
    y = 145
    for chk, col in checks:
        draw.text((465, y), chk, font=f_bold, fill=col)
        y += 30
        
    draw.text((465, 305), "Решение мастера-приемщика (Карим):", font=f_bold, fill="#475569")
    draw.rectangle([465, 330, 865, 395], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((475, 340), "Автомобиль допущен к проведению чип-тюнинга Stage 1.\nПротивопоказаний по двигателю B57 не обнаружено.\nВыписан электронный наряд-заказ. Направлен на согласование сметы.", font=f_sm, fill="#334155")

    # Action buttons
    draw.rounded_rectangle([465, 410, 670, 440], radius=4, fill="#16A34A")
    draw.text((485, 420), "Утвердить и передать в цех", font=f_bold, fill="#FFFFFF")
    draw.rounded_rectangle([685, 410, 810, 440], radius=4, fill="#F1F5F9", outline="#CBD5E1")
    draw.text((700, 420), "Печать наряда", font=f_bold, fill="#475569")

    out_p = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС\ui_step2_order.png'
    im.save(out_p)
    print("Saved", out_p)

def make_ui_step3():
    w, h = 900, 480
    im = Image.new("RGB", (w, h), "#FFFFFF")
    draw = ImageDraw.Draw(im)
    draw_window_frame(draw, w, h, "Форма: Согласование условий калибровки и сметы [Данил / Карим]")
    
    draw.text((15, 48), "Главная  /  Согласование  /  Наряд-заказ № WO-2026-00441  /  Калькуляция", font=f_sm, fill="#64748B")
    draw.rounded_rectangle([w - 180, 45, w - 15, 69], radius=12, fill="#DCFCE7", outline="#22C55E")
    draw.text((w - 165, 50), "• Статус: Согласовано", font=f_sm, fill="#15803D")
    
    # Table of estimate
    draw.text((25, 95), "Спецификация калибровочных работ и калькуляция сметы:", font=f_sub, fill="#0F172A")
    
    # Table header
    draw.rectangle([25, 125, 875, 155], fill="#1E293B")
    draw.text((35, 133), "№", font=f_bold, fill="#FFFFFF")
    draw.text((70, 133), "Наименование технологической операции", font=f_bold, fill="#FFFFFF")
    draw.text((480, 133), "Нормо-час", font=f_bold, fill="#FFFFFF")
    draw.text((580, 133), "Цена (руб)", font=f_bold, fill="#FFFFFF")
    draw.text((690, 133), "Скидка (УУНиТ)", font=f_bold, fill="#FFFFFF")
    draw.text((795, 133), "Итого (руб)", font=f_bold, fill="#FFFFFF")
    
    rows = [
        ("1", "Входная и выходная диагностика ЭСУД BMW Rheingold", "0.8", "3 000,00", "15%", "2 550,00"),
        ("2", "Инженерная вычитка дампа и архивация в КИС (SHA-256)", "0.5", "2 500,00", "15%", "2 125,00"),
        ("3", "Калибровка карт ПО ЭБУ Bosch MD1 в WinOLS 5.0 (Stage 1)", "2.0", "28 000,00", "15%", "23 800,00"),
        ("4", "Программная конфигурация Euro-2 (De-cat / DPF off)", "1.0", "8 000,00", "15%", "6 800,00"),
        ("5", "Запись прошивки, пересчет CRC и сброс адаптаций", "0.7", "4 000,00", "15%", "3 400,00"),
        ("6", "Выходной тест-драйв со снятием dynamic-логов", "1.0", "3 500,00", "15%", "2 975,00"),
    ]
    y = 155
    for r in rows:
        draw.rectangle([25, y, 875, y + 30], fill="#F8FAFC" if int(r[0]) % 2 == 1 else "#FFFFFF", outline="#E2E8F0")
        draw.text((35, y + 7), r[0], font=f_reg, fill="#0F172A")
        draw.text((70, y + 7), r[1], font=f_reg, fill="#0F172A")
        draw.text((495, y + 7), r[2], font=f_reg, fill="#0F172A")
        draw.text((590, y + 7), r[3], font=f_reg, fill="#0F172A")
        draw.text((710, y + 7), r[4], font=f_reg, fill="#16A34A")
        draw.text((800, y + 7), r[5], font=f_bold, fill="#0F172A")
        y += 30
        
    # Total bar
    draw.rectangle([25, y, 875, y + 35], fill="#EFF6FF", outline="#BFDBFE")
    draw.text((480, y + 8), "ИТОГО К ОПЛАТЕ СО СКИДКОЙ КОРПОРАТИВНОГО КЛИЕНТА:", font=f_bold, fill="#1E40AF")
    draw.text((790, y + 8), "41 650,00 руб.", font=f_bold, fill="#1D4ED8")
    
    # Dialog box
    y += 50
    draw.rectangle([25, y, 875, y + 70], fill="#F8FAFC", outline="#CBD5E1")
    draw.text((35, y + 10), "История обмена информацией (Чат согласования КИС):", font=f_bold, fill="#334155")
    draw.text((35, y + 30), "• [10:28] Карим (Мастер): Смета сформирована с корпоративной скидкой УУНиТ 15%. Мощность: +61 л.с., момент: +100 Нм.", font=f_sm, fill="#475569")
    draw.text((35, y + 48), "• [10:34] Данил (УУНиТ): Смета согласована в полном объеме. Подтверждаю прошивку Stage 1 + Euro-2 по безналичному расчету.", font=f_sm, fill="#15803D")

    out_p = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС\ui_step3_agreement.png'
    im.save(out_p)
    print("Saved", out_p)

def make_ui_step4():
    w, h = 900, 480
    im = Image.new("RGB", (w, h), "#FFFFFF")
    draw = ImageDraw.Draw(im)
    draw_window_frame(draw, w, h, "Форма: Исполнение наряда, прошивка и ТЕХ. КОММЕНТАРИЙ [Карим]")
    
    draw.text((15, 48), "Главная  /  Инженерный модуль  /  Прошивка ЭБУ  /  ID: WO-2026-00441", font=f_sm, fill="#64748B")
    draw.rounded_rectangle([w - 180, 45, w - 15, 69], radius=12, fill="#FEF3C7", outline="#F59E0B")
    draw.text((w - 165, 50), "• Запись завершена", font=f_sm, fill="#B45309")
    
    # Left Column: Calibration Parameters
    draw.rectangle([20, 90, 430, 450], fill="#F8FAFC", outline="#E2E8F0", width=1)
    draw.text((35, 105), "1. Параметры прошивки и архив дампов", font=f_sub, fill="#0F172A")
    fields = [
        ("Программатор:", "AutoTuner Tool (CAN-FD OBD Bus)"),
        ("Процессор ЭБУ:", "Infineon TriCore TC298TP"),
        ("Напряжение сети:", "14.28 V (Стабильно, бустер GYS OK)"),
        ("Файл стока:", "Stock_B57_MD1_0083B200.bin"),
        ("SHA-256 стока:", "8f4a2b91...c31e (Архивирован в КИС)"),
        ("Файл тюнинга:", "BMW102_B57_Stage1_Euro2_v2.bin"),
        ("SHA-256 мода:", "d51c7709...4a88 (Контрольная сумма OK)"),
        ("Давление наддува:", "2.45 bar (абс.) [Сток: 2.15 bar]"),
        ("Макс. EGT (выхлоп):", "720 °C [Безопасный лимит: 800 °C]"),
    ]
    y = 135
    for label, val in fields:
        draw.text((35, y), label, font=f_bold, fill="#475569")
        draw.text((160, y), val, font=f_reg, fill="#0F172A")
        y += 32

    # Right Column: Technical Comment of Karim
    draw.rectangle([450, 90, 880, 450], fill="#F8FAFC", outline="#E2E8F0", width=1)
    draw.text((465, 105), "2. ТЕХНИЧЕСКИЙ КОММЕНТАРИЙ МАСТЕРА", font=f_sub, fill="#B91C1C")
    
    draw.text((465, 135), "Обязательное поле фиксации технических нюансов:", font=f_bold, fill="#475569")
    draw.rectangle([465, 160, 865, 340], fill="#FEFCE8", outline="#CA8A04")
    tech_comment = (
        "МАСТЕР-КАЛИБРОВЩИК: Гарифуллин Карим Рамилевич\n"
        "ДАТА И ВРЕМЯ: 25.03.2026 13:40\n\n"
        "ТЕХНИЧЕСКИЙ ОТЧЕТ:\n"
        "1. Чтение и запись произведены штатно через OBD-II без разбора ЭБУ.\n"
        "2. Выполнена адаптация топливных карт: давление в Common Rail поднято\n"
        "   на 80 bar (до 2580 bar в пике). Время впрыска оптимизировано.\n"
        "3. Отключен опрос датчиков дифференциального давления DPF и лямбд.\n"
        "4. Проведен сброс адаптаций АКПП ZF8HP и сброс сервисных интервалов.\n"
        "5. Выходной лог тест-драйва подтвердил стабильный разгон 0-100 км/ч:\n"
        "   5.1 сек (было 6.0 сек). Детонация отсутствует, дымность 0.00%.\n\n"
        "РЕКОМЕНДАЦИИ ЗАКАЗЧИКУ: Замена масла каждые 7500 км (BMW LL-04 5W-30)."
    )
    draw.text((475, 168), tech_comment, font=f_mono, fill="#713F12")
    
    # Action buttons
    draw.rounded_rectangle([465, 410, 680, 440], radius=4, fill="#16A34A")
    draw.text((485, 420), "Завершить и выпустить Акт", font=f_bold, fill="#FFFFFF")
    draw.rounded_rectangle([695, 410, 840, 440], radius=4, fill="#F1F5F9", outline="#CBD5E1")
    draw.text((710, 420), "Сохранить лог", font=f_bold, fill="#475569")

    out_p = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС\ui_step4_execution.png'
    im.save(out_p)
    print("Saved", out_p)

def make_ui_step5():
    w, h = 900, 480
    im = Image.new("RGB", (w, h), "#FFFFFF")
    draw = ImageDraw.Draw(im)
    draw_window_frame(draw, w, h, "Форма: Выходной контроль и генерация Акта [Карим]")
    
    draw.text((15, 48), "Главная  /  Документооборот  /  Акт выполненных работ № WCA-2026-00388", font=f_sm, fill="#64748B")
    draw.rounded_rectangle([w - 180, 45, w - 15, 69], radius=12, fill="#DCFCE7", outline="#22C55E")
    draw.text((w - 165, 50), "• Готов к выдаче", font=f_sm, fill="#15803D")
    
    # Left Column: Quality Assurance
    draw.rectangle([20, 90, 430, 450], fill="#F8FAFC", outline="#E2E8F0", width=1)
    draw.text((35, 105), "1. Выходной контроль качества (ОТК)", font=f_sub, fill="#0F172A")
    fields = [
        ("Номер Акта:", "WCA-2026-00388 от 25.03.2026"),
        ("Основание:", "Наряд-заказ WO-2026-00441"),
        ("Контролер/Мастер:", "Гарифуллин Карим Рамилевич"),
        ("Тест-драйв:", "Пройден успешно (12 км, смешанный режим)"),
        ("Ошибки в блоках:", "DTC: 0 (Все 44 блока управления чистые)"),
        ("Мощность (стенд):", "310 л.с. (+61 л.с.) / 720 Нм (+100 Нм)"),
        ("Сертификат:", "CERT-102-G30-00994 (Пожизненная гарантия)"),
        ("Стоянка готовых ТС:", "Сектор А-04 (Ключи переданы на выдачу)"),
        ("Уведомление клиенту:", "Отправлено по SMS и в личный кабинет УУНиТ"),
    ]
    y = 135
    for label, val in fields:
        draw.text((35, y), label, font=f_bold, fill="#475569")
        draw.text((160, y), val, font=f_reg, fill="#0F172A")
        y += 32

    # Right Column: Act Preview
    draw.rectangle([450, 90, 880, 450], fill="#F8FAFC", outline="#E2E8F0", width=1)
    draw.text((465, 105), "2. Проект печатной формы Акта (КС-2)", font=f_sub, fill="#0F172A")
    draw.rectangle([465, 135, 865, 395], fill="#FFFFFF", outline="#CBD5E1")
    act_text = (
        "АКТ ВЫПОЛНЕННЫХ РАБОТ № WCA-2026-00388\n"
        "Заказчик: ФГБОУ ВО «УУНиТ»\n"
        "Исполнитель: ООО «BMW102»\n"
        "Автомобиль: BMW 530d xDrive (VIN: WBAJA710X0B102994)\n"
        "----------------------------------------------------------\n"
        "Выполнен комплекс работ по модернизации программного\n"
        "обеспечения ЭБУ (Stage 1) и программной калибровке Euro-2.\n"
        "Претензий по качеству и срокам Заказчик не имеет.\n"
        "Итоговая сумма: 41 650,00 руб. (В т.ч. НДС 20% - 6 941,67 руб.)\n"
        "Гарантия на ПО: Пожизненная (при условии сохранения стока).\n"
        "Исполнитель: / Гарифуллин К.Р. /   Заказчик: / Гайнанов Д.И. /"
    )
    draw.text((475, 145), act_text, font=f_mono, fill="#1E293B")
    
    # Action buttons
    draw.rounded_rectangle([465, 410, 680, 440], radius=4, fill="#2563EB")
    draw.text((485, 420), "Сформировать УПД в Диадок", font=f_bold, fill="#FFFFFF")
    draw.rounded_rectangle([695, 410, 840, 440], radius=4, fill="#F1F5F9", outline="#CBD5E1")
    draw.text((715, 420), "Печать Акта", font=f_bold, fill="#475569")

    out_p = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС\ui_step5_completion.png'
    im.save(out_p)
    print("Saved", out_p)

def make_ui_step6():
    w, h = 900, 480
    im = Image.new("RGB", (w, h), "#FFFFFF")
    draw = ImageDraw.Draw(im)
    draw_window_frame(draw, w, h, "Форма: Выдача ТС и закрытие наряда [Данил / Карим]")
    
    draw.text((15, 48), "Главная  /  Выдача автотранспорта  /  Наряд-заказ № WO-2026-00441", font=f_sm, fill="#64748B")
    draw.rounded_rectangle([w - 180, 45, w - 15, 69], radius=12, fill="#E0E7FF", outline="#6366F1")
    draw.text((w - 165, 50), "• Закрыт / В архиве", font=f_sm, fill="#4338CA")
    
    # Left Column: Handover Checklist
    draw.rectangle([20, 90, 430, 450], fill="#F8FAFC", outline="#E2E8F0", width=1)
    draw.text((35, 105), "1. Протокол выдачи автомобиля заказчику", font=f_sub, fill="#0F172A")
    fields = [
        ("Представитель УУНиТ:", "Гайнанов Данил Ильдарович (Доверенность №14)"),
        ("Мастер сервиса:", "Гарифуллин Карим Рамилевич"),
        ("Дата фактич. выдачи:", "25.03.2026 15:45:00"),
        ("Визуальный осмотр:", "ТС чистое, без новых дефектов. Принято."),
        ("Контрольный пуск ДВС:", "Запуск штатный, обороты ровные, Chek OFF."),
        ("Тестовое ускорение:", "Динамика существенно возросла, отклик на газ."),
        ("Передано заказчику:", "Ключи (2 шт), СТС, Акт, Гарантийный талон"),
        ("Статус оплаты:", "Безналичный расчет (П/П № 4402 от 25.03.2026)"),
        ("Архивация в КИС:", "Заказ переведен в статус ЗАКРЫТ. Хэши в БД."),
    ]
    y = 135
    for label, val in fields:
        draw.text((35, y), label, font=f_bold, fill="#475569")
        draw.text((160, y), val, font=f_reg, fill="#0F172A")
        y += 32

    # Right Column: Certificate and Customer Sign
    draw.rectangle([450, 90, 880, 450], fill="#F8FAFC", outline="#E2E8F0", width=1)
    draw.text((465, 105), "2. Электронная подпись приемки и гарантия", font=f_sub, fill="#0F172A")
    
    draw.rectangle([465, 135, 865, 275], fill="#EFF6FF", outline="#93C5FD")
    cert_text = (
        "ГАРАНТИЙНЫЙ СЕРТИФИКАТ № CERT-102-G30-00994\n"
        "Настоящим автотехцентр «BMW102» удостоверяет, что на\n"
        "автомобиле BMW 530d xDrive произведена лицензионная\n"
        "калибровка ПО ЭБУ. Оригинальный сток-дамп сохранен\n"
        "в защищенном хранилище КИС бессрочно.\n"
        "Предоставляется право на бесплатный откат на сток."
    )
    draw.text((475, 145), cert_text, font=f_sm, fill="#1E40AF")
    
    draw.text((465, 290), "Электронная подпись приемки (ПЭП):", font=f_bold, fill="#475569")
    draw.rectangle([465, 315, 865, 395], fill="#FFFFFF", outline="#CBD5E1")
    draw.text((475, 325), "Документ подписан простой электронной подписью (ПЭП):\nПодписант: Гайнанов Данил Ильдарович (УУНиТ)\nСертификат: SMS-OTP Код подтверждения 849201\nДата/время: 25.03.2026 15:42:18 MSK\nРезультат: Работа принята в полном объеме, замечаний нет.", font=f_sm, fill="#15803D")

    # Action buttons
    draw.rounded_rectangle([465, 410, 680, 440], radius=4, fill="#16A34A")
    draw.text((495, 420), "Закрыть наряд и архивировать", font=f_bold, fill="#FFFFFF")
    draw.rounded_rectangle([695, 410, 850, 440], radius=4, fill="#F1F5F9", outline="#CBD5E1")
    draw.text((710, 420), "Выгрузить в 1С:Гараж", font=f_bold, fill="#475569")

    out_p = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС\ui_step6_handover.png'
    im.save(out_p)
    print("Saved", out_p)

make_ui_step1()
make_ui_step2()
make_ui_step3()
make_ui_step4()
make_ui_step5()
make_ui_step6()
print("All 6 UI mockups created successfully!")

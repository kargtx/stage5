import os
from PIL import Image, ImageDraw, ImageFont

print("=== Generating Corporate Screen Forms (UI Mockups) for DBS in КИС ===")

out_dir = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС"

font_title = ImageFont.truetype(r"C:\Windows\Fonts\segoeuib.ttf", 15)
font_bold = ImageFont.truetype(r"C:\Windows\Fonts\segoeuib.ttf", 12)
font_regular = ImageFont.truetype(r"C:\Windows\Fonts\segoeui.ttf", 12)
font_small = ImageFont.truetype(r"C:\Windows\Fonts\segoeui.ttf", 10)
font_mono = ImageFont.truetype(r"C:\Windows\Fonts\consola.ttf", 11)

def draw_window_frame(draw, w, h, title_text, breadcrumbs="КИС «Автосервис: Чип-Тюнинг» v3.4"):
    # Window background
    draw.rectangle([0, 0, w, h], fill=(248, 250, 252))
    
    # Outer border
    draw.rectangle([0, 0, w-1, h-1], outline=(148, 163, 184), width=1)
    
    # Header bar (Windows / 1C style dark navy)
    draw.rectangle([1, 1, w-2, 38], fill=(30, 58, 138))
    draw.text((16, 10), title_text, fill=(255, 255, 255), font=font_title)
    
    # Window controls (min, max, close)
    draw.rectangle([w-30, 8, w-10, 28], fill=(239, 68, 68))
    draw.text((w-23, 8), "✕", fill=(255, 255, 255), font=font_bold)
    draw.rectangle([w-60, 8, w-40, 28], fill=(51, 65, 85))
    draw.text((w-53, 6), "□", fill=(255, 255, 255), font=font_bold)
    draw.rectangle([w-90, 8, w-70, 28], fill=(51, 65, 85))
    draw.text((w-83, 6), "—", fill=(255, 255, 255), font=font_bold)
    
    # Breadcrumbs toolbar
    draw.rectangle([1, 39, w-2, 70], fill=(241, 245, 249), outline=(226, 232, 240), width=1)
    draw.text((16, 47), f"Модуль: {breadcrumbs}", fill=(71, 85, 105), font=font_small)

def draw_button(draw, x, y, w, h, text, is_primary=False):
    fill = (29, 78, 216) if is_primary else (255, 255, 255)
    border = (30, 58, 138) if is_primary else (203, 213, 225)
    txt_color = (255, 255, 255) if is_primary else (30, 41, 59)
    draw.rounded_rectangle([x, y, x+w, y+h], radius=4, fill=fill, outline=border, width=1)
    draw.text((x + 12, y + 6), text, fill=txt_color, font=font_bold)

def draw_input(draw, x, y, w, h, label, value, is_readonly=False):
    draw.text((x, y - 18), label, fill=(71, 85, 105), font=font_small)
    fill = (241, 245, 249) if is_readonly else (255, 255, 255)
    draw.rounded_rectangle([x, y, x+w, y+h], radius=3, fill=fill, outline=(203, 213, 225), width=1)
    draw.text((x + 8, y + 6), str(value), fill=(15, 23, 42), font=font_regular)

def draw_badge(draw, x, y, text, badge_type="success"):
    colors = {
        "success": ((220, 252, 231), (22, 163, 74), (22, 101, 52)),
        "warning": ((254, 243, 199), (217, 119, 6), (180, 83, 9)),
        "info": ((219, 234, 254), (37, 99, 235), (30, 58, 138)),
        "danger": ((254, 226, 226), (220, 38, 38), (153, 27, 27))
    }
    bg, border, txt = colors.get(badge_type, colors["info"])
    draw.rounded_rectangle([x, y, x+120, y+24], radius=12, fill=bg, outline=border, width=1)
    draw.text((x+14, y+4), text, fill=txt, font=font_bold)

# ================= 1. ЭКРАН 1: РЕГИСТРАЦИЯ КЛИЕНТА И ТС =================
def make_screen_1():
    W, H = 850, 480
    im = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(im)
    draw_window_frame(d, W, H, "КИС «Автосервис» — [Сеанс 1.3] Карточка клиента и транспортного средства", "CRM / База автовладельцев / Регистрация")
    
    # Action buttons
    draw_button(d, 16, 80, 160, 30, "💾 Записать и закрыть", is_primary=True)
    draw_button(d, 186, 80, 140, 30, "➕ Создать заявку")
    draw_button(d, 336, 80, 140, 30, "📜 История визитов")
    draw_badge(d, W - 140, 84, "✔ Активен", "success")

    # Section 1: Данные клиента
    d.rectangle([16, 125, W-16, 235], fill=(255, 255, 255), outline=(226, 232, 240), width=1)
    d.text((26, 133), "1. СВЕДЕНИЯ О КЛИЕНТЕ (АВТОВЛАДЕЛЬЦЕ)", fill=(30, 58, 138), font=font_bold)
    
    draw_input(d, 26, 175, 260, 28, "ФИО владельца / Юрлицо", "Смирнов Алексей Викторович")
    draw_input(d, 306, 175, 170, 28, "Номер телефона", "+7 (917) 845-12-34")
    draw_input(d, 496, 175, 150, 28, "Тип контрагента", "Физическое лицо")
    draw_input(d, 666, 175, 140, 28, "ID клиента (PK)", "CL-00849", is_readonly=True)

    # Section 2: Данные автомобиля
    d.rectangle([16, 255, W-16, 445], fill=(255, 255, 255), outline=(226, 232, 240), width=1)
    d.text((26, 263), "2. ТРАНСПОРТНОЕ СРЕДСТВО И СПЕЦИФИКАЦИЯ ЭБУ", fill=(30, 58, 138), font=font_bold)

    draw_input(d, 26, 305, 240, 28, "VIN-код автомобиля (PK)", "WBAJB71090G128941", is_readonly=False)
    draw_input(d, 286, 305, 150, 28, "Марка ТС", "BMW")
    draw_input(d, 456, 305, 170, 28, "Модель ТС", "530d G30 M-Sport")
    draw_input(d, 646, 305, 160, 28, "Год выпуска", "2021")

    draw_input(d, 26, 365, 240, 28, "Модель двигателя", "3.0d B57D30 (249 л.с.)")
    draw_input(d, 286, 365, 240, 28, "Блок управления (ЭБУ)", "Bosch MD1CS001")
    draw_input(d, 546, 365, 130, 28, "Текущий пробег", "64 250 км")
    draw_input(d, 696, 365, 110, 28, "Госномер", "У 714 РХ 102")

    p = os.path.join(out_dir, "ДБС_Экран_1_Клиент_Автомобиль.png")
    im.save(p)
    print("Saved:", p)

# ================= 2. ЭКРАН 2: ЗАЯВКА НА ТЮНИНГ =================
def make_screen_2():
    W, H = 850, 480
    im = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(im)
    draw_window_frame(d, W, H, "КИС «Автосервис» — [Сеанс 1.5] Первичная заявка на чип-тюнинг и подбор ПО", "Заявки / Согласование пакета / Параметры")
    
    draw_button(d, 16, 80, 170, 30, "✔ Согласовать смету", is_primary=True)
    draw_button(d, 196, 80, 150, 30, "📄 Наряд-заказ (Данил)")
    draw_badge(d, W - 160, 84, "⏳ На согласовании", "warning")

    # Header info
    draw_input(d, 16, 135, 180, 28, "Номер заявки (PK)", "REQ-2026-0312", is_readonly=True)
    draw_input(d, 216, 135, 180, 28, "Дата создания", "26.10.2026 10:15")
    draw_input(d, 416, 135, 200, 28, "Автомобиль (VIN)", "WBAJB71090G128941", is_readonly=True)
    draw_input(d, 636, 135, 180, 28, "Мастер-приемщик", "Гайнанов Данил (ЭАС-514)")

    # Tuning package comparison table
    d.rectangle([16, 190, W-16, 445], fill=(255, 255, 255), outline=(226, 232, 240), width=1)
    d.text((26, 200), "ВЫБОР ПАКЕТА МОДИФИКАЦИИ И РАСЧЕТНЫЕ ХАРАКТЕРИСТИКИ", fill=(30, 58, 138), font=font_bold)

    # Table headers
    headers = ["Параметр", "Заводские значения", "Пакет Stage 1 (Выбран)", "Пакет Stage 2"]
    cols_x = [30, 230, 430, 640]
    d.rectangle([26, 225, W-26, 255], fill=(238, 242, 255))
    for i, h_txt in enumerate(headers):
        d.text((cols_x[i], 232), h_txt, fill=(30, 58, 138), font=font_bold)

    rows = [
        ("Мощность ДВС (л.с.)", "249 л.с. @ 4000 rpm", "310 л.с. (+61 л.с.)", "345 л.с. (+96 л.с.)"),
        ("Крутящий момент (Нм)", "620 Нм @ 2000 rpm", "720 Нм (+100 Нм)", "780 Нм (+160 Нм)"),
        ("Экологический стандарт", "Евро-6 (DPF + AdBlue)", "Евро-6 (Штатно)", "Евро-2 (DPF/EGR off)"),
        ("Ограничитель скорости", "250 км/ч (Vmax on)", "Снят (Vmax off)", "Снят (Vmax off)"),
        ("Стоимость пакета", "—", "42 000 руб.", "68 000 руб.")
    ]
    curr_y = 260
    for r in rows:
        d.line([26, curr_y+24, W-26, curr_y+24], fill=(241, 245, 249), width=1)
        for i, val in enumerate(r):
            fnt = font_bold if i == 2 else font_regular
            col = (29, 78, 216) if i == 2 else (15, 23, 42)
            d.text((cols_x[i], curr_y+4), val, fill=col, font=fnt)
        curr_y += 30

    d.text((30, 420), "Выбранный пакет: Stage 1 DynoTune | Согласованная смета: 42 000 руб. | Предоплата: 0%", fill=(22, 101, 52), font=font_bold)

    p = os.path.join(out_dir, "ДБС_Экран_2_Заявка_Тюнинг.png")
    im.save(p)
    print("Saved:", p)

# ================= 3. ЭКРАН 3: НАРЯД-ЗАКАЗ (ДАНИЛ) =================
def make_screen_3():
    W, H = 850, 480
    im = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(im)
    draw_window_frame(d, W, H, "КИС «Автосервис» — [Сеанс 1.8] Наряд-заказ № НЗ-2026/04-118 (Документ Данила)", "Производство / Наряд-заказы / Первичный документ")
    
    draw_button(d, 16, 80, 170, 30, "🖨 Печать наряда", is_primary=True)
    draw_button(d, 196, 80, 160, 30, "🔧 Передать в работу")
    draw_button(d, 366, 80, 140, 30, "🔍 Диагностика")
    draw_badge(d, W - 150, 84, "✔ Утвержден", "success")

    # Document Header Card
    d.rectangle([16, 125, W-16, 215], fill=(255, 255, 255), outline=(226, 232, 240), width=1)
    draw_input(d, 26, 145, 170, 26, "Номер наряд-заказа", "НЗ-2026/04-118", is_readonly=True)
    draw_input(d, 206, 145, 150, 26, "Дата открытия", "26.10.2026 10:30")
    draw_input(d, 366, 145, 220, 26, "Мастер-приемщик (Данил)", "Гайнанов Д.И. (ЭАС-514)")
    draw_input(d, 596, 145, 220, 26, "Клиент", "Смирнов А.В. (+79178451234)")

    draw_input(d, 26, 185, 240, 26, "Автомобиль", "BMW 530d G30 (2021)")
    draw_input(d, 276, 185, 240, 26, "VIN-код", "WBAJB71090G128941", is_readonly=True)
    draw_input(d, 526, 185, 290, 26, "Заявка-основание", "REQ-2026-0312 от 26.10.2026")

    # Table of Services
    d.rectangle([16, 230, W-16, 445], fill=(255, 255, 255), outline=(226, 232, 240), width=1)
    d.text((26, 240), "СПЕЦИФИКАЦИЯ РАБОТ И НОРМО-ЧАСОВ", fill=(30, 58, 138), font=font_bold)

    headers = ["№", "Код услуги", "Наименование технологической операции", "Норма (ч)", "Ставка", "Сумма (руб.)"]
    cols_x = [30, 60, 160, 560, 640, 720]
    d.rectangle([26, 265, W-26, 290], fill=(219, 234, 254))
    for i, h_txt in enumerate(headers):
        d.text((cols_x[i], 270), h_txt, fill=(30, 58, 138), font=font_bold)

    items = [
        ("1", "SRV-01", "Компьютерная диагностика систем ДВС (входной контроль)", "1.0", "3 000", "3 000.00"),
        ("2", "TUN-03", "Программная оптимизация ЭБУ Stage 1 DynoTune (B57)", "4.0", "8 500", "34 000.00"),
        ("3", "TST-02", "Замер мощностных параметров и логирование на стенде", "1.5", "3 333", "5 000.00")
    ]
    curr_y = 295
    for item in items:
        d.line([26, curr_y+24, W-26, curr_y+24], fill=(241, 245, 249), width=1)
        for i, val in enumerate(item):
            d.text((cols_x[i], curr_y+4), val, fill=(15, 23, 42), font=font_regular)
        curr_y += 30

    d.line([26, 395, W-26, 395], fill=(30, 58, 138), width=2)
    d.text((450, 410), "ИТОГО ПО НАРЯД-ЗАКАЗУ (СОГЛАСОВАНО):", fill=(30, 58, 138), font=font_bold)
    d.text((710, 408), "42 000.00 руб.", fill=(29, 78, 216), font=font_title)

    p = os.path.join(out_dir, "ДБС_Экран_3_Наряд_Заказ.png")
    im.save(p)
    print("Saved:", p)

# ================= 4. ЭКРАН 4: ДИАГНОСТИЧЕСКАЯ КАРТА =================
def make_screen_4():
    W, H = 850, 480
    im = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(im)
    draw_window_frame(d, W, H, "КИС «Автосервис» — [Сеанс 1.10] Диагностическая карта ДВС № ДК-118", "Диагностика / Входной контроль / Сканирование OBD-II")
    
    draw_button(d, 16, 80, 160, 30, "✔ Допуск к прошивке", is_primary=True)
    draw_button(d, 186, 80, 150, 30, "🖨 Распечатать отчет")
    draw_badge(d, W - 150, 84, "✔ Ошибок нет", "success")

    # Header Card
    d.rectangle([16, 125, W-16, 205], fill=(255, 255, 255), outline=(226, 232, 240), width=1)
    draw_input(d, 26, 145, 160, 26, "Карта диагностики (PK)", "DIAG-2026-0118", is_readonly=True)
    draw_input(d, 196, 145, 160, 26, "Наряд-заказ (FK)", "НЗ-2026/04-118", is_readonly=True)
    draw_input(d, 366, 145, 180, 26, "Диагност", "Хакимов Р.А. (Таб. №14)")
    draw_input(d, 556, 145, 140, 26, "Этап проверки", "Входной контроль")
    draw_input(d, 706, 145, 110, 26, "DTC статус", "NONE", is_readonly=True)

    # Diagnostic Sensors Table
    d.rectangle([16, 220, W-16, 445], fill=(255, 255, 255), outline=(226, 232, 240), width=1)
    d.text((26, 230), "РЕЗУЛЬТАТЫ СКАНИРОВАНИЯ И ТЕЛЕМЕТРИЯ СИСТЕМ ДВС", fill=(30, 58, 138), font=font_bold)

    headers = ["Система / Параметр", "Фактическое значение", "Эталонный допуск", "Статус параметра"]
    cols_x = [30, 260, 480, 680]
    d.rectangle([26, 255, W-26, 280], fill=(238, 242, 255))
    for i, h_txt in enumerate(headers):
        d.text((cols_x[i], 260), h_txt, fill=(30, 58, 138), font=font_bold)

    data = [
        ("Коды неисправностей (DTC)", "Кодов нет (dtc_fault_code = NONE)", "Отсутствие критических ошибок", "НОРМА (OK)"),
        ("Давление наддува турбокомпрессора", "1 620 мбар @ 3000 rpm", "1 550 - 1 650 мбар", "НОРМА (OK)"),
        ("Давление топлива в рампе Rail", "2 180 бар @ Full Throttle", "2 100 - 2 250 бар", "НОРМА (OK)"),
        ("Температура отработавших газов (EGT)", "680 °C (пик)", "до 820 °C", "НОРМА (OK)"),
        ("Балансировка форсунок (коррекция)", "+0.15 ... -0.12 мг/цикл", "-0.8 ... +0.8 мг/цикл", "НОРМА (OK)")
    ]
    curr_y = 285
    for item in data:
        d.line([26, curr_y+24, W-26, curr_y+24], fill=(241, 245, 249), width=1)
        for i, val in enumerate(item):
            col = (22, 101, 52) if i == 3 else (15, 23, 42)
            d.text((cols_x[i], curr_y+4), val, fill=col, font=font_regular)
        curr_y += 28

    d.text((30, 425), "ЗАКЛЮЧЕНИЕ ДИАГНОСТА: ДВС технически исправен. Допущен к программной калибровке Stage 1.", fill=(22, 101, 52), font=font_bold)

    p = os.path.join(out_dir, "ДБС_Экран_4_Диагностическая_Карта.png")
    im.save(p)
    print("Saved:", p)

# ================= 5. ЭКРАН 5: АРМ КАЛИБРОВЩИКА + КОММЕНТАРИЙ (КАРИМ) =================
def make_screen_5():
    W, H = 850, 500
    im = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(im)
    draw_window_frame(d, W, H, "КИС «Автосервис» — [Сеанс 1.13] АРМ инженера-калибровщика: Прошивка и комментарий (Карим)", "Цех чип-тюнинга / Запись калибровок / Технический отчет")
    
    draw_button(d, 16, 80, 180, 30, "⚡ Записать в ЭБУ", is_primary=True)
    draw_button(d, 206, 80, 160, 30, "💾 Сохранить отчет")
    draw_badge(d, W - 180, 84, "✔ Завершено успешно", "success")

    # Header params
    draw_input(d, 16, 135, 170, 26, "Наряд-заказ (FK)", "НЗ-2026/04-118", is_readonly=True)
    draw_input(d, 196, 135, 200, 26, "Инженер-калибровщик (Карим)", "Гарифуллин К.Р. (Таб. №18)")
    draw_input(d, 406, 135, 190, 26, "Аппаратный ЭБУ", "Bosch MD1CS001 (TC298)")
    draw_input(d, 606, 135, 210, 26, "Программатор / Протокол", "Autotuner Bench Mode OBD-II")

    # Firmware details
    d.rectangle([16, 180, W-16, 280], fill=(255, 255, 255), outline=(226, 232, 240), width=1)
    d.text((26, 190), "ДАННЫЕ ПРОШИВКИ И КОНТРОЛЬНЫХ СУММ (3NF: ORDER_EXECUTION)", fill=(30, 58, 138), font=font_bold)

    draw_input(d, 26, 230, 250, 26, "Файл заводской прошивки (Stock)", "BMW_G30_530d_B57_Stock.bin", is_readonly=True)
    draw_input(d, 286, 230, 250, 26, "Модифицированная версия", "Stage1_DynoTune_v3.2.bin", is_readonly=True)
    draw_input(d, 546, 230, 140, 26, "Чексумма (CRC32)", "0x8F4A2C19 (OK)", is_readonly=True)
    draw_input(d, 696, 230, 120, 26, "Прирост мощности", "+61 л.с. / +100 Нм")

    # Highlighted: TECHNICAL COMMENT (Teacher's requirement for Karim)
    d.rectangle([16, 295, W-16, 485], fill=(254, 252, 232), outline=(234, 179, 8), width=2)
    d.text((26, 305), "⭐ ОБЯЗАТЕЛЬНЫЙ ТЕХНИЧЕСКИЙ КОММЕНТАРИЙ МАСТЕРА-КАЛИБРОВЩИКА (ПО ЗАДАНИЮ):", fill=(161, 98, 7), font=font_bold)

    comment_lines = [
        "«Чтение и запись произведены в режиме Bench Mode через протокол Aurix TC298 без вскрытия блока.",
        "Резервная копия оригинальной прошивки сохранена в архиве КИС (хэш SHA-256: e3b0c44298fc1c149afbf4c8996fb924).",
        "Выполнена оптимизация карт топливоподачи Rail до 2400 бар, целевое давление наддува скорректировано до 1.82 бар.",
        "Убран лимитатор максимальной скорости Vmax, оптимизирован отклик электронной педали акселератора.",
        "Выполнена программная адаптация гидротрансформатора АКПП ZF 8HP75 под увеличенный крутящий момент 720 Нм.",
        "По результатам выходного тест-драйва и логирования на диностенде детонация и аварийные откаты УОЗ отсутствуют.",
        "Температура выхлопа EGT в пике не превышает 740 °C (допустимый предел 820 °C). Рекомендован АИ-98 / ДТ Танеко.»"
    ]
    curr_y = 330
    for l in comment_lines:
        d.text((28, curr_y), l, fill=(15, 23, 42), font=font_regular)
        curr_y += 20

    p = os.path.join(out_dir, "ДБС_Экран_5_АРМ_Калибровщика_Коммент.png")
    im.save(p)
    print("Saved:", p)

# ================= 6. ЭКРАН 6: АКТ ВЫПОЛНЕННЫХ РАБОТ (КАРИМ) =================
def make_screen_6():
    W, H = 850, 480
    im = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(im)
    draw_window_frame(d, W, H, "КИС «Автосервис» — [Сеанс 1.16] Акт выполненных работ № АКТ-2026/04-118 (Документ Карима)", "Финансы / Закрытие наряда / Акт сдачи-приемки и гарантия")
    
    draw_button(d, 16, 80, 160, 30, "🖨 Печать акта", is_primary=True)
    draw_button(d, 186, 80, 170, 30, "💳 Провести оплату")
    draw_badge(d, W - 150, 84, "✔ Подписан", "success")

    # Header Card
    d.rectangle([16, 125, W-16, 215], fill=(255, 255, 255), outline=(226, 232, 240), width=1)
    draw_input(d, 26, 145, 170, 26, "Номер акта (PK)", "АКТ-2026/04-118", is_readonly=True)
    draw_input(d, 206, 145, 150, 26, "Дата закрытия", "26.10.2026 15:45")
    draw_input(d, 366, 145, 220, 26, "Наряд-заказ-основание", "НЗ-2026/04-118", is_readonly=True)
    draw_input(d, 596, 145, 220, 26, "Ответственные мастера", "Гайнанов Д.И. / Гарифуллин К.Р.")

    draw_input(d, 26, 185, 270, 26, "Заказчик", "Смирнов Алексей Викторович")
    draw_input(d, 306, 185, 270, 26, "Автомобиль", "BMW 530d G30 (VIN: WBAJB71090G128941)")
    draw_input(d, 586, 185, 230, 26, "Срок гарантии на ПО", "24 месяца (Безусловная)")

    # Financial settlement
    d.rectangle([16, 230, W-16, 445], fill=(255, 255, 255), outline=(226, 232, 240), width=1)
    d.text((26, 240), "РАСЧЕТ ВЗАИМОРАСЧЕТОВ И ЗАКРЫТИЕ ЗАКАЗА", fill=(30, 58, 138), font=font_bold)

    headers = ["Наименование закрываемой статьи", "Сумма по наряду", "Скидка клиента", "К оплате (руб.)"]
    cols_x = [30, 360, 540, 680]
    d.rectangle([26, 265, W-26, 290], fill=(219, 234, 254))
    for i, h_txt in enumerate(headers):
        d.text((cols_x[i], 270), h_txt, fill=(30, 58, 138), font=font_bold)

    items = [
        ("Компьютерная диагностика и замер ДВС", "8 000.00 руб.", "0.00 руб.", "8 000.00 руб."),
        ("Программная модификация Stage 1 DynoTune (B57)", "34 000.00 руб.", "0.00 руб.", "34 000.00 руб.")
    ]
    curr_y = 295
    for item in items:
        d.line([26, curr_y+24, W-26, curr_y+24], fill=(241, 245, 249), width=1)
        for i, val in enumerate(item):
            d.text((cols_x[i], curr_y+4), val, fill=(15, 23, 42), font=font_regular)
        curr_y += 30

    d.line([26, 375, W-26, 375], fill=(30, 58, 138), width=2)
    d.text((420, 395), "ИТОГО ОПЛАЧЕНО ПОЛНОСТЬЮ:", fill=(30, 58, 138), font=font_bold)
    d.text((680, 390), "42 000.00 руб.", fill=(22, 101, 52), font=font_title)

    d.text((30, 422), "Отметка: Акт подписан Заказчиком без претензий. Автомобиль выдан автовладельцу.", fill=(71, 85, 105), font=font_bold)

    p = os.path.join(out_dir, "ДБС_Экран_6_Акт_Выполненных_Работ.png")
    im.save(p)
    print("Saved:", p)

make_screen_1()
make_screen_2()
make_screen_3()
make_screen_4()
make_screen_5()
make_screen_6()
print("SUCCESS: All 6 Corporate UI Mockups generated successfully!")

# -*- coding: utf-8 -*-
"""
Генератор презентации по теме:
«Особенности архитектуры и сопровождения ИС управления надежностью и ТОиР на базе комплекса «ТРИМ» (TRIM)»
Дисциплина: Эксплуатация автоматизированных систем специального назначения (ЭАССН)
Кафедра АСУ, УУНиТ, 2026 г.

Докладчик: Гайнанов Данил Ильдарович, гр. ЭАС-514С
Преподаватель: к.т.н., доцент Ризванов Константин Анварович

Стиль оформления:
- Светлая профессиональная тема (чистый белый фон, темный высококонтрастный текст)
- Крупные шрифты для четкой читаемости на проекторе в аудитории
- Минималистичная, ненагруженная компоновка (воздух, крупные акценты, краткие тезисы)
"""

import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

current_dir = os.path.dirname(os.path.abspath(__file__))
fig_dir = os.path.join(current_dir, "figures")
pptx_path = os.path.join(current_dir, "EAM_ТРИМ_в_АССН_Презентация.pptx")

prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# ==================== ЦВЕТОВАЯ ПАЛИТРА (СВЕТЛАЯ ИНЖЕНЕРНАЯ ТЕМА) ====================
BG_COLOR = RGBColor(255, 255, 255)          # Чистый белый фон
CARD_BG = RGBColor(248, 250, 252)           # Slate 50 (мягкий светлый фон карточек)
CARD_BORDER = RGBColor(203, 213, 225)       # Slate 300 (четкая граница для проектора)

TEXT_HEAD = RGBColor(15, 23, 42)            # Slate 900 (глубокий темно-синий для заголовков)
TEXT_BODY = RGBColor(30, 41, 59)            # Slate 800 (темный высококонтрастный текст)
TEXT_MUTED = RGBColor(71, 85, 105)          # Slate 600 (поясняющий текст достаточной контрастности)

# Акцентные цвета карточек и плашек (насыщенные для проектора)
BLUE = RGBColor(2, 132, 199)                # Sky 600
NAVY = RGBColor(30, 58, 138)                # Blue 900
EMERALD = RGBColor(16, 149, 106)            # Emerald 600
AMBER = RGBColor(217, 119, 6)               # Amber 600
RED_ACCENT = RGBColor(220, 38, 38)          # Red 600
PURPLE = RGBColor(126, 34, 206)             # Purple 700


def set_slide_background(slide):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR


def add_header(slide, category, title, subtitle=None):
    """Лаконичный крупный заголовок слайда с категорией и пояснением (хорошо читается на проекторе)."""
    # 1. Надзаголовок (категория)
    cat_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.35), Inches(11.9), Inches(0.35))
    tf_cat = cat_box.text_frame
    tf_cat.word_wrap = True
    tf_cat.margin_left = tf_cat.margin_top = tf_cat.margin_right = tf_cat.margin_bottom = 0
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category.upper()
    p_cat.font.name = "Arial"
    p_cat.font.size = Pt(12)
    p_cat.font.bold = True
    p_cat.font.color.rgb = BLUE

    # 2. Основной заголовок слайда (крупный 24 pt)
    title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.68), Inches(11.9), Inches(0.65))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
    p_t = tf_t.paragraphs[0]
    p_t.text = title
    p_t.font.name = "Arial"
    p_t.font.size = Pt(24)
    p_t.font.bold = True
    p_t.font.color.rgb = TEXT_HEAD

    # 3. Подзаголовок / контекст (13.5 pt)
    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.35), Inches(11.9), Inches(0.38))
        tf_s = sub_box.text_frame
        tf_s.word_wrap = True
        tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0
        p_s = tf_s.paragraphs[0]
        p_s.text = subtitle
        p_s.font.name = "Arial"
        p_s.font.size = Pt(13)
        p_s.font.color.rgb = TEXT_MUTED


def add_card(slide, left, top, width, height, title, body_items, accent_color=BLUE, badge_text=None, title_size=15.5, body_size=13):
    """Создает ненагруженную карточку с увеличенным шрифтом для проектора."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = CARD_BORDER
    card.line.width = Pt(1.2)

    # Верхняя цветная акцентная полоска
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left + Inches(0.12), top, width - Inches(0.24), Inches(0.05))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = accent_color
    stripe.line.fill.background()

    tb = slide.shapes.add_textbox(left + Inches(0.18), top + Inches(0.14), width - Inches(0.36), height - Inches(0.22))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    # Заголовок карточки
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = "Arial"
    p0.font.size = Pt(title_size)
    p0.font.bold = True
    p0.font.color.rgb = accent_color
    p0.space_after = Pt(7)

    if badge_text:
        run_b = p0.add_run()
        run_b.text = f"  [{badge_text}]"
        run_b.font.size = Pt(12)
        run_b.font.bold = True
        run_b.font.color.rgb = TEXT_MUTED

    for item in body_items:
        p = tf.add_paragraph()
        if isinstance(item, tuple):
            prefix, text = item
            run_p = p.add_run()
            run_p.text = "• " + prefix + " — "
            run_p.font.name = "Arial"
            run_p.font.bold = True
            run_p.font.color.rgb = TEXT_HEAD
            run_p.font.size = Pt(body_size)

            run_t = p.add_run()
            run_t.text = text
            run_t.font.name = "Arial"
            run_t.font.bold = False
            run_t.font.color.rgb = TEXT_BODY
            run_t.font.size = Pt(body_size)
        else:
            run = p.add_run()
            run.text = "• " + item
            run.font.name = "Arial"
            run.font.size = Pt(body_size)
            run.font.color.rgb = TEXT_BODY
        p.space_after = Pt(6)


# ==================== SLIDE 1: ТИТУЛЬНЫЙ СЛАЙД ====================
def slide_1():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)

    # Верхний бейдж кафедры
    badge = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(0.75), Inches(9.6), Inches(0.48))
    badge.fill.solid()
    badge.fill.fore_color.rgb = RGBColor(241, 245, 249)
    badge.line.color.rgb = RGBColor(203, 213, 225)
    badge.line.width = Pt(1.2)
    tf_b = badge.text_frame
    p_b = tf_b.paragraphs[0]
    p_b.text = "УФИМСКИЙ УНИВЕРСИТЕТ НАУКИ И ТЕХНОЛОГИЙ • КАФЕДРА АСУ • 2026"
    p_b.font.name = "Arial"
    p_b.font.size = Pt(12)
    p_b.font.bold = True
    p_b.font.color.rgb = BLUE

    # Заголовок доклада (крупный 31 pt)
    title_box = s.shapes.add_textbox(Inches(0.7), Inches(1.55), Inches(11.9), Inches(2.4))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    p1 = tf_t.paragraphs[0]
    p1.text = "Архитектура и эксплуатация EAM-комплекса «ТРИМ»\nна критических объектах специального назначения"
    p1.font.name = "Arial"
    p1.font.size = Pt(30)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_HEAD
    p1.space_after = Pt(14)

    p2 = tf_t.add_paragraph()
    p2.text = "Дисциплина: «Эксплуатация автоматизированных систем специального назначения» (ЭАССН)"
    p2.font.name = "Arial"
    p2.font.size = Pt(16)
    p2.font.bold = True
    p2.font.color.rgb = AMBER

    # Карточки докладчика и руководителя с крупным шрифтом 14 pt
    add_card(s, Inches(0.7), Inches(4.55), Inches(5.8), Inches(2.35), "ДОКЛАДЧИК", [
        ("Студент", "Гайнанов Данил Ильдарович"),
        ("Группа", "ЭАС-514С (5 курс)"),
        ("Специальность", "АС специального назначения"),
    ], accent_color=BLUE, title_size=17, body_size=13.5)

    add_card(s, Inches(6.8), Inches(4.55), Inches(5.8), Inches(2.35), "НАУЧНЫЙ РУКОВОДИТЕЛЬ", [
        ("Преподаватель", "Ризванов Константин Анварович"),
        ("Ученая степень", "Кандидат технических наук, доцент"),
        ("Кафедра", "Автоматизированных систем управления (АСУ)"),
    ], accent_color=EMERALD, title_size=17, body_size=13.5)


# ==================== SLIDE 2: СПЕЦИФИКА СПЕЦОБЪЕКТОВ ====================
def slide_2():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "СПЕЦИФИКА ОБЪЕКТОВ", "Условия эксплуатации критических автономных спецобъектов",
               "Атомный ледокольный флот («Арктика», «Сибирь») и плавучие атомные станции (ПАТЭС)")

    w = Inches(2.8)
    gap = Inches(0.23)
    top = Inches(1.88)
    h = Inches(5.05)

    add_card(s, Inches(0.7), top, w, h, "1. АВТОНОМНОСТЬ", [
        ("Срок похода", "4–9 месяцев плавания без баз снабжения."),
        ("Радиомолчание", "Недели плавания без внешней связи."),
        ("Автономия ИС", "100% самодостаточный бортовой контур.")
    ], accent_color=BLUE, title_size=15, body_size=13)

    add_card(s, Inches(0.7) + w + gap, top, w, h, "2. ЦЕНА ОТКАЗА", [
        ("Ядерная безопасность", "1-я категория, сбой реактора недопустим."),
        ("Ледовый плен", "Потеря хода угрожает всему каравану судов."),
        ("Упреждение сбоя", "Прогноз отказа за сотни часов до поломки.")
    ], accent_color=RED_ACCENT, title_size=15, body_size=13)

    add_card(s, Inches(0.7) + (w + gap)*2, top, w, h, "3. ДЕФИЦИТ ЗИП", [
        ("Изоляция склада", "Доставка запчастей в Арктику невозможна."),
        ("Лимит объема", "Судовой ЗИП строго ограничен трюмами."),
        ("Точный расчет", "Предиктивный заказ ЗИП до выхода в море.")
    ], accent_color=AMBER, title_size=15, body_size=13)

    add_card(s, Inches(0.7) + (w + gap)*3, top, w, h, "4. ЭКСТРИМ СРЕДЫ", [
        ("Вибрация и качка", "Удары о торосы, солевой морской туман."),
        ("Блэкауты сети", "Просадки напряжения при тяжелых пусках."),
        ("Защита серверов", "Морское вибро- и влагостойкое исполнение.")
    ], accent_color=EMERALD, title_size=15, body_size=13)


# ==================== SLIDE 3: ОБЗОР ПЛАТФОРМЫ ТРИМ ====================
def slide_3():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ПЛАТФОРМА ИС", "Программный комплекс «ТРИМ» (TRIM) — EAM для спецобъектов",
               "Отечественная система Enterprise Asset Management от НПП «СпецТек» (г. Санкт-Петербург)")

    top = Inches(1.88)
    w = Inches(3.78)
    gap = Inches(0.28)
    h = Inches(5.05)

    add_card(s, Inches(0.7), top, w, h, "ПРОФИЛЬ И ВНЕДРЕНИЯ", [
        ("Спец-назначение", "Создан для тяжелой морской и атомной техники."),
        ("Атомный флот", "Ледоколы «Арктика», «Сибирь», ПАТЭС «Ломоносов»."),
        ("Сертификация", "Одобрен Морским (РМРС) и Речным регистром."),
        ("Суверенитет", "100% независим от зарубежных ERP-систем.")
    ], accent_color=BLUE, title_size=16, body_size=13.5)

    add_card(s, Inches(0.7) + w + gap, top, w, h, "МОДУЛИ СИСТЕМЫ", [
        ("TRIM-Maintenance", "Управление ТОиР, электронные наряды и регламенты."),
        ("TRIM-Warehouse", "Партионный учет ЗИП, сроки годности смазок и РТИ."),
        ("TRIM-Reliability", "Расчет надежности, вероятности отказов и FMECA."),
        ("TRIM-Doc", "Интеграция с интерактивной документацией (ИЭТР).")
    ], accent_color=EMERALD, title_size=16, body_size=13.5)

    add_card(s, Inches(0.7) + (w + gap)*2, top, w, h, "АРХИТЕКТУРА", [
        ("Иерархия", "Трехуровневая модель: Борт <—> Берег <—> Штаб."),
        ("Узкие каналы", "Оптимизация пакетов под связь от 9.6 кбит/с."),
        ("Доверенная ОС", "Работа в среде защищенной Astra Linux SE."),
        ("Аттестация КИИ", "Высшая категория критической инфраструктуры.")
    ], accent_color=PURPLE, title_size=16, body_size=13.5)


# ==================== SLIDE 4: АСИНХРОННАЯ РЕПЛИКАЦИЯ ====================
def slide_4():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "СЕТЕВАЯ АРХИТЕКТУРА", "Асинхронная пакетная репликация ТРИМ: «Борт <—> Берег»",
               "Гарантированная синхронизация баз данных в условиях узкополосной спутниковой связи")

    img_path = os.path.join(fig_dir, "diagram_trim_replication_architecture.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.7), Inches(1.85), width=Inches(7.9))

    add_card(s, Inches(8.8), Inches(1.85), Inches(3.8), Inches(2.5), "ПАКЕТНЫЙ МЕХАНИЗМ", [
        ("Delta-пакеты", "Передаются только измененные записи (XML)."),
        ("Очередь Outbox", "Автономное накопление данных неделями."),
        ("Защита ГОСТ", "Подписание и шифрование перед отправкой.")
    ], accent_color=BLUE, title_size=15.5, body_size=13)

    add_card(s, Inches(8.8), Inches(4.5), Inches(3.8), Inches(2.45), "БЕРЕГОВОЙ ЦЕНТР", [
        ("Сеансовая связь", "Синхронизация за секунды спутникового контакта."),
        ("Приоритет борта", "При коллизиях приоритет имеют судовые записи."),
        ("Мониторинг флота", "Единая картина техсостояния всех атомоходов.")
    ], accent_color=EMERALD, title_size=15.5, body_size=13)


# ==================== SLIDE 5: КОНТУР НАДЕЖНОСТИ RCM ====================
def slide_5():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "МЕТОДОЛОГИЯ НАДЕЖНОСТИ", "Контур RCM (Reliability-Centered Maintenance) в ТРИМ",
               "Переход от календарного ППР к обслуживанию по фактическому техническому состоянию")

    img_path = os.path.join(fig_dir, "diagram_trim_rcm_workflow.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.7), Inches(1.85), width=Inches(7.9))

    add_card(s, Inches(8.8), Inches(1.85), Inches(3.8), Inches(2.5), "МАТРИЦА FMECA", [
        ("Индекс RPN", "Критичность = Вероятность × Тяжесть × Обнаружение."),
        ("Фокус на рисках", "Приоритет узлам с наивысшим риском отказа."),
        ("P-F интервал", "Выявление дефекта задолго до наступления сбоя.")
    ], accent_color=AMBER, badge_text="RPN", title_size=15.5, body_size=13)

    add_card(s, Inches(8.8), Inches(4.5), Inches(3.8), Inches(2.45), "ЭФФЕКТ ДЛЯ ОБЪЕКТА", [
        ("Экономия ЗИП", "Снижение расхода дорогостоящих узлов на 25–35%."),
        ("Готовность судна", "Рост коэффициента готовности (КТГ) до 0.98–0.99."),
        ("Отказ от лишнего ППР", "Исключение опасных ненужных переборок.")
    ], accent_color=EMERALD, badge_text="КТГ > 0.98", title_size=15.5, body_size=13)


# ==================== SLIDE 6: ЭЛЕКТРОННЫЙ ПАСПОРТ ====================
def slide_6():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ЭКСПЛУАТАЦИОННЫЙ УЧЕТ", "Структура электронного паспорта и формуляра оборудования",
               "Цифровой двойник технического состояния спецобъекта по ГОСТ 27.002-2015")

    img_path = os.path.join(fig_dir, "diagram_trim_electronic_logbook.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.7), Inches(1.85), width=Inches(7.9))

    add_card(s, Inches(8.8), Inches(1.85), Inches(3.8), Inches(2.5), "ИЕРАРХИЯ ОБЪЕКТА", [
        ("10 000+ узлов", "Декомпозиция: судно -> установка -> узел -> деталь."),
        ("Кодирование KKS", "Однозначный код каждого насоса и клапана."),
        ("История цикла", "Полный след: монтаж -> наработка -> ремонты.")
    ], accent_color=BLUE, badge_text="10 000+ узлов", title_size=15.5, body_size=13)

    add_card(s, Inches(8.8), Inches(4.5), Inches(3.8), Inches(2.45), "ЮРИДИЧЕСКАЯ СИЛА", [
        ("Личная ЭЦП", "Каждая операция подписывается механиком."),
        ("Неизменяемость", "Защита от искажения наработки задним числом."),
        ("Морской Регистр", "Признание формуляров инспекторами РМРС.")
    ], accent_color=PURPLE, badge_text="ГОСТ ЭЦП", title_size=15.5, body_size=13)


# ==================== SLIDE 7: ИНТЕГРАЦИЯ С SCADA И ДИАГНОСТИКОЙ ====================
def slide_7():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ИНТЕГРАЦИЯ ДАННЫХ", "Интеграция ТРИМ с бортовыми АСУ ТП и вибродиагностикой",
               "Автоматический сбор параметров телеметрии без участия вахтенного персонала")

    top = Inches(1.88)
    w = Inches(3.78)
    gap = Inches(0.28)
    h = Inches(5.05)

    add_card(s, Inches(0.7), top, w, h, "ИСТОЧНИКИ ТЕЛЕМЕТРИИ", [
        ("Судовая SCADA", "Авто-съем наработки в моточасах от контроллеров ГЭУ."),
        ("Вибромониторинг", "Прямой импорт спектров подшипников турбин."),
        ("Тепловизоры", "Температурные тренды генераторов и щитов."),
        ("Протоколы", "Безопасный сбор через OPC UA и Modbus TCP.")
    ], accent_color=BLUE, title_size=16, body_size=13.5)

    add_card(s, Inches(0.7) + w + gap, top, w, h, "ПРЕДИКТИВНЫЕ ТРИГГЕРЫ", [
        ("Контроль ISO 10816", "Авто-детекция выхода вибрации за уставки тревог."),
        ("Авто-наряд на ТО", "Система сама выпускает наряд на виброналадку."),
        ("Резерв деталей", "Авто-проверка наличия подшипников в кладовой."),
        ("Ранняя локализация", "Устранение дефекта на стадии микротрещины.")
    ], accent_color=AMBER, badge_text="ISO 10816", title_size=16, body_size=13.5)

    add_card(s, Inches(0.7) + (w + gap)*2, top, w, h, "РАЗГРУЗКА ЭКИПАЖА", [
        ("Без ручного ввода", "Механики освобождены от рутинной переписи журналов."),
        ("100% достоверность", "Исключен человеческий фактор сокрытия наработки."),
        ("Точный график", "Календарь ТО пересчитывается по реальной нагрузке.")
    ], accent_color=EMERALD, title_size=16, body_size=13.5)


# ==================== SLIDE 8: ЖИВУЧЕСТЬ СЕРВЕРА ====================
def slide_8():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ЖИВУЧЕСТЬ СИСТЕМЫ", "Обеспечение живучести локального сервера ТРИМ в море",
               "Аппаратно-программная защита базы данных при блэкаутах и авариях судна")

    img_path = os.path.join(fig_dir, "diagram_trim_fault_tolerance.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.7), Inches(1.85), width=Inches(7.9))

    add_card(s, Inches(8.8), Inches(1.85), Inches(3.8), Inches(2.5), "ЗАЩИТА ПИТАНИЯ", [
        ("Морские On-Line ИБП", "Фильтрация бросков судовой сети 380 В."),
        ("Сброс WAL < 2 сек", "Аварийная фиксация транзакций СУБД на SSD."),
        ("Стойкие корпуса", "Вибро- и влагозащита серверных стоек.")
    ], accent_color=RED_ACCENT, badge_text="< 2 сек", title_size=15.5, body_size=13)

    add_card(s, Inches(8.8), Inches(4.5), Inches(3.8), Inches(2.45), "ДУБЛИРОВАНИЕ СЕРВЕРОВ", [
        ("Разнесение бортов", "Серверы разнесены по разным отсекам судна."),
        ("Синхронная репликация", "Нулевая потеря данных при аварии (RPO = 0)."),
        ("Авто-переключение", "Мгновенная смена ролей при затоплении отсека.")
    ], accent_color=EMERALD, badge_text="RPO = 0", title_size=15.5, body_size=13)


# ==================== SLIDE 9: СКЛАД ЗИП В ПОХОДЕ ====================
def slide_9():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "МАТЕРИАЛЬНОЕ СНАБЖЕНИЕ", "Управление складом ЗИП в автономном арктическом походе",
               "Математическое прогнозирование достаточности запаса запчастей на многомесячный рейс")

    top = Inches(1.88)
    w = Inches(3.78)
    gap = Inches(0.28)
    h = Inches(5.05)

    add_card(s, Inches(0.7), top, w, h, "СПЕЦИФИКА СУДОВОГО СКЛАДА", [
        ("Неснижаемый запас", "Критический ЗИП по обязательным нормам РМРС."),
        ("Сложная топология", "Учет размещения деталей в десятках выгородок и трюмов."),
        ("Условия хранения", "Контроль температуры для специальных РТИ и смазок.")
    ], accent_color=BLUE, title_size=16, body_size=13.5)

    add_card(s, Inches(0.7) + w + gap, top, w, h, "МАТЕМАТИКА РАСЧЕТА ЗИП", [
        ("Закон Пуассона", "Расчет вероятности отказов по интенсивности λ."),
        ("План рейса", "Формирование заказа под ледовый маршрут."),
        ("Сроки годности", "Контроль предельных сроков консервации узлов.")
    ], accent_color=AMBER, badge_text="Модель λ", title_size=16, body_size=13.5)

    add_card(s, Inches(0.7) + (w + gap)*2, top, w, h, "АДРЕСНЫЙ УЧЕТ И ВЫДАЧА", [
        ("Штрихкодирование", "Идентификация взрывозащищенными терминалами."),
        ("Поиск за секунды", "Механик сразу видит отсек, стеллаж и полку."),
        ("Списание по наряду", "Деталь списывается строго при закрытии наряда.")
    ], accent_color=EMERALD, title_size=16, body_size=13.5)


# ==================== SLIDE 10: ИМПОРТОЗАМЕЩЕНИЕ И БЕЗОПАСНОСТЬ ====================
def slide_10():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ДОВЕРЕННЫЙ СТЕК", "Переход комплекса ТРИМ на отечественный защищенный стек",
               "Функционирование в среде Astra Linux Special Edition и защищенной СУБД Postgres Pro")

    top = Inches(1.88)
    w = Inches(2.8)
    gap = Inches(0.23)
    h = Inches(5.05)

    add_card(s, Inches(0.7), top, w, h, "ОС ASTRA LINUX SE", [
        ("Релиз «Смоленск»", "Защищенная среда спецназначения."),
        ("Мандатный доступ", "Изоляция служебных контуров судна."),
        ("Замкнутая среда", "Запрет неавторизованного кода.")
    ], accent_color=BLUE, title_size=15, body_size=13)

    add_card(s, Inches(0.7) + w + gap, top, w, h, "СУБД POSTGRES PRO", [
        ("Отказ от MS SQL", "Сертифицированная СУБД РФ."),
        ("Сжатие CFS", "Экономия ресурса судовых SSD."),
        ("Отказоустойчивость", "Кластерная архитектура высокой готовности.")
    ], accent_color=AMBER, title_size=15, body_size=13)

    add_card(s, Inches(0.7) + (w + gap)*2, top, w, h, "ВЕБ-КЛИЕНТ ТРИМ", [
        ("«Нулевой клиент»", "Работа через Chromium-ГОСТ."),
        ("Планшеты у узлов", "Работа инженеров по защищенному Wi-Fi."),
        ("Кроссплатформенность", "Единый доступ без установки стороннего ПО.")
    ], accent_color=PURPLE, title_size=15, body_size=13)

    add_card(s, Inches(0.7) + (w + gap)*3, top, w, h, "СТАТУС И ЛИЦЕНЗИИ", [
        ("Реестр Минцифры", "Запись № 1245 в реестре ПО РФ."),
        ("Свидетельство РМРС", "Морской регистр судоходства."),
        ("Объекты КИИ-1", "Высшая категория критической инфраструктуры.")
    ], accent_color=EMERALD, title_size=15, body_size=13)


# ==================== SLIDE 11: СРАВНИТЕЛЬНАЯ ТАБЛИЦА ====================
def slide_11():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "СРАВНИТЕЛЬНЫЙ АНАЛИЗ", "Сравнительная матрица: Традиционный ППР против EAM «ТРИМ»",
               "Количественные показатели эксплуатационной эффективности внедрения комплекса")

    rows = 6
    cols = 4
    left = Inches(0.7)
    top = Inches(1.88)
    width = Inches(11.9)
    height = Inches(5.05)

    table_shape = s.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    table.columns[0].width = Inches(3.0)
    table.columns[1].width = Inches(3.0)
    table.columns[2].width = Inches(3.3)
    table.columns[3].width = Inches(2.6)

    headers = ["Эксплуатационный фактор", "Традиционный плановый ремонт (ППР)", "EAM-комплекс ТРИМ (RCM)", "Выигрыш для спецобъекта"]
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(241, 245, 249)
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.name = "Arial"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = BLUE
        p.alignment = PP_ALIGN.CENTER

    data = [
        ("Стратегия обслуживания", "Жесткий график по календарным часам", "По фактическому техсостоянию и наработке", "Оптимизация ремонтов"),
        ("Внезапные отказы агрегатов", "Высокий риск (скрытые дефекты не видны)", "Минимальный (предиктивное обнаружение)", "Снижение отказов на 70%"),
        ("Расход судового ЗИП", "Избыточная плановая замена исправных узлов", "Замена строго по достижению лимита износа", "Экономия ЗИП 25–35%"),
        ("Учет наработки оборудования", "Ручные вахтенные журналы, риск искажений", "Автоматический импорт из АСУ ТП, ЭЦП", "100% достоверность"),
        ("Синхронизация с берегом", "Бумажные отчеты по возвращении в порт", "Асинхронные XML-пакеты по спутнику", "Контроль флота 24/7")
    ]

    for i, row in enumerate(data):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255) if i % 2 == 0 else RGBColor(248, 250, 252)
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Arial"
            p.font.size = Pt(12.5)
            if j == 0:
                p.font.bold = True
                p.font.color.rgb = TEXT_HEAD
            elif j == 3:
                p.font.bold = True
                p.font.color.rgb = EMERALD
            else:
                p.font.color.rgb = TEXT_BODY


# ==================== SLIDE 12: ЗАКЛЮЧЕНИЕ ====================
def slide_12():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ЗАКЛЮЧЕНИЕ", "Ключевые результаты и эксплуатационные рекомендации",
               "ТРИМ как эталон надежного информационного сопровождения объектов специального назначения")

    top = Inches(1.88)
    w = Inches(5.8)
    gap = Inches(0.3)
    h = Inches(2.4)

    add_card(s, Inches(0.7), top, w, h, "1. АВТОНОМНОСТЬ И НАДЕЖНОСТЬ", [
        ("Полная независимость", "Доказана надежность в арктических походах при радиомолчании."),
        ("Пакетная репликация", "Эффективная передача данных по узкополосным каналам от 9.6 кбит/с.")
    ], accent_color=BLUE, title_size=16, body_size=13.5)

    add_card(s, Inches(0.7) + w + gap, top, w, h, "2. НАУЧНЫЙ ПОДХОД RCM", [
        ("Управление рисками", "Анализ FMECA прогнозирует отказ за сотни часов до аварии."),
        ("Рост готовности", "Коэффициент готовности энергоустановок (КТГ) достигает 0.98–0.99.")
    ], accent_color=AMBER, badge_text="КТГ > 0.98", title_size=16, body_size=13.5)

    add_card(s, Inches(0.7), top + h + Inches(0.25), w, h, "3. СКВОЗНАЯ ИНТЕГРАЦИЯ", [
        ("Связь с АСУ ТП", "Бесшовная передача вибрации и моточасов исключает ошибки персонала."),
        ("Юридический след", "Электронный формуляр с личной ЭЦП вахтенных механиков.")
    ], accent_color=EMERALD, title_size=16, body_size=13.5)

    add_card(s, Inches(0.7) + w + gap, top + h + Inches(0.25), w, h, "4. ТЕХНОЛОГИЧЕСКИЙ СУВЕРЕНИТЕТ", [
        ("Доверенная среда", "Успешная эксплуатация на базе Astra Linux SE и Postgres Pro."),
        ("Стандарт отрасли", "Многолетний опыт ледоколов «Атомфлота» — эталон для спецсистем РФ.")
    ], accent_color=PURPLE, badge_text="КИИ РФ", title_size=16, body_size=13.5)


# ==================== СБОРКА ПРЕЗЕНТАЦИИ ====================
slide_1()
slide_2()
slide_3()
slide_4()
slide_5()
slide_6()
slide_7()
slide_8()
slide_9()
slide_10()
slide_11()
slide_12()

prs.save(pptx_path)
print("TRIM Presentation saved successfully with enlarged projector fonts at:", pptx_path)

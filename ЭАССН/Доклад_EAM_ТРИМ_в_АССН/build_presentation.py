# -*- coding: utf-8 -*-
import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

out_dir = r"c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\Доклад_EAM_ТРИМ_в_АССН"
fig_dir = os.path.join(out_dir, "figures")
pptx_path = os.path.join(out_dir, "EAM_ТРИМ_в_АССН_Презентация.pptx")

prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# Colors (Executive Dark Tech Theme)
BG_COLOR = RGBColor(15, 23, 42)       # Slate 900
CARD_BG = RGBColor(30, 41, 59)        # Slate 800
CARD_BORDER = RGBColor(51, 65, 85)    # Slate 700
CYAN = RGBColor(56, 189, 248)         # Sky 400
EMERALD = RGBColor(52, 211, 153)      # Emerald 400
AMBER = RGBColor(251, 191, 36)        # Amber 400
PURPLE = RGBColor(192, 132, 252)      # Purple 400
WHITE = RGBColor(248, 250, 252)       # Slate 50
TEXT_MUTED = RGBColor(148, 163, 184)  # Slate 400
RED_ACCENT = RGBColor(248, 113, 113)  # Red 400

def set_slide_background(slide):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

def add_header(slide, category, title, subtitle=None):
    cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.35))
    tf_cat = cat_box.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category.upper()
    p_cat.font.name = "Arial"
    p_cat.font.size = Pt(11)
    p_cat.font.bold = True
    p_cat.font.color.rgb = CYAN

    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(11.7), Inches(0.65))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.text = title
    p_t.font.name = "Arial"
    p_t.font.size = Pt(21)
    p_t.font.bold = True
    p_t.font.color.rgb = WHITE

    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(11.7), Inches(0.4))
        tf_s = sub_box.text_frame
        tf_s.word_wrap = True
        p_s = tf_s.paragraphs[0]
        p_s.text = subtitle
        p_s.font.name = "Arial"
        p_s.font.size = Pt(12)
        p_s.font.color.rgb = TEXT_MUTED

def add_card(slide, left, top, width, height, title, body_paragraphs, accent_color=CYAN, title_size=15):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = CARD_BORDER
    card.line.width = Pt(1.2)

    tb = slide.shapes.add_textbox(left + Inches(0.18), top + Inches(0.15), width - Inches(0.36), height - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True

    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = "Arial"
    p0.font.size = Pt(title_size)
    p0.font.bold = True
    p0.font.color.rgb = accent_color
    p0.space_after = Pt(8)

    for item in body_paragraphs:
        p = tf.add_paragraph()
        if isinstance(item, tuple):
            prefix, text = item
            run1 = p.add_run()
            run1.text = prefix + ": "
            run1.font.bold = True
            run1.font.color.rgb = WHITE
            run1.font.size = Pt(11)

            run2 = p.add_run()
            run2.text = text
            run2.font.bold = False
            run2.font.color.rgb = TEXT_MUTED
            run2.font.size = Pt(11)
        else:
            run = p.add_run()
            run.text = item
            run.font.size = Pt(11)
            run.font.color.rgb = TEXT_MUTED
        p.space_after = Pt(5)

# ==================== SLIDE 1: TITLE ====================
def slide_1():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)

    badge = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.9), Inches(8.5), Inches(0.45))
    badge.fill.solid()
    badge.fill.fore_color.rgb = CARD_BG
    badge.line.color.rgb = CARD_BORDER
    badge.line.width = Pt(1)
    tf_b = badge.text_frame
    p_b = tf_b.paragraphs[0]
    p_b.text = "УФИМСКИЙ УНИВЕРСИТЕТ НАУКИ И ТЕХНОЛОГИЙ • КАФЕДРА АСУ • 2026"
    p_b.font.name = "Arial"
    p_b.font.size = Pt(10.5)
    p_b.font.bold = True
    p_b.font.color.rgb = CYAN

    title_box = s.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.7), Inches(2.2))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    p1 = tf_t.paragraphs[0]
    p1.text = "Особенности архитектуры и сопровождения ИС\nуправления надежностью и ТОиР на базе комплекса «ТРИМ»"
    p1.font.name = "Arial"
    p1.font.size = Pt(29)
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    p1.space_after = Pt(14)

    p2 = tf_t.add_paragraph()
    p2.text = "Дисциплина: «Эксплуатация автоматизированных систем специального назначения» (ЭАССН)"
    p2.font.name = "Arial"
    p2.font.size = Pt(15)
    p2.font.color.rgb = AMBER

    add_card(s, Inches(0.8), Inches(4.7), Inches(5.6), Inches(2.1), "ДОКЛАДЧИК", [
        ("Студент", "Гайнанов Данил Ильдарович"),
        ("Группа", "ЭАС-514С (5 курс)"),
        ("Специальность", "Эксплуатация автоматизированных систем специального назначения"),
    ], accent_color=CYAN)

    add_card(s, Inches(6.8), Inches(4.7), Inches(5.7), Inches(2.1), "НАУЧНЫЙ РУКОВОДИТЕЛЬ", [
        ("Преподаватель", "Ризванов Константин Анварович"),
        ("Ученая степень", "Кандидат технических наук, доцент"),
        ("Кафедра", "Автоматизированных систем управления (АСУ)"),
    ], accent_color=EMERALD)

# ==================== SLIDE 2: SPECIAL OPERATIONAL CONTEXT ====================
def slide_2():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "СПЕЦИФИКА ОБЪЕКТОВ", "Условия эксплуатации критических автономных спецобъектов",
               "Атомный ледокольный флот, ПАТЭС и морские платформы диктуют уникальные требования к ИС")

    w = Inches(2.75)
    gap = Inches(0.23)
    top = Inches(2.0)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "1. ДЛИТЕЛЬНАЯ АВТОНОМНОСТЬ", [
        ("Автономный рейс", "Объекты находятся в море от 4 до 9 месяцев без захода в порт."),
        ("Отсутствие связи", "Недели плавания в условиях арктического радиомолчания."),
        ("Автономия ИС", "Локальный бортовой контур ИС обязан быть 100% самодостаточным.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "2. ЦЕНА ОТКАЗА АГРЕГАТОВ", [
        ("Ядерная безопасность", "Отказ реакторной установки или ГЭУ грозит катастрофой 1-й категории."),
        ("Ледовый плен", "Потеря хода во льдах толщиной 3 метра ставит под угрозу весь караван судов."),
        ("Недопустимость сбоя", "ИС должна предупреждать отказ за сотни часов до инцидента.")
    ], accent_color=RED_ACCENT)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "3. ДЕФИЦИТ ЗИП", [
        ("Изолированный склад", "Доставить сломанный насос или датчик в Карское море невозможно."),
        ("Лимит номенклатуры", "Судовой состав ЗИП строго ограничен вместимостью трюмов."),
        ("Точный прогноз", "ИС должна предиктивно рассчитывать расход запчастей до выхода в море.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8) + (w + gap)*3, top, w, h, "4. АГРЕССИВНАЯ СРЕДА", [
        ("Вибрация и качка", "Постоянные перегрузки, удары о торосы, солевой туман."),
        ("Блэкауты (Blackouts)", "Резкие просадки судовой сети при пуске тяжелых компрессоров."),
        ("Аппаратная защита", "Серверы в морском защищенном исполнении с защитой от сбоев питания.")
    ], accent_color=EMERALD)

# ==================== SLIDE 3: TRIM OVERVIEW ====================
def slide_3():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ПЛАТФОРМА ИС", "Программный комплекс «ТРИМ» (TRIM) — EAM для спецобъектов",
               "Отечественная система Enterprise Asset Management от НПП «СпецТек» (г. Санкт-Петербург)")

    top = Inches(2.0)
    w = Inches(3.7)
    gap = Inches(0.3)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "ПРОФИЛЬНАЯ СПЕЦИАЛИЗАЦИЯ", [
        ("Не универсальная ERP", "Создана исключительно для тяжелой морской, атомной и военной техники."),
        ("Флагманские внедрения", "Ледоколы ФГУП «Атомфлот» («Арктика», «Сибирь»), ПАТЭС «Академик Ломоносов», объекты Минобороны."),
        ("Сертификация", "Одобрена Российским морским регистром судоходства (РМРС) и Речным регистром (РКО).")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "ФУНКЦИОНАЛЬНЫЕ МОДУЛИ", [
        ("TRIM-Maintenance", "Управление ТОиР, электронные технологические карты и планировщик работ."),
        ("TRIM-Warehouse", "Учет судового ЗИП, партионный учет, контроль сроков годности уплотнений и смазок."),
        ("TRIM-Reliability (RCM)", "Расчет надежности, вероятности безотказной работы и матриц рисков FMECA."),
        ("TRIM-Doc", "Интеграция с электронной эксплуатационной документацией изделия (ИЭТР).")
    ], accent_color=EMERALD)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "АРХИТЕКТУРНЫЙ БАЗИС", [
        ("Многоуровневость", "Иерархия: Судовые локальные СУБД <—> Береговой кластер DWH."),
        ("Независимость от ОС", "Поддержка защищенных дистрибутивов Astra Linux Special Edition."),
        ("Низкий сетевой оверхед", "Пакетная передача данных, оптимизированная под каналы со скоростью от 9.6 кбит/с.")
    ], accent_color=PURPLE)

# ==================== SLIDE 4: REPLICATION ARCHITECTURE ====================
def slide_4():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "СЕТЕВАЯ АРХИТЕКТУРА", "Асинхронная пакетная репликация ТРИМ: «Борт <—> Берег»",
               "Гарантированная синхронизация баз данных в условиях узкополосной спутниковой связи")

    img_path = os.path.join(fig_dir, "diagram_trim_replication_architecture.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "ПАКЕТНЫЙ МЕХАНИЗМ", [
        ("Delta-пакеты", "Передаются только измененные записи в формате сжатого бинарного XML."),
        ("Очередь Outbox", "При отсутствии связи пакеты буферизируются на борту неделями."),
    ], accent_color=CYAN)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "БЕРЕГОВОЙ ЦЕНТР", [
        ("Сводная аналитика", "Береговая служба видит техническое состояние всех атомоходов в реальном времени."),
        ("Разрешение коллизий", "Приоритет отдается фактическим бортовым записям наработки."),
    ], accent_color=EMERALD)

# ==================== SLIDE 5: RCM WORKFLOW ====================
def slide_5():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "МЕТОДОЛОГИЯ НАДЕЖНОСТИ", "Контур RCM (Reliability-Centered Maintenance) в ТРИМ",
               "Переход от устаревшего жесткого регламента часов к обслуживанию по техническому состоянию")

    img_path = os.path.join(fig_dir, "diagram_trim_rcm_workflow.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "МАТРИЦА FMECA", [
        ("Индекс критичности", "Оценка риска: RPN = Вероятность × Тяжесть последствий × Обнаруживаемость."),
        ("Приоритет агрегатов", "Концентрация внимания экипажа на оборудовании с наивысшим риском отказа."),
    ], accent_color=AMBER)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "ЭКОНОМИЧЕСКИЙ ЭФФЕКТ", [
        ("Экономия ЗИП", "Снижение расхода дорогостоящих запчастей на 25–35%."),
        ("Коэффициент готовности", "Рост коэффициента технической готовности (КТГ) судна до 0.98–0.99."),
    ], accent_color=EMERALD)

# ==================== SLIDE 6: ELECTRONIC LOGBOOK ====================
def slide_6():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ЭКСПЛУАТАЦИОННЫЙ УЧЕТ", "Структура электронного паспорта и формуляра оборудования",
               "Цифровой двойник технического состояния спецобъекта по стандартам ГОСТ 27.002-2015")

    img_path = os.path.join(fig_dir, "diagram_trim_electronic_logbook.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "ИЕРАРХИЯ ОБЪЕКТА", [
        ("До 10 000 позиций", "Полная декомпозиция: судно -> установка -> подсистема -> узел -> деталь."),
        ("Уникальный KKS-код", "Маркировка каждого клапана и насоса по международной системе KKS/RDS-PP."),
    ], accent_color=CYAN)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "ЮРИДИЧЕСКАЯ СИЛА", [
        ("Электронная подпись", "Каждая запись механика о ремонте подписывается личной ЭЦП."),
        ("Защита от фальсификаций", "Невозможно задним числом исказить данные о наработке узла."),
    ], accent_color=PURPLE)

# ==================== SLIDE 7: SCADA & TELEMETRY INTEGRATION ====================
def slide_7():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ИНТЕГРАЦИЯ ДАННЫХ", "Интеграция ТРИМ с бортовыми АСУ ТП и вибродиагностикой",
               "Автоматический сбор параметров телеметрии без участия вахтенного персонала")

    top = Inches(2.0)
    w = Inches(3.7)
    gap = Inches(0.3)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "ИСТОЧНИКИ ТЕЛЕМЕТРИИ", [
        ("Судовая SCADA", "Автоматическая передача счетчиков наработки в моточасах от контроллеров ГЭУ."),
        ("Вибромониторинг", "Прямой импорт спектров вибрации подшипников главных турбогенераторов."),
        ("Тепловизионный контроль", "Температурные тренды электрических машин и распределительных щитов."),
        ("Протоколы обмена", "Использование промышленных шин OPC UA и Modbus TCP в защищенном контуре.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "ПРЕДИКТИВНЫЕ ТРИГГЕРЫ", [
        ("Авто-детекция аномалий", "Фиксация выхода вибрации за пороговое предупредительное поле (ISO 10816)."),
        ("Генерация предписания", "ТРИМ автоматически создает дефектную ведомость и наряд на виброналадку."),
        ("Резервирование подшипника", "Система автоматически проверяет наличие ремкомплекта в судовой кладовой."),
        ("Предотвращение заклинивания", "Отказ локализуется на ранней стадии зарождения микротрещины.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "РАЗГРУЗКА ЭКИПАЖА", [
        ("Исключение ручного ввода", "Механики освобождены от ежедневного переписывания показаний счетчиков."),
        ("Устранение искажений", "Исключен человеческий фактор занижения наработки перед проверкой."),
        ("Точность планирования", "Календарные графики ТО автоматически пересчитываются по реальной нагрузке судна.")
    ], accent_color=EMERALD)

# ==================== SLIDE 8: FAULT TOLERANCE ====================
def slide_8():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ЖИВУЧЕСТЬ СИСТЕМЫ", "Обеспечение живучести локального сервера ТРИМ в море",
               "Аппаратно-программные решения для сохранения целостности БД при блэкаутах и авариях")

    img_path = os.path.join(fig_dir, "diagram_trim_fault_tolerance.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "ЗАЩИТА ПИТАНИЯ", [
        ("Морские ИБП On-Line", "Сглаживание резких скачков напряжения судовой сети 380В."),
        ("Сброс WAL буферов", "При падении питания СУБД форсированно фиксирует транзакции за < 2 сек."),
    ], accent_color=RED_ACCENT)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "ДУБЛИРОВАНИЕ СЕРВЕРОВ", [
        ("Разнесение по отсекам", "Основной и резервный серверы физически разнесены по разным бортам судна."),
        ("Синхронная репликация", "RPO = 0 при затоплении или пожаре в одном из серверных помещений."),
    ], accent_color=EMERALD)

# ==================== SLIDE 9: SPARE PARTS IN ISOLATION ====================
def slide_9():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "МАТЕРИАЛЬНОЕ СНАБЖЕНИЕ", "Управление складом ЗИП в автономном арктическом походе",
               "Математическое прогнозирование достаточности запаса запчастей на многомесячный рейс")

    top = Inches(2.0)
    w = Inches(3.7)
    gap = Inches(0.3)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "СПЕЦИФИКА СУДОВОГО ЗИП", [
        ("Неснижаемый аварийный запас", "Критическая номенклатура, отсутствие которой запрещает выход в рейс по требованиям РМРС."),
        ("Многоместное хранение", "Учет точных координат размещения деталей в десятках выгородок и трюмов."),
        ("Особые условия хранения", "Контроль температурного режима для специальных резинотехнических изделий (РТИ).")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "МАТЕМАТИКА РАСЧЕТА ЗАПАСА", [
        ("Закон распределения Пуассона", r"Расчет вероятности потребности в запасных частях с учетом интенсивности отказов $\lambda$."),
        ("Планирование рейса", "ТРИМ формирует заказ ЗИП на берег с учетом запланированного ледового маршрута."),
        ("Учет сроков годности", "Автоматическое оповещение об истечении гарантийного срока хранения смазок и фильтров.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "АДРЕСНЫЙ ПОИСК ДЕТАЛЕЙ", [
        ("Штрихкодирование деталей", "Идентификация узлов взрывозащищенными терминалами сбора данных."),
        ("Мгновенная локализация", "При аварии механик за секунды видит: на каком стеллаже и в каком ящике лежит нужный клапан."),
        ("Автоматическое списание", "Списание детали из остатков происходит строго по закрытии наряд-заказа.")
    ], accent_color=EMERALD)

# ==================== SLIDE 10: IMPORT SUBSTITUTION & COMPLIANCE ====================
def slide_10():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ИМПОРТОЗАМЕЩЕНИЕ", "Переход комплекса ТРИМ на доверенный отечественный стек",
               "Эксплуатация в среде Astra Linux Special Edition и защищенной СУБД Postgres Pro")

    top = Inches(2.0)
    w = Inches(2.75)
    gap = Inches(0.23)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "ОПЕРАЦИОННАЯ СИСТЕМА", [
        ("Astra Linux SE", "Полный переход серверной и клиентской частей на релиз «Смоленск» / «Воронеж»."),
        ("Замкнутая среда", "Запрет запуска неавторизованных исполняемых файлов на АРМ экипажа."),
        ("Мандатный доступ", "Изоляция служебных контуров судна по уровням конфиденциальности.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "СУБД СПЕЦНАЗНАЧЕНИЯ", [
        ("Postgres Pro Enterprise", "Отказ от MS SQL Server и Oracle в пользу сертифицированной российской СУБД."),
        ("Сжатие данных", "Эффективная компрессия CFS для экономии места на судовых SSD-накопителях."),
        ("Отказоустойчивость", "Встроенные механизмы отказоустойчивых кластеров высокой готовности.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "ВЕБ-ИНТЕРФЕЙС ТРИМ", [
        ("Кроссплатформенность", "Работа операторов через защищенный отечественный браузер («Chromium-ГОСТ»)."),
        ("Адаптивность под планшеты", "Работа судовых инженеров непосредственно у агрегатов по защищенному Wi-Fi."),
        ("Нулевой клиент", "Отсутствие необходимости установки тяжелого ПО на рабочие станции.")
    ], accent_color=PURPLE)

    add_card(s, Inches(0.8) + (w + gap)*3, top, w, h, "РЕГУЛЯТОРНЫЙ СТАТУС", [
        ("Реестр Минцифры", "ПК «ТРИМ» включен в Единый реестр российских программ (№ 1245)."),
        ("Сертификат РМРС", "Свидетельство о типовом одобрении программного комплекса морского применения."),
        ("Аттестация КИИ", "Готовность к развертыванию на объектах критической информационной инфраструктуры 1-й категории.")
    ], accent_color=EMERALD)

# ==================== SLIDE 11: SUMMARY MATRIX ====================
def slide_11():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "СРАВНИТЕЛЬНЫЙ АНАЛИЗ", "Сравнительная матрица: Традиционный ППР против EAM «ТРИМ»",
               "Количественные показатели эксплуатационной эффективности внедрения комплекса")

    rows = 6
    cols = 4
    left = Inches(0.8)
    top = Inches(2.0)
    width = Inches(11.7)
    height = Inches(4.7)

    table_shape = s.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    table.columns[0].width = Inches(3.2)
    table.columns[1].width = Inches(2.8)
    table.columns[2].width = Inches(3.2)
    table.columns[3].width = Inches(2.5)

    headers = ["Эксплуатационный фактор", "Традиционный плановый ремонт (ППР)", "EAM-комплекс ТРИМ (RCM)", "Выигрыш для спецобъекта"]
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = CARD_BG
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = CYAN
        p.alignment = PP_ALIGN.CENTER

    data = [
        ("Стратегия обслуживания", "Жесткий график по календарным часам", "По фактическому техсостоянию и наработке", "Оптимизация ремонтов"),
        ("Внезапные отказы агрегатов", "Высокий риск (скрытые дефекты не видны)", "Минимальный (предиктивное обнаружение)", "Снижение аварийности на 70%"),
        ("Расход судового ЗИП", "Избыточная плановая замена исправных узлов", "Замена строго по достижению лимита износа", "Экономия ЗИП 25–35%"),
        ("Учет наработки оборудования", "Ручные вахтенные журналы, риск искажений", "Автоматический импорт из АСУ ТП, ЭЦП", "100% достоверность"),
        ("Синхронизация с берегом", "Бумажные отчеты по возвращении в порт", "Асинхронные XML-пакеты по спутнику", "Контроль флота 24/7")
    ]

    for i, row in enumerate(data):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(24, 33, 47) if i % 2 == 0 else CARD_BG
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Arial"
            p.font.size = Pt(10)
            if j == 0:
                p.font.bold = True
                p.font.color.rgb = WHITE
            elif j == 3:
                p.font.bold = True
                p.font.color.rgb = EMERALD
            else:
                p.font.color.rgb = TEXT_MUTED

# ==================== SLIDE 12: CONCLUSION ====================
def slide_12():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ЗАКЛЮЧЕНИЕ", "Ключевые результаты и эксплуатационные рекомендации",
               "ТРИМ как эталон надежного информационного сопровождения объектов специального назначения")

    top = Inches(2.0)
    w = Inches(5.6)
    gap = Inches(0.5)
    h = Inches(2.2)

    add_card(s, Inches(0.8), top, w, h, "1. АВТОНОМНОСТЬ И НАДЕЖНОСТЬ", [
        ("Полная независимость", "Архитектура ТРИМ доказала надежность функционирования в условиях арктической изоляции и радиомолчания."),
        ("Пакетная репликация", "Эффективная синхронизация берегового и бортового сегментов по узкополосным каналам связи.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "2. НАУЧНЫЙ ПОДХОД RCM", [
        ("Управление рисками", "Внедрение анализа FMECA позволяет прогнозировать отказы критических агрегатов задолго до возникновения аварии."),
        ("Рост готовности", "Обеспечение коэффициента технической готовности атомных энергетических установок свыше 0.98.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8), top + h + Inches(0.3), w, h, "3. СКВОЗНАЯ ИНТЕГРАЦИЯ", [
        ("Связь с АСУ ТП", "Бесшовная передача параметров вибродиагностики и моточасов исключает ошибки человеческого фактора."),
        ("Юридический след", "Электронный формуляр с ЭЦП обеспечивает прозрачный контроль жизненного цикла оборудования.")
    ], accent_color=EMERALD)

    add_card(s, Inches(0.8) + w + gap, top + h + Inches(0.3), w, h, "4. ТЕХНОЛОГИЧЕСКИЙ СУВЕРЕНИТЕТ", [
        ("Доверенная среда", "Успешная эксплуатация на базе Astra Linux SE и Postgres Pro подтверждает готовность к импортозамещению на объектах КИИ 1-й категории."),
        ("Стандарт отрасли", "Многолетний опыт эксплуатации на ледоколах ФГУП «Атомфлот» является эталоном для спецсистем РФ.")
    ], accent_color=PURPLE)

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
print("TRIM Presentation saved successfully at:", pptx_path)

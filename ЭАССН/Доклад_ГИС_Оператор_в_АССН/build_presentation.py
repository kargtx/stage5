# -*- coding: utf-8 -*-
import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

out_dir = r"c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\Доклад_ГИС_Оператор_в_АССН"
fig_dir = os.path.join(out_dir, "figures")
pptx_path = os.path.join(out_dir, "ГИС_Оператор_в_АССН_Презентация.pptx")

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
    p1.text = "Эксплуатационные аспекты и безопасность\nраспределенной ГИС на примере комплекса «Оператор»"
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
        ("Студент", "Каргапольцев Тимофей Рамилевич"),
        ("Группа", "ЭАС-514С (5 курс)"),
        ("Специальность", "Эксплуатация автоматизированных систем специального назначения"),
    ], accent_color=CYAN)

    add_card(s, Inches(6.8), Inches(4.7), Inches(5.7), Inches(2.1), "НАУЧНЫЙ РУКОВОДИТЕЛЬ", [
        ("Преподаватель", "Ризванов Камиль Анварович"),
        ("Ученая степень", "Кандидат технических наук, доцент"),
        ("Кафедра", "Автоматизированных систем управления (АСУ)"),
    ], accent_color=EMERALD)

# ==================== SLIDE 2: SPECIAL GIS CONTEXT ====================
def slide_2():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "СПЕЦИФИКА ЗАДАЧ", "Геоинформационные системы в контуре спецназначения",
               "Оперативное звено управления, командные пункты и жесткие требования к живучести данных")

    w = Inches(2.75)
    gap = Inches(0.23)
    top = Inches(2.0)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "1. СКОРОСТЬ ОТКЛИКА", [
        ("Реальное время", "Отображение тысяч подвижных объектов (авиация, БПЛА, техника) без задержек."),
        ("Кадровая частота", "Плавный рендеринг карты при масштабировании (не менее 25–30 FPS)."),
        ("Поддержка решений", "Время расчета зон досягаемости оружия должно занимать доли секунды.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "2. СЕКРЕТНОСТЬ ДАННЫХ", [
        ("Многоуровневость", "На одной картоснове совмещаются слои разного грифа (открытые, ДСП, Секретно)."),
        ("Мандатная защита", "Каждый слой защищен меткой безопасности операционной системы."),
        ("Изоляция сетей", "Полный запрет выхода во внешние публичные картографические сервисы.")
    ], accent_color=RED_ACCENT)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "3. АВТОНОМНОСТЬ (OFF-LINE)", [
        ("Полевые условия", "Работа на переносных планшетах при полном радиоэлектронном подавлении (РЭБ)."),
        ("Локальный кэш", "Все необходимые матрицы рельефа и тайлы хранятся на защищенном накопителе."),
        ("Отсутствие облаков", "100% независимость от удаленных внешних серверов геоданных.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8) + (w + gap)*3, top, w, h, "4. СТАНДАРТЫ ОФОРМЛЕНИЯ", [
        ("Классификаторы КЭО", "Строгое соответствие условным знакам оперативно-тактической обстановки (ГОСТ)."),
        ("Точность геодезии", "Использование эллипсоида ПЗ-90.11 и СК-42/95 с миллиметровой привязкой."),
        ("3D-моделирование", "Построение зон прямой видимости и профилей высот местности.")
    ], accent_color=EMERALD)

# ==================== SLIDE 3: GIS OPERATOR OVERVIEW ====================
def slide_3():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ПЛАТФОРМА ГИС", "ГИС «Оператор» (КБ «Панорама») — базовый стандарт ведомств",
               "Штатный программный комплекс геопространственного обеспечения ВС РФ и силовых структур")

    top = Inches(2.0)
    w = Inches(3.7)
    gap = Inches(0.3)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "ПРИНЯТИЕ НА СНАБЖЕНИЕ", [
        ("Штатный статус", "Принят на снабжение ВС РФ приказом Министра обороны РФ."),
        ("Сертификация безопасности", "Сертификат МО РФ по 2-му уровню контроля отсутствия недекларированных возможностей (НДВ)."),
        ("Спецведомства", "Широко применяется в МЧС, Росгвардии и ситуационных центрах безопасности.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "ФОРМАТНЫЙ БАЗИС", [
        ("Форматы SXF / SITX", "Отечественные векторные форматы пространственных данных открытого и закрытого типа."),
        ("Матрицы рельефа MTW", "Высокоточные матрицы высот для трехмерного моделирования театра действий."),
        ("Пакетные тайлы Mbtiles", "Эффективная нарезка растровых подложек для мобильных переносных АРМ.")
    ], accent_color=EMERALD)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "СИСТЕМНЫЙ СТЕК", [
        ("Astra Linux SE", "Основная операционная система функционирования серверной и клиентской частей."),
        ("Ядро СУБД", "Интеграция с защищенными СУБД Postgres Pro Enterprise и Линтер-ВС."),
        ("Аппаратные архитектуры", "Поддержка платформ x86-64, «Эльбрус» (e2k) и «Байкал» (ARM).")
    ], accent_color=PURPLE)

# ==================== SLIDE 4: ARCHITECTURE DIAGRAM ====================
def slide_4():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "АРХИТЕКТУРА СИСТЕМЫ", "Сквозной контур защищенной геоинформационной системы",
               "Интеграция ГИС Сервера, мандатной фильтрации, полевой связи и терминалов операторов")

    img_path = os.path.join(fig_dir, "diagram_gis_operator_architecture.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "ГИС СЕРВЕР", [
        ("Банк данных", "Централизованное защищенное хранение карт масштабов от 1:1 000 000 до 1:10 000."),
        ("Многопоточный доступ", "Одновременное обслуживание сотен рабочих мест без падения производительности."),
    ], accent_color=CYAN)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "ЗАЩИЩЕННЫЙ ТРАНСПОРТ", [
        ("Полевая связь КАС ТР", "Передача обновлений обстановки по защищенным радиомодемам в условиях помех."),
        ("ГОСТ-шифрование", "Защита каждого передаваемого координатного пакета аппаратным СКЗИ."),
    ], accent_color=EMERALD)

# ==================== SLIDE 5: TILE PYRAMIDS & CACHE ====================
def slide_5():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ОПТИМИЗАЦИЯ ДАННЫХ", "Пирамиды пространственных тайлов и автономный кэш",
               "Технологии мгновенного зумирования и автономной работы полевых АРМ без связи")

    img_path = os.path.join(fig_dir, "diagram_gis_tile_pyramid_cache.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "ПИРАМИДНАЯ НАРЕЗКА", [
        ("Генерализация", "Каждому масштабу соответствует свой уровень детализации векторных объектов."),
        ("Мгновенный рендер", "Отрисовываются только тайлы, попадающие в видимый видовой экран оператора."),
    ], accent_color=AMBER)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "АВТОНОМНЫЙ ПАКЕТ", [
        ("Предзагрузка района", "Перед выездом на планшет загружается зашифрованный тайловый архив района дежурства."),
        ("Нулевой сетевой трафик", "Картооснова вообще не запрашивается из сети во время выполнения задачи."),
    ], accent_color=EMERALD)

# ==================== SLIDE 6: REALTIME TACTICAL PICTURE ====================
def slide_6():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ОПЕРАТИВНАЯ ОБСТАНОВКА", "Отображение динамической тактической обстановки реального времени",
               "Сквозной пайплайн: от радиосигналов ГЛОНАСС до оперативного экрана командующего")

    img_path = os.path.join(fig_dir, "diagram_gis_tactical_realtime.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "КЛАССИФИКАТОР КЭО", [
        ("Стандартизация", "Использование цифрового классификатора оперативных условных знаков."),
        ("Векторные знаки", "Масштабирование символов без потери четкости при любом приближении."),
    ], accent_color=CYAN)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "ПРОИЗВОДИТЕЛЬНОСТЬ", [
        ("До 5 000 объектов", "Отображение перемещения воздушных и наземных целей в реальном масштабе времени."),
        ("Экстраполяция", "При задержке радиосигнала ГИС рассчитывает предполагаемое положение цели по вектору скорости."),
    ], accent_color=RED_ACCENT)

# ==================== SLIDE 7: MANDATORY ACCESS CONTROL ====================
def slide_7():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ", "Мандатное разграничение доступа (МКД) к картографическим слоям",
               "Изоляция секретных пространственных данных на базе механизмов ОС Astra Linux Special Edition")

    img_path = os.path.join(fig_dir, "diagram_gis_security_levels.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "МАНДАТНАЯ МЕТКА СЛОЯ", [
        ("Гриф секретности", "Каждый слой карты (дороги, мосты, дислокация, замысел) имеет свою мандатную категорию."),
        ("Авто-скрытие", "Оператор без допуска формы №2 физически не видит закрытые слои в списке и на экране."),
    ], accent_color=RED_ACCENT)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "КОНТРОЛЬ ЦЕЛОСТНОСТИ", [
        ("Защита от подмены", "Цифровая подпись картографических файлов исключает искажение координат диверсантами."),
        ("Журнал аудита", "Фиксация каждого факта экспорта фрагмента карты или печати боевого документа."),
    ], accent_color=EMERALD)

# ==================== SLIDE 8: DATA SYNC OVER LOW-BANDWIDTH COMMS ====================
def slide_8():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "СЕТЕВАЯ ЭКСПЛУАТАЦИЯ", "Синхронизация обстановки по низкоскоростным радиоканалам",
               "Инкрементальная передача тактических слоев со скоростями от 1.2 до 19.2 кбит/с")

    top = Inches(2.0)
    w = Inches(3.7)
    gap = Inches(0.3)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "ИНКРЕМЕНТАЛЬНЫЙ ОБМЕН", [
        ("Передача только дельты", "Передаются не карты, а только короткие бинарные телеграммы изменений (ID объекта, X, Y, H, статус)."),
        ("Размер пакета", "Средний объем оперативного донесения — всего от 64 до 256 байт."),
        ("Сжатие заголовков", "Специализированные протоколы минимизации сетевого служебного оверхеда.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "ПОМЕХОУСТОЙЧИВОСТЬ", [
        ("Упреждающая коррекция (FEC)", "Восстановление поврежденных радиопакетов без повторного запроса."),
        ("Контрольные суммы ГОСТ", "Исключение искажения координат при воздействии комплексов радиоподавления (РЭБ)."),
        ("Очередь с приоритетами", "Приоритетная передача сигналов бедствия и тревог перед рутинными сводками.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "МЕЖСЕТЕВОЙ ЭКРАН", [
        ("Защита периметра", "Фильтрация входящего картографического трафика на аппаратных шлюзах."),
        ("Защита от инъекций", "Блокировка вредоносного бинарного кода, скрытого в структурах векторных карт."),
        ("Изоляция подсетей", "Штабные серверы отделены от радиосетей передовых отрядов демилитаризованной зоной.")
    ], accent_color=EMERALD)

# ==================== SLIDE 9: FIELD TERMINALS & EMSEC ====================
def slide_9():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ЖИВУЧЕСТЬ АРМ", "Надежность и безопасность полевых терминалов ГИС",
               "Защищенные переносные ЭВМ командиров, экранирование ПЭМИН и стойкость к повреждениям")

    top = Inches(2.0)
    w = Inches(3.7)
    gap = Inches(0.3)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "КЛИМАТИЧЕСКАЯ СТОЙКОСТЬ", [
        ("Исполнение IP67 / ГОСТ РВ", "Полная пыле- и влагозащита, погружение в воду, работа под дождем и снегом."),
        ("Температурный диапазон", "Гарантированная работа от -40°C до +50°C с подогревом экранной матрицы."),
        ("Ударопрочность", "Выдерживает падение на бетон с высоты 1.5 метра и вибрацию гусеничной бронетехники.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "ЗАЩИТА ОТ УТЕЧЕК (ПЭМИН)", [
        ("Экранирование корпуса", "Подавление побочных электромагнитных излучений и наводок для исключения перехвата экрана."),
        ("Оптическая фильтрация", "Специальные поляризационные фильтры дисплея, скрывающие изображение от взглядов сбоку."),
        ("Режим маскировки", "Специальная ночная подсветка экрана, не демаскирующая позицию приборами ночного видения.")
    ], accent_color=RED_ACCENT)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "ЭКСТРЕННОЕ УНИЧТОЖЕНИЕ", [
        ("Аппаратная кнопка стирания", "Мгновенное физическое или криптографическое уничтожение ключей и картографической базы."),
        ("Защита от захвата", "При попадании планшета в руки противника секретная оперативная обстановка не может быть прочитана."),
        ("Блокировка по датчикам", "Автоматическая блокировка АРМ при несанкционированном вскрытии корпуса.")
    ], accent_color=EMERALD)

# ==================== SLIDE 10: DOMESTIC TECH STACK ====================
def slide_10():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ИМПОРТОНЕЗАВИСИМОСТЬ", "Технологический суверенитет программного ядра ГИС «Оператор»",
               "Полный отказ от зарубежных библиотек, нативная кроссплатформенность C++ / Qt")

    top = Inches(2.0)
    w = Inches(2.75)
    gap = Inches(0.23)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "СОБСТВЕННОЕ ЯДРО", [
        ("С++ / Qt Framework", "Высокооптимизированный нативный код без использования уязвимых виртуальных машин."),
        ("Чистая разработка", "Отсутствие сторонних закрытых библиотек Google, Esri или OpenLayers."),
        ("Исходный код РФ", "100% разработка на территории Российской Федерации (г. Ногинск / Москва).")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "СУБД СПЕЦНАЗНАЧЕНИЯ", [
        ("Пространственный сервер", "ГИС Сервер работает напрямую с СУБД Postgres Pro / PostGIS и Линтер-ВС."),
        ("Транзакционная защита", "Многоверсионность и надежная защита целостности геобаз при сбоях."),
        ("Масштабирование", "Поддержка терабайтных пространственных банков данных.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "ПРОЦЕССОРЫ РФ", [
        ("Эльбрус (e2k)", "Официальная сборка и сертификация под отечественную микропроцессорную архитектуру «Эльбрус»."),
        ("Байкал (ARM)", "Стабильная работа на энергоэффективных мобильных отечественных процессорах."),
        ("Оптимизация команд", "Использование аппаратно-специфичных инструкций для ускорения графики.")
    ], accent_color=PURPLE)

    add_card(s, Inches(0.8) + (w + gap)*3, top, w, h, "АТТЕСТАЦИЯ И СТАНДАРТЫ", [
        ("Приказ МО РФ №111", "Соответствие жестким требованиям нормативно-технических документов геодезической службы."),
        ("Единый реестр ПО", "Включена в реестр российского программного обеспечения Минцифры РФ."),
        ("Готовность к бою", "Проверена многолетним опытом боевого дежурства и учений ВС РФ.")
    ], accent_color=EMERALD)

# ==================== SLIDE 11: SUMMARY MATRIX ====================
def slide_11():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "СРАВНИТЕЛЬНЫЙ АНАЛИЗ", "Сравнительная матрица: Гражданские ГИС vs ГИС «Оператор»",
               "Ключевые эксплуатационные различия коммерческих и специализированных геосистем")

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

    headers = ["Критерий эксплуатации", "Гражданские ГИС (Esri ArcGIS / QGIS)", "ГИС «Оператор» (КБ «Панорама»)", "Значимость для АССН"]
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
        ("Сетевая зависимость", "Требуют стабильного интернет-доступа к серверам", "100% автономная работа с локальным кэшем", "Критически важно в поле"),
        ("Разграничение доступа", "Дискреционное (логин / пароль / роли)", "Мандатное (МКД Astra Linux по метке слоя)", "Защита гостайны"),
        ("Классификатор символов", "Произвольные гражданские иконки", "Стандартизованные условные знаки ГОСТ/КЭО", "Единообразие управления"),
        ("Сертификация безопасности", "Сертификаты отсутствуют либо минимальны", "Сертификат МО РФ (2-й уровень контроля НДВ)", "Официальный допуск"),
        ("Работа по радиомодемам", "Не поддерживается (тяжелый трафик Web-API)", "Специализированные дельта-протоколы", "Связь при помехах")
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
    add_header(s, "ЗАКЛЮЧЕНИЕ", "Ключевые результаты и эксплуатационные выводы",
               "ГИС «Оператор» как фундаментальная платформа ситуационной осведомленности спецсистем")

    top = Inches(2.0)
    w = Inches(5.6)
    gap = Inches(0.5)
    h = Inches(2.2)

    add_card(s, Inches(0.8), top, w, h, "1. СИТУАЦИОННОЕ ПРЕИМУЩЕСТВО", [
        ("Реальное время", "Отображение динамической тактической обстановки с задержкой менее секунды обеспечивает превосходство в принятии решений."),
        ("Плавный рендеринг", "Пирамиды тайлов гарантируют мгновенную навигацию при любых масштабах карт.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "2. БЕЗУПРЕЧНАЯ БЕЗОПАСНОСТЬ", [
        ("Мандатная фильтрация", "МКД в среде Astra Linux SE делает невозможной утечку секретных слоев операторам без соответствующей формы допуска."),
        ("Контроль целостности", "ГОСТ-подпись исключает искажение картографической информации.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8), top + h + Inches(0.3), w, h, "3. НАДЕЖНОСТЬ В ПОЛЕ", [
        ("Узкополосная связь", "Инкрементальный обмен по радиомодемам обеспечивает актуальность данных при скорости каналов от 1.2 кбит/с."),
        ("Автономный кэш", "Полная боеспособность полевых планшетов при радиоэлектронном подавлении.")
    ], accent_color=EMERALD)

    add_card(s, Inches(0.8) + w + gap, top + h + Inches(0.3), w, h, "4. ТЕХНОЛОГИЧЕСКИЙ СУВЕРЕНИТЕТ", [
        ("100% отечественный стек", "Нативная поддержка Astra Linux, защищенных СУБД и отечественных процессоров «Эльбрус»."),
        ("Основа АСУ СН", "Базовый геоинформационный компонент для современных систем управления специального назначения.")
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
print("GIS Operator Presentation saved successfully at:", pptx_path)

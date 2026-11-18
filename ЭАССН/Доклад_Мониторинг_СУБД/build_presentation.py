# -*- coding: utf-8 -*-
import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

out_dir = r"c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\Доклад_Мониторинг_СУБД"
fig_dir = os.path.join(out_dir, "figures")
pptx_path = os.path.join(out_dir, "Мониторинг_и_диагностика_СУБД_Презентация.pptx")

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
    # Category tag
    cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.35))
    tf_cat = cat_box.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category.upper()
    p_cat.font.name = "Arial"
    p_cat.font.size = Pt(11)
    p_cat.font.bold = True
    p_cat.font.color.rgb = CYAN

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(11.7), Inches(0.65))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.text = title
    p_t.font.name = "Arial"
    p_t.font.size = Pt(21)
    p_t.font.bold = True
    p_t.font.color.rgb = WHITE

    # Subtitle if present
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
    # Card shape
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = CARD_BORDER
    card.line.width = Pt(1.2)

    # Text box
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

    # University Badge
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

    # Big Title
    title_box = s.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.7), Inches(2.2))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    p1 = tf_t.paragraphs[0]
    p1.text = "Комплексный мониторинг, аудит\nи диагностика производительности СУБД\nв режиме реального времени"
    p1.font.name = "Arial"
    p1.font.size = Pt(32)
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    p1.space_after = Pt(14)

    p2 = tf_t.add_paragraph()
    p2.text = "Дисциплина: «Эксплуатация автоматизированных систем специального назначения» (ЭАССН)"
    p2.font.name = "Arial"
    p2.font.size = Pt(15)
    p2.font.color.rgb = AMBER

    # Bottom Meta Cards
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

# ==================== SLIDE 2: CONTEXT & SPECIFICS ====================
def slide_2():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "АКТУАЛЬНОСТЬ И СПЕЦИФИКА", "Специфика эксплуатации СУБД в АС специального назначения",
               "Высокая критичность данных и цена простоя диктуют бескомпромиссные требования к наблюдаемости")

    w = Inches(2.75)
    gap = Inches(0.23)
    top = Inches(2.0)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "1. ЖЕСТКИЕ SLA / 24/7", [
        ("Непрерывность", "Недопустимость технологического простоя в боевых и диспетчерских контурах."),
        ("RTO ~ 0", "Восстановление работоспособности за секунды."),
        ("Риск отказа", "Отказ СУБД влечет потерю управляемости всем комплексом АССН.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "2. «ТИХАЯ» ДЕГРАДАЦИЯ", [
        ("Коварство", "Замедление ответов критичнее полного сбоя: система работает, но операторы блокированы."),
        ("Каскадный сбой", "Накопление очередей транзакций и лавинообразный отказ пулов соединений."),
        ("Детекция", "Требуется упреждающее выявление аномалий до жалоб пользователей.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "3. ТРЕБОВАНИЯ ИБ", [
        ("Регуляторы", "Соответствие требованиям ФСТЭК России по защите информации ограниченного доступа."),
        ("Контроль прав", "Тотальный аудит действий суперпользователей (DBA) и привилегированных ролей."),
        ("Защита периметра", "Работа в изолированных контурах без выхода во внешние облака.")
    ], accent_color=RED_ACCENT)

    add_card(s, Inches(0.8) + (w + gap)*3, top, w, h, "4. МИНИМАЛЬНЫЙ OVERHEAD", [
        ("Лимит ресурсов", "Накладные расходы мониторинга не должны превышать 2-3% CPU/RAM."),
        ("Неинвазивность", "Сбор телеметрии не должен провоцировать конкуренцию за блокировки ячеек памяти."),
        ("Автономность", "Сохранение работоспособности при сбоях сетевой подсистемы.")
    ], accent_color=EMERALD)

# ==================== SLIDE 3: THREE PILLARS OF OBSERVABILITY ====================
def slide_3():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "МЕТОДОЛОГИЯ НАБЛЮДАЕМОСТИ", "Концепция сквозной наблюдаемости (Full-Stack Observability)",
               "Разделение контуров телеметрии: от операционной системы до единичных предикатов SQL-запроса")

    top = Inches(2.0)
    w = Inches(3.7)
    gap = Inches(0.3)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "ИНФРАСТРУКТУРНЫЙ УРОВЕНЬ", [
        ("CPU & Steal", "Загрузка физических ядер, прерывания, троттлинг виртуализации."),
        ("Дисковая шина (I/O)", "Задержки ввода-вывода (IOPS, Disk Latency, Queue Length). Деградация дисков — главный тормоз СУБД."),
        ("Оперативная память", "Использование RAM, давление на HugePages, недопустимость свопирования (SWAP = смерть СУБД)."),
        ("Сетевой стек", "Пропускная способность адаптеров, дропы пакетов, задержки сокетов межсерверной репликации.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "УРОВЕНЬ ЯДРА СУБД (KERNEL)", [
        ("SGA & Buffer Cache", "Эффективность кэширования блоков данных (Hit Ratio > 95%), перевыделение памяти."),
        ("Журналирование (Redo)", "Скорость работы процесса LGWR, ротация Redo-журналов, утилизация буфера повтора."),
        ("Защелки (Latches/Mutexes)", "Конкуренция за структуры оперативной памяти ядра, устранение Buffer Busy Waits."),
        ("Сессионный пул", "Баланс активных и фоновых процессов, мониторинг очередей квотирования ресурсов (Resource Manager).")
    ], accent_color=PURPLE)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "ПРИКЛАДНОЙ УРОВЕНЬ (SQL)", [
        ("Латентность (Latency)", "Время исполнения транзакций в 95-м и 99-м перцентилях (P95/P99)."),
        ("Пропускная способность", "Количество транзакций и коммитов в секунду (TPS / QPS)."),
        ("Регрессия планов", "Внезапные изменения планов выполнения SQL-запросов Стоимостным оптимизатором (CBO)."),
        ("Очереди блокировок", "Взаимные блокировки транзакций (Deadlocks) и долгие ожидания на строках таблиц.")
    ], accent_color=EMERALD)

# ==================== SLIDE 4: ARCHITECTURE DIAGRAM ====================
def slide_4():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "АРХИТЕКТУРНОЕ РЕШЕНИЕ", "Сквозной контур мониторинга, аудита и диагностики СУБД",
               "Интеграция телеметрии ядра, сбора метрик временных рядов и ситуационного центра")

    img_path = os.path.join(fig_dir, "diagram_monitoring_architecture.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    # Right side explanatory cards
    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "СБОР ТЕЛЕМЕТРИИ", [
        ("Неблокирующий опрос", "Экспортеры читают легковесные $V-представления и системные счетчики."),
        ("Сжатие трафика", "Протокол pull-модели Prometheus снижает нагрузку на сеть АССН."),
    ], accent_color=CYAN)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "ПРЕДИКТИВНЫЙ АЛЕРТИНГ", [
        ("Упреждение сбоев", "Alertmanager предупреждает о росте задержек за 10-15 минут до исчерпания буферов."),
        ("Экспорт в SIEM", "Аудит передается по защищенному Syslog без задержки транзакций."),
    ], accent_color=EMERALD)

# ==================== SLIDE 5: ORACLE TELEMETRY (ASH, AWR, ADDM) ====================
def slide_5():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ВНУТРЕННИЕ МЕХАНИЗМЫ СУБД", "Интроспекция рабочей нагрузки Oracle: ASH, AWR и ADDM",
               "Глубокий анализ ядра без существенных накладных расходов на процессорную подсистему")

    top = Inches(2.0)
    w = Inches(2.75)
    gap = Inches(0.23)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "ACTIVE SESSION HISTORY", [
        ("Семплирование 1 сек", "Каждую секунду фиксируется состояние всех не-idle сессий."),
        ("Кольцевой буфер SGA", "Хранение в оперативной памяти (V$ACTIVE_SESSION_HISTORY) — нулевой диск. оверхед."),
        ("Точный срез", "Показывает: кто выполнял, какой SQL_ID, какой объект БД и какое ожидание зафиксировано.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "WORKLOAD REPOSITORY", [
        ("AWR Снапшоты", "Регулярный сброс (раз в час) кумулятивных метрик производительности в системные таблицы."),
        ("Исторический тренд", "Возможность сравнить профиль нагрузки «сегодня» с «прошлой неделей»."),
        ("Базовые линии (Baselines)", "Фиксация нормального рабочего профиля нагрузки АССН для поиска аномалий.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "DIAGNOSTIC MONITOR", [
        ("ADDM Анализ", "Встроенная экспертная система искусственного интеллекта СУБД."),
        ("Корневые причины", "ADDM автоматически ранжирует узкие места по их вкладу в совокупное время DB Time."),
        ("Готовые рекомендации", "Предлагает: добавить индекс, изменить размер SGA, переписать неоптимальный SQL.")
    ], accent_color=EMERALD)

    add_card(s, Inches(0.8) + (w + gap)*3, top, w, h, "СЛОВАРЬ ПРОИЗВОДИТЕЛЬНОСТИ", [
        ("V$SESSION", "Текущие пользовательские сессии, статус, блокирующий SID (blocking_session)."),
        ("V$SQL_AREA", "Кэш разобранных SQL-запросов, число выполнений, время ЦПУ, дисковые чтения."),
        ("V$SYSTEM_EVENT", "Сводная статистика ожиданий с момента старта инстанса СУБД.")
    ], accent_color=PURPLE)

# ==================== SLIDE 6: WAIT MODEL DIAGRAM ====================
def slide_6():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "МОДЕЛЬ СОБЫТИЙ ОЖИДАНИЯ", "Декомпозиция DB Time и философия Wait Interface",
               "Эффективная методология тюнинга: не оптимизировать то, что не тратит время пользователей")

    img_path = os.path.join(fig_dir, "diagram_oracle_wait_model.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "ПРАВИЛО РАСПРЕДЕЛЕНИЯ", [
        ("Формула баланса", "DB Time = Время работы CPU + Время не-холостых ожиданий (Waits)."),
        ("Приоритет тюнинга", "Оптимизируется компонент, занимающий наибольший сегмент в диаграмме времени."),
    ], accent_color=CYAN)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "ИНДИКАТОРЫ УЗКИХ МЕСТ", [
        ("db file sequential read", "Одиночное чтение блоков. Норма для OLTP, но рост задержки говорит о перегрузке дисков."),
        ("enq: TX - row lock", "Блокировка строк. Признак неоптимальной архитектуры транзакций в ПО."),
    ], accent_color=RED_ACCENT)

# ==================== SLIDE 7: FGA & SECURITY AUDIT ====================
def slide_7():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ", "Тонко гранулированный аудит (Fine-Grained Auditing — FGA)",
               "Адресный контроль доступа к секретным атрибутам без захламления системных дисков")

    img_path = os.path.join(fig_dir, "diagram_audit_fga_workflow.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "ПАКЕТ DBMS_FGA", [
        ("Условие аудита", "Политика активируется только при совпадении предиката (например, доступ к спецкатегории)."),
        ("Контроль столбцов", "Аудит инициируется при чтении конкретных полей (гостайна, персональные данные)."),
    ], accent_color=EMERALD)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "ЗАЩИТА ОТ ИНСАЙДЕРОВ", [
        ("Неуязвимость лога", "Журнал Unified Audit защищен от модификации даже привилегированным DBA."),
        ("Полный контекст", "Фиксируется реальный IP терминала, имя ОС, стек вызовов и значения bind-переменных."),
    ], accent_color=AMBER)

# ==================== SLIDE 8: CBO & EXECUTION PLANS ====================
def slide_8():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ОПТИМИЗАЦИЯ И СТАБИЛИЗАЦИЯ", "Стоимостной оптимизатор (CBO) и планы выполнения SQL",
               "Предотвращение внезапных деградаций планов выполнения в высоконагруженных спецсистемах")

    top = Inches(2.0)
    w = Inches(3.7)
    gap = Inches(0.3)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "СТОИМОСТНОЙ ОПТИМИЗАТОР (CBO)", [
        ("Принцип работы", "Оценивает тысячи вариантов выполнения запроса и выбирает план с наименьшей оценочной стоимостью (Cost)."),
        ("Роль статистики", "Критическая зависимость от свежести статистики объектов (DBMS_STATS)."),
        ("Ошибки кардинальности", "Искаженная гистограмма распределения значений ведет к катастрофическому выбору Full Scan вместо индекса.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "ИНСТРУМЕНТЫ АНАЛИЗА ПЛАНОВ", [
        ("EXPLAIN PLAN", "Предварительный теоретический план выполнения запроса."),
        ("V$SQL_PLAN / DISPLAY_CURSOR", "Фактический план выполнения из кэша библиотеки с реальным числом обработанных строк (A-Rows)."),
        ("Трассировка 10046 (tkprof)", "Детальный пошаговый профиль выполнения: задержки на уровне каждого шага плана и системных вызовов."),
        ("Обнаружение деградации", "Автоматическое выявление смены Plan Hash Value при росте времени исполнения.")
    ], accent_color=PURPLE)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "СТАБИЛИЗАЦИЯ (SQL BASELINES)", [
        ("SQL Plan Baselines", "Фиксация эталонного проверенного плана выполнения в репозитории СУБД."),
        ("Защита от деградации", "Оптимизатор не имеет права переключиться на непроверенный план, даже если посчитает его дешевле."),
        ("Эволюция планов", "Новый план тестируется в фоновом режиме и принимается только после доказанного ускорения."),
        ("Надежность для АССН", "100% защита от непредсказуемых сбоев при обновлении версий СУБД или накатке патчей.")
    ], accent_color=EMERALD)

# ==================== SLIDE 9: TUNING CYCLE ====================
def slide_9():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "РЕГЛАМЕНТ ЭКСПЛУАТАЦИИ", "Замкнутый цикл устранения деградации производительности",
               "Пошаговый регламент инженера АССН: от получения алерта до фиксации стабильности")

    img_path = os.path.join(fig_dir, "diagram_performance_tuning_cycle.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "СКОРОСТЬ РЕАКЦИИ", [
        ("Норматив MTTR", "Среднее время локализации узкого места в АССН не должно превышать 3-5 минут."),
        ("Изоляция без остановки", "Использование ASH позволяет выявить виновный SQL без остановки инстанса БД."),
    ], accent_color=AMBER)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "ПОСТ-ИНЦИДЕНТНЫЙ АНАЛИЗ", [
        ("Фиксация ретроспективы", "Сохранение среза AWR за период инцидента для расследования причин."),
        ("Защита от рецидивов", "Создание постоянного правила мониторинга или профиля плана запроса."),
    ], accent_color=EMERALD)

# ==================== SLIDE 10: IMPORT SUBSTITUTION (POSTGRES PRO) ====================
def slide_10():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ИМПОРТОЗАМЕЩЕНИЕ И ОТЕЧЕСТВЕННЫЙ СТЕК", "Специфика мониторинга и аудита в Postgres Pro Enterprise",
               "Перенос компетенций администрирования Oracle на сертифицированные российские СУБД")

    top = Inches(2.0)
    w = Inches(2.75)
    gap = Inches(0.23)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "ИНТРОСПЕКЦИЯ ЯДРА", [
        ("pg_wait_sampling", "Отечественный аналог Oracle ASH в Postgres Pro. Обеспечивает семплирование ожиданий сессий."),
        ("pg_stat_statements", "Аналог V$SQL: статистика выполнения всех запросов, время CPU, дисковые чтения."),
        ("pg_stat_activity", "Текущие активные транзакции, задержки запросов, граф блокировок (lock contention).")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "СПЕЦИФИКА MVCC И ВАКУУМ", [
        ("Очистка (VACUUM)", "В отличие от Oracle Undo, в PostgreSQL версионность строк требует своевременного автовакуума."),
        ("Мониторинг Bloat", "Контроль разрастания «мертвых» строк в таблицах и индексах."),
        ("Транзакционный счетчик", "Мониторинг отставания автовакуума для защиты от Transaction ID Wraparound.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "СЕРТИФИЦИРОВАННЫЙ АУДИТ", [
        ("Модуль pgaudit", "Глубокое протоколирование действий пользователей для выполнения требований безопасности."),
        ("ФСТЭК России", "Сертификация Postgres Pro Enterprise по 1 классу защиты и 1 уровню доверия."),
        ("Защищенный журнал", "Передача событий аудита в изолированный контур SIEM без задержки пользовательских сессий.")
    ], accent_color=RED_ACCENT)

    add_card(s, Inches(0.8) + (w + gap)*3, top, w, h, "ПЛАТФОРМЫ УПРАВЛЕНИЯ", [
        ("PPEM Platform", "Postgres Pro Enterprise Manager — комплексный мониторинг кластеров и резервного копирования."),
        ("Tantor Platform", "Российская платформа мониторинга и оптимизации производительности PostgreSQL."),
        ("Open-source стек", "Интеграция с Prometheus postgres_exporter и специализированными дашбордами Grafana.")
    ], accent_color=EMERALD)

# ==================== SLIDE 11: SUMMARY MATRIX ====================
def slide_11():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "СРАВНИТЕЛЬНЫЙ АНАЛИЗ", "Сводная матрица инструментария мониторинга, аудита и диагностики",
               "Оценка применимости различных технологических уровней в контуре АС специального назначения")

    # Table
    rows = 5
    cols = 5
    left = Inches(0.8)
    top = Inches(2.0)
    width = Inches(11.7)
    height = Inches(4.7)

    table_shape = s.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    table.columns[0].width = Inches(2.7)
    table.columns[1].width = Inches(2.25)
    table.columns[2].width = Inches(2.25)
    table.columns[3].width = Inches(2.25)
    table.columns[4].width = Inches(2.25)

    headers = ["Инструментальный уровень", "Детализация данных", "Накладные расходы (Overhead)", "Скорость выявления инцидентов", "Применимость в АССН"]
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
        ("Внутренняя интроспекция ядра\n(Oracle ASH/AWR / pg_wait_sampling)",
         "Максимальная\n(до строки SQL и события)",
         "Минимальные\n(< 1-2% CPU, кольцевой буфер)",
         "В реальном времени\n(семплирование каждую секунду)",
         "Критически обязательно\n(основа диагностики)"),

        ("Внешний Time-Series мониторинг\n(Prometheus + Grafana + Exporters)",
         "Агрегированные метрики\n(TPS, Latency, очереди, диски)",
         "Низкие\n(опрос по таймеру раз в 5-15 с)",
         "Секундная детекция трендов\n(предиктивные алерты)",
         "Обязательно\n(ситуационный экран)"),

        ("Тонко гранулированный аудит\n(Oracle FGA / pgaudit)",
         "Контекст безопасности\n(IP, пользователь, точный SQL)",
         "Низкие\n(только по целевым предикатам)",
         "Немедленно при выполнении\n(экспорт в SIEM)",
         "Обязательно по нормам ФСТЭК\n(контроль спецданных)"),

        ("Тяжелая трассировка вызовов\n(SQL Trace 10046, расширенный лог)",
         "Исчерпывающая\n(каждый системный вызов)",
         "Высокие\n(до 20-40% деградации CPU/диска)",
         "Ретроспективная\n(анализ постфактум)",
         "Исключительно на тестовых\nстендах или локально")
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
            elif j == 4:
                p.font.bold = True
                p.font.color.rgb = EMERALD if "Обязательно" in val else AMBER
            else:
                p.font.color.rgb = TEXT_MUTED

# ==================== SLIDE 12: CONCLUSION ====================
def slide_12():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ЗАКЛЮЧЕНИЕ", "Ключевые результаты и эксплуатационные рекомендации",
               "Построение надежной и безопасной инфраструктуры данных для систем специального назначения")

    top = Inches(2.0)
    w = Inches(5.6)
    gap = Inches(0.5)
    h = Inches(2.2)

    add_card(s, Inches(0.8), top, w, h, "1. СИНЕРГИЯ ТЕЛЕМЕТРИИ", [
        ("Сквозной контроль", "Эффективная эксплуатация возможна только при объединении метрик ОС, интроспекции ядра (ASH) и мониторинга запросов в единый контур."),
        ("Отсутствие слепых зон", "Исключаются скрытые задержки между дисковой подсистемой и планировщиком СУБД.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "2. ПРЕДИКТИВНОСТЬ ВМЕСТО РЕАКЦИИ", [
        ("Упреждение сбоя", "Мониторинг должен выявлять тренд деградации за 10-20 минут до нарушения SLA и исчерпания пулов транзакций."),
        ("Снижение MTTR", "Автоматическая локализация виновного SQL_ID сокращает время устранения аварии до минут.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8), top + h + Inches(0.3), w, h, "3. БЕЗОПАСНОСТЬ БЕЗ ТОРМОЗОВ", [
        ("Контекстный аудит (FGA)", "Обеспечивает строгое выполнение требований ФСТЭК России без деградации производительности СУБД."),
        ("Защита от инсайдеров", "Гарантированная фиксация доступа к критическим данным с передачей событий в изолированный SIEM.")
    ], accent_color=EMERALD)

    add_card(s, Inches(0.8) + w + gap, top + h + Inches(0.3), w, h, "4. ГОТОВНОСТЬ К ИМПОРТОЗАМЕЩЕНИЮ", [
        ("Унификация методологии", "Принципы Wait Interface и интроспекции сессий едины как для Oracle (AWR/ASH), так и для Postgres Pro (pg_wait_sampling)."),
        ("Бесшовная миграция", "Обеспечивается непрерывный мониторинг производительности на всех этапах перехода на отечественный стек.")
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
print("Presentation saved successfully at:", pptx_path)

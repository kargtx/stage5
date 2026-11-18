# -*- coding: utf-8 -*-
import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

out_dir = r"c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\Доклад_SRE_IaC_в_АССН"
fig_dir = os.path.join(out_dir, "figures")
pptx_path = os.path.join(out_dir, "SRE_и_IaC_в_АССН_Презентация.pptx")

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
    p1.text = "Применение концепций SRE и Infrastructure as Code\nдля сопровождения АС специального назначения"
    p1.font.name = "Arial"
    p1.font.size = Pt(30)
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

# ==================== SLIDE 2: CHALLENGES OF TRADITIONAL OPS ====================
def slide_2():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "АКТУАЛЬНОСТЬ И ВЫЗОВЫ", "Проблематика сопровождения АССН в рамках классического Ops",
               "Ручные методы администрирования исчерпали себя и стали главным источником инцидентов")

    w = Inches(2.75)
    gap = Inches(0.23)
    top = Inches(2.0)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "1. ЧЕЛОВЕЧЕСКИЙ ФАКТОР", [
        ("Статистика сбоев", "Более 70% аварий в АС происходят из-за непреднамеренных ошибок администраторов."),
        ("Ручные операции", "Ручной ввод команд в консоль на боевых серверах неизбежно приводит к опечаткам."),
        ("Цена ошибки", "Сбой в АССН влечет нарушение управления критическими объектами.")
    ], accent_color=RED_ACCENT)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "2. ДРЕЙФ КОНФИГУРАЦИЙ", [
        ("Configuration Drift", "Различие версий библиотек, пакетов и настроек между полигоном и боевым контуром."),
        ("Эффект «Снежинок»", "Серверы со временем превращаются в уникальные нестандартные системы («Snowflakes»)."),
        ("Непредсказуемость", "Обновление ПО, успешно прошедшее на стенде, падает на боевом комплексе.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "3. РУТИНА (TOIL)", [
        ("Перегрузка смен", "Дежурные инженеры тратят свыше 80% времени на монотонные механические задачи."),
        ("Отсутствие роста", "Вместо развития надежности персонал занят ручной очисткой логов и перезапусками служб."),
        ("Выгорание кадров", "Низкая эффективность реагирования при возникновении нестандартных аварий.")
    ], accent_color=PURPLE)

    add_card(s, Inches(0.8) + (w + gap)*3, top, w, h, "4. ИЗОЛЯЦИЯ (AIR-GAP)", [
        ("Закрытый контур", "Невозможность использовать публичные репозитории и облачные сервисы."),
        ("Трудность обновления", "Сложный регламент доставки патчей и верификации контрольных сумм."),
        ("Требования ФСТЭК", "Строгий контроль неизменности программной среды и аттестация стендов.")
    ], accent_color=EMERALD)

# ==================== SLIDE 3: SRE PHILOSOPHY ====================
def slide_3():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "МЕТОДОЛОГИЯ SRE", "Site Reliability Engineering: надежность как программная дисциплина",
               "«SRE — это то, что происходит, когда эксплуатацию поручают инженерам-программистам»")

    top = Inches(2.0)
    w = Inches(3.7)
    gap = Inches(0.3)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "НАДЕЖНОСТЬ КАК КОД", [
        ("Программный подход", "Инфраструктура и правила эксплуатации реализуются в виде исходного кода и скриптов."),
        ("Автоматизация взамен рутины", "Если действие выполняется более двух раз, оно обязано быть автоматизировано."),
        ("Инженерный лимит", "Минимум 50% рабочего времени инженера SRE тратится на проектирование и разработку автоматики.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "МАТЕМАТИЧЕСКИЙ БАЛАНС", [
        ("Отказ от 100% надежности", "100% надежность — недостижимая и экономически бессмысленная цель."),
        ("Бюджет ошибок", "Оставшаяся доля недоступности (например, 0.05%) используется как ресурс для модернизации."),
        ("Объективный арбитраж", "Данные телеметрии, а не эмоции разработчиков, решают, можно ли накатывать релиз.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "КУЛЬТУРА BLAMELESS", [
        ("Разбор без чувства вины", "Ошибся не конкретный дежурный, а система, которая позволила человеку ошибиться."),
        ("Системные исправления", "По итогам инцидента создается автоматический барьер безопасности в пайплайне."),
        ("База знаний отказов", "Открытые постмортемы позволяют всей организации учиться на ошибках.")
    ], accent_color=EMERALD)

# ==================== SLIDE 4: SLI, SLO, ERROR BUDGET ====================
def slide_4():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "МАТЕМАТИКА НАДЕЖНОСТИ", "Триада метрик SRE: SLI, SLO и Бюджет ошибок (Error Budget)",
               "Четкие числовые ориентиры надежности и автоматические правила блокировки релизов")

    img_path = os.path.join(fig_dir, "diagram_sre_sli_slo_error_budget.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "ПРАВИЛО РАСЧЕТА", [
        ("Формула баланса", "Error Budget = 100% – SLO."),
        ("Пример расчета", "При целевом SLO = 99.95% бюджет простоя составляет ровно 21.6 минуты в месяц."),
        ("Транзакционный SLI", "Доля запросов с временем отклика < 50 мс.")
    ], accent_color=CYAN)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "ERROR BUDGET POLICY", [
        ("Защита боевого контура", "Если за месяц лимит исчерпан авариями — развертывание новых функций блокируется."),
        ("Фокус на стабильность", "Вся команда занимается только устранением причин сбоев и техдолга."),
    ], accent_color=RED_ACCENT)

# ==================== SLIDE 5: INFRASTRUCTURE AS CODE ====================
def slide_5():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ПАРАДИГМА IAC", "Infrastructure as Code: декларативное управление средой АССН",
               "Переход от ручных скриптов к декларативным спецификациям и контролю версий")

    top = Inches(2.0)
    w = Inches(3.7)
    gap = Inches(0.3)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "ДЕКЛАРАТИВНЫЙ ПОДХОД", [
        ("Целевое состояние", "Описывается ЧТО должно получиться (State), а не пошаговые команды КАК делать."),
        ("Движок оркестрации", "Terraform / OpenTofu самостоятельно определяет граф зависимостей и порядок создания."),
        ("Единый источник правды", "Репозиторий Git содержит полный слепок всех параметров инфраструктуры АССН.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "ИДЕМПОТЕНТНОСТЬ (IDEMPOTENCY)", [
        ("Гарантия результата", "Многократный повторный запуск сценария Ansible приводит к строго одинаковому состоянию."),
        ("Безопасность перезапуска", "Если конфигурация уже соответствует целевой, система не совершает лишних действий."),
        ("Устранение сбоев сети", "При разрыве связи сценарий можно безболезненно перезапустить с места прерывания.")
    ], accent_color=EMERALD)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "ИММУТАБЕЛЬНОСТЬ (IMMUTABLE)", [
        ("Золотые образы (Golden)", "Операционная система с проверенными пакетами собирается утилитой Packer один раз."),
        ("Запрет ручных правок", "Прямой доступ по SSH на боевые серверы заблокирован политиками безопасности."),
        ("Полная повторяемость", "Развертывание 100 идентичных серверов занимает столько же времени, сколько одного.")
    ], accent_color=PURPLE)

# ==================== SLIDE 6: AIR-GAPPED IAC PIPELINE ====================
def slide_6():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "АРХИТЕКТУРА И БЕЗОПАСНОСТЬ", "Пайплайн IaC в изолированном контуре АССН (Air-Gapped)",
               "Полная автономность, отсутствие внешней связности и сквозной контроль качества DevSecOps")

    img_path = os.path.join(fig_dir, "diagram_iac_airgap_pipeline.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "ДОВЕРЕННЫЙ РЕЕСТР", [
        ("Локальный контур", "Все артефакты, образы и пакеты хранятся в защищенном локальном Nexus / Harbor."),
        ("ЭЦП манифестов", "Все коммиты в Git подписываются ГОСТ-ключами разработчиков."),
    ], accent_color=CYAN)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "DEVSECOPS GATE", [
        ("Статический анализ", "Линтеры tflint и ansible-lint ловят синтаксические и логические ошибки до деплоя."),
        ("Сканирование уязвимостей", "Утилита Trivy проверяет компоненты по базам БДУ ФСТЭК России."),
    ], accent_color=EMERALD)

# ==================== SLIDE 7: ZERO-DOWNTIME DEPLOYMENT ====================
def slide_7():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "НЕПРЕРЫВНОСТЬ ФУНКЦИОНИРОВАНИЯ", "Развертывание без простоя (Zero-Downtime Blue-Green & Canary)",
               "Обновление боевых комплексов 24/7 с гарантией мгновенного отката при сбоях")

    img_path = os.path.join(fig_dir, "diagram_zero_downtime_deployment.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "МЕТОД BLUE-GREEN", [
        ("Параллельные контуры", "Полная копия боевой среды разворачивается рядом кодом IaC."),
        ("Мгновенный свитч", "Переключение трафика балансировщиком за миллисекунды."),
        ("Нулевой простой", "Пользователи не замечают технологического окна обновления."),
    ], accent_color=CYAN)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "ПАТТЕРН EXPAND-CONTRACT", [
        ("Миграция СУБД", "Схема базы данных расширяется с обратной совместимостью (Expand)."),
        ("Очистка (Contract)", "Удаление старых полей только ПОСЛЕ успешного перевода 100% трафика."),
    ], accent_color=AMBER)

# ==================== SLIDE 8: TOIL REDUCTION & INCIDENTS ====================
def slide_8():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "УПРАВЛЕНИЕ НАДЕЖНОСТЬЮ", "Борьба с рутиной (Eliminating Toil) и замкнутый цикл инцидентов",
               "Трансформация опыта устранения аварий в автоматизированные сценарии надежности")

    img_path = os.path.join(fig_dir, "diagram_incident_lifecycle_sre.png")
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0.8), Inches(1.9), width=Inches(8.2))

    add_card(s, Inches(9.2), Inches(1.9), Inches(3.3), Inches(2.4), "ЧТО ТАКОЕ TOIL?", [
        ("Признаки рутины", "Ручная, повторяемая работа, не требующая инноваций, масштабируемая линейно с сервисом."),
        ("Норматив SRE", "Жесткое ограничение: рутина обязана составлять менее 50% рабочего времени."),
    ], accent_color=PURPLE)

    add_card(s, Inches(9.2), Inches(4.5), Inches(3.3), Inches(2.4), "АВТО-СМЯГЧЕНИЕ", [
        ("Self-Healing", "Автоматический перезапуск зависших подов и failover кластеров баз данных."),
        ("Защита дежурного", "Инженер привлекается только при нетиповых авариях."),
    ], accent_color=EMERALD)

# ==================== SLIDE 9: CHAOS ENGINEERING IN CRITICAL SYSTEMS ====================
def slide_9():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ИНЖЕНЕРИЯ ХАОСА", "Стресс-тестирование и учения по надежности в АССН",
               "Проверка механизмов автоматического восстановления в боевых условиях до реальной аварии")

    top = Inches(2.0)
    w = Inches(3.7)
    gap = Inches(0.3)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "КОНЦЕПЦИЯ CHAOS ENGINEERING", [
        ("Проактивная провокация", "Искусственное внедрение сбоев в контролируемое время в присутствии дежурной смены."),
        ("Устранение скрытых багов", "Выявление неработающих конфигураций резервирования, которые никто не проверял годами."),
        ("GameDays", "Регулярные учения инженерного состава по отработке нештатных сценариев.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "ТИПОВЫЕ ТЕСТОВЫЕ СЦЕНАРИИ", [
        ("Сетевой раскол (Split-Brain)", "Искусственная изоляция одного из ЦОД и проверка кворума Raft/Etcd."),
        ("Внезапная смерть узла", "Резкое отключение мастер-сервера СУБД (Postgres Pro / Patroni) под нагрузкой."),
        ("Деградация дисковой шины", "Искусственное введение задержек в 100 мс на дисковые операции через tc/netem."),
        ("Исчерпание ресурсов", "Стресс-тестирование памяти (OOM-killer) и насыщение пулов соединений.")
    ], accent_color=RED_ACCENT)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "РЕЗУЛЬТАТЫ ДЛЯ АССН", [
        ("Подтвержденный RTO", "Доказанное время автоматического восстановления менее 10–15 секунд."),
        ("Уверенность в автоматике", "Инженеры точно знают, как поведет себя комплекс при реальном отказе стойки или канала."),
        ("Обновление IaC-кода", "Любой выявленный нюанс сразу фиксируется в виде изменений в манифестах инфраструктуры.")
    ], accent_color=EMERALD)

# ==================== SLIDE 10: RUSSIAN TRUSTED TECH STACK ====================
def slide_10():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "ИМПОРТОЗАМЕЩЕНИЕ И КОМПЛАЕНС", "Реализация SRE и IaC на отечественном доверенном стеке",
               "Совместимость с требованиями регуляторов ФСТЭК России и ГОСТ")

    top = Inches(2.0)
    w = Inches(2.75)
    gap = Inches(0.23)
    h = Inches(4.8)

    add_card(s, Inches(0.8), top, w, h, "ОТЕЧЕСТВЕННЫЕ ОС", [
        ("Astra Linux SE", "Режим замкнутой программной среды, мандатный контроль доступа (МКД)."),
        ("РЕД ОС / Альт 8 СП", "Сертифицированные дистрибутивы общего и специального назначения."),
        ("Интеграция с IaC", "Управление через доверенные модули Ansible без агентов (Agentless SSH).")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "СРЕДЫ ВИРТУАЛИЗАЦИИ", [
        ("ПК СВ «Брест»", "Отечественный защищенный комплекс виртуализации на базе Astra Linux."),
        ("zVirt / РЕД Вирт.", "Управление пулами вычислительных ресурсов спецназначения."),
        ("OpenTofu Провайдеры", "Декларативное развертывание виртуальных машин через API гипервизоров.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8) + (w + gap)*2, top, w, h, "КОНТЕЙНЕРИЗАЦИЯ", [
        ("Защищенный Kube", "Отечественные платформы управления контейнерами (Deckhouse, Штурвал)."),
        ("Локальные реестры", "Изолированное хранение образов в доверенных средах."),
        ("Контроль целостности", "Контроль контрольных сумм контейнеров в соответствии с требованиями ФСТЭК.")
    ], accent_color=PURPLE)

    add_card(s, Inches(0.8) + (w + gap)*3, top, w, h, "СТАНДАРТЫ И НОРМЫ", [
        ("ГОСТ Р 56939-2016", "Защита информации. Безопасная разработка программного обеспечения (DevSecOps)."),
        ("ГОСТ Р 59548-2021", "Автоматизированные системы в защищенном исполнении."),
        ("Аттестация контуров", "IaC-манифесты выступают юридическим доказательством эталонной настройки.")
    ], accent_color=EMERALD)

# ==================== SLIDE 11: SUMMARY COMPARISON MATRIX ====================
def slide_11():
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s)
    add_header(s, "СРАВНИТЕЛЬНЫЙ АНАЛИЗ", "Сравнительная матрица: Традиционный Ops против SRE + IaC",
               "Количественная и качественная оценка эффективности перехода на современные практики")

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

    headers = ["Параметр сопровождения", "Традиционный подход (Ops)", "Современный подход (SRE + IaC)", "Эффект для АССН"]
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
        ("Развертывание узла / стенда", "Ручная установка: 2–5 дней", "Автоматизация IaC: 10–15 минут", "Ускорение в 30+ раз"),
        ("Контроль конфигураций", "Дрейф настроек, уникальные «снежинки»", "Git Single Source of Truth, аудит версий", "100% повторяемость"),
        ("Человеческий фактор", "Высокий (>70% причин сбоев)", "Минимальный (код валидируется линтерами)", "Снижение аварий на 80%"),
        ("Время восстановления (MTTR)", "Часы (поиск причины вслепую)", "Минуты (Self-healing, авто-failover)", "Соответствие RTO ≈ 0"),
        ("Борьба с рутиной (Toil)", "Рутина занимает 80–90% времени", "Рутина жестко ограничена < 50%", "Инженерное развитие")
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
               "Новый технологический стандарт надежности комплексов специального назначения")

    top = Inches(2.0)
    w = Inches(5.6)
    gap = Inches(0.5)
    h = Inches(2.2)

    add_card(s, Inches(0.8), top, w, h, "1. ИНЖЕНЕРНЫЙ ПОДХОД", [
        ("Надежность как продукт", "SRE переводит эксплуатацию из режима хаотичного тушения пожаров в строгую инженерную деятельность с четкими метриками SLI/SLO."),
        ("Защита от перегрузок", "Бюджет ошибок объективно регулирует темп релизов и защищает систему от дестабилизации.")
    ], accent_color=CYAN)

    add_card(s, Inches(0.8) + w + gap, top, w, h, "2. ИСКЛЮЧЕНИЕ СНЕЖИНОК ЧЕРЕЗ IAC", [
        ("Абсолютная идентичность", "Декларативное описание инфраструктуры полностью устраняет расхождение между полигонами тестирования и боевыми контурами."),
        ("Скорость масштабирования", "Ввод резервных вычислительных мощностей при авариях сокращается с суток до десятков минут.")
    ], accent_color=AMBER)

    add_card(s, Inches(0.8), top + h + Inches(0.3), w, h, "3. БЕЗОПАСНОСТЬ И КОМПЛАЕНС", [
        ("DevSecOps в изолированных сетях", "Успешная реализация в полностью закрытых контурах (Air-Gapped) с локальными реестрами и ГОСТ ЭЦП."),
        ("Выполнение требований ФСТЭК", "Инфраструктурный код является самодокументируемым юридическим доказательством корректности настроек безопасности.")
    ], accent_color=EMERALD)

    add_card(s, Inches(0.8) + w + gap, top + h + Inches(0.3), w, h, "4. ГОТОВНОСТЬ К СУВЕРЕНИТЕТУ", [
        ("Отечественный стек", "Стек успешно апробирован на базе Astra Linux SE, СУБД Postgres Pro Enterprise и российских гипервизоров."),
        ("Экономический эффект", "Снижение совокупной стоимости владения (TCO) и многократный рост надежности специальных систем.")
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
print("SRE Presentation saved successfully at:", pptx_path)

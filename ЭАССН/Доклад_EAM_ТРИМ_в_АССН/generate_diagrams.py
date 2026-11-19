# -*- coding: utf-8 -*-
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches

fig_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "figures")
os.makedirs(fig_dir, exist_ok=True)

# 1. Asynchronous Packet Replication Architecture
def make_replication_diagram():
    fig, ax = plt.subplots(figsize=(12, 7.2), dpi=300)
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 8.0)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(6, 7.6, "Архитектура асинхронной пакетной репликации EAM TRIM (Борт <—> Берег)",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    # 3 Main Zones
    # Zone 1: Бортовой сегмент
    bort_box = patches.FancyBboxPatch((0.6, 1.2), 3.8, 5.8, boxstyle="round,pad=0.1,rounding_size=0.15",
                                     facecolor='#E0F2FE', edgecolor='#0284C7', linewidth=1.5)
    ax.add_patch(bort_box)
    ax.text(2.5, 6.7, "АВТОНОМНЫЙ СПЕЦОБЪЕКТ\n(Ледокол / ПАТЭС / Платформа)",
            fontsize=10, fontweight='bold', ha='center', color='#0369A1', fontfamily='sans-serif')

    bort_items = [
        ("АРМ судовых механиков", "Фиксация дефектов, наряды, списание ЗИП"),
        ("Локальная СУБД TRIM", "Автономная база данных (Postgres Pro / MS SQL)"),
        ("Очередь изменений (Outbox)", "Накопление дифференциальных пакетов (XML)"),
        ("Шифрование ГОСТ", "Подписание и сжатие пакета перед сеансом")
    ]
    y_pos = 5.6
    for title, desc in bort_items:
        r = patches.FancyBboxPatch((0.9, y_pos - 0.45), 3.2, 0.8, boxstyle="round,pad=0.08,rounding_size=0.1",
                                   facecolor='#FFFFFF', edgecolor='#BAE6FD', linewidth=1.0)
        ax.add_patch(r)
        ax.text(2.5, y_pos - 0.1, title, fontsize=8.5, fontweight='bold', ha='center', color='#0F172A')
        ax.text(2.5, y_pos - 0.32, desc, fontsize=7.2, ha='center', color='#64748B')
        y_pos -= 1.05

    # Zone 2: Канал связи
    chan_box = patches.FancyBboxPatch((4.7, 2.0), 2.6, 4.2, boxstyle="round,pad=0.1,rounding_size=0.15",
                                     facecolor='#FEF3C7', edgecolor='#D97706', linewidth=1.5)
    ax.add_patch(chan_box)
    ax.text(6.0, 5.8, "УЗКОПОЛОСНЫЙ КАНАЛ\n(Периодическая связь)",
            fontsize=9.5, fontweight='bold', ha='center', color='#92400E', fontfamily='sans-serif')

    chan_items = ["Спутниковая связь\n(Гонец / Экспресс)", "КВ/УКВ радиолинии\nморской связи", "Режим радиомолчания\n(полная буферизация)", "Контроль целостности\n(CRC32 + ГОСТ ЭЦП)"]
    y_pos = 4.8
    for ci in chan_items:
        r = patches.FancyBboxPatch((4.9, y_pos - 0.4), 2.2, 0.65, boxstyle="round,pad=0.06,rounding_size=0.08",
                                   facecolor='#FFFFFF', edgecolor='#FDE68A', linewidth=1.0)
        ax.add_patch(r)
        ax.text(6.0, y_pos - 0.12, ci, fontsize=7.5, ha='center', color='#1E293B', fontweight='bold')
        y_pos -= 0.85

    # Zone 3: Береговой центр
    bereg_box = patches.FancyBboxPatch((7.6, 1.2), 3.8, 5.8, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor='#DCFCE7', edgecolor='#16A34A', linewidth=1.5)
    ax.add_patch(bereg_box)
    ax.text(9.5, 6.7, "БЕРЕГОВОЙ ЦЕНТР\n(ФГУП «Атомфлот» / Штаб флота)",
            fontsize=10, fontweight='bold', ha='center', color='#15803D', fontfamily='sans-serif')

    bereg_items = [
        ("Шлюз приема репликаций", "Валидация пакетов, разрешение коллизий"),
        ("Центральное DWH TRIM", "Консолидированная БД всего флота"),
        ("Модуль надежности RCM", "Анализ дефектов и наработки по флоту"),
        ("АРМ Службы ТОиР", "Централизованное снабжение и докование")
    ]
    y_pos = 5.6
    for title, desc in bereg_items:
        r = patches.FancyBboxPatch((7.9, y_pos - 0.45), 3.2, 0.8, boxstyle="round,pad=0.08,rounding_size=0.1",
                                   facecolor='#FFFFFF', edgecolor='#BBF7D0', linewidth=1.0)
        ax.add_patch(r)
        ax.text(9.5, y_pos - 0.1, title, fontsize=8.5, fontweight='bold', ha='center', color='#0F172A')
        ax.text(9.5, y_pos - 0.32, desc, fontsize=7.2, ha='center', color='#64748B')
        y_pos -= 1.05

    # Bidirectional arrows
    arrow_props = dict(facecolor='#0284C7', edgecolor='#0284C7', width=1.5, headwidth=6, headlength=6)
    ax.annotate('', xy=(4.7, 4.2), xytext=(4.4, 4.2), arrowprops=arrow_props)
    ax.annotate('', xy=(7.6, 4.2), xytext=(7.3, 4.2), arrowprops=arrow_props)
    ax.annotate('', xy=(4.4, 3.6), xytext=(4.7, 3.6), arrowprops=dict(facecolor='#16A34A', edgecolor='#16A34A', width=1.5, headwidth=6, headlength=6))
    ax.annotate('', xy=(7.3, 3.6), xytext=(7.6, 3.6), arrowprops=dict(facecolor='#16A34A', edgecolor='#16A34A', width=1.5, headwidth=6, headlength=6))

    # Bottom note
    bot_r = patches.FancyBboxPatch((0.6, 0.3), 10.8, 0.65, boxstyle="round,pad=0.06,rounding_size=0.08",
                                  facecolor='#F1F5F9', edgecolor='#94A3B8', linewidth=1.0)
    ax.add_patch(bot_r)
    ax.text(6.0, 0.62, "Ключевая особенность: отсутствие онлайн-зависимости, гарантированная доставка изменений при сеансах связи",
            fontsize=8.5, style='italic', ha='center', color='#334155')

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_trim_replication_architecture.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

# 2. RCM Methodology Workflow in TRIM
def make_rcm_workflow_diagram():
    fig, ax = plt.subplots(figsize=(11, 6.8), dpi=300)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 7.5)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(5.5, 7.1, "Контур управления надежностью RCM в EAM-комплексе ТРИМ",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    steps = [
        ("1. Сбор параметров", 0.6, 4.3, 2.2, 2.2, "#DBEAFE", "#1E40AF",
         ["• Фактическая наработка", "• Число пусков/циклов", "• Вибродиагностика", "• Температурные тренды"]),
        ("2. FMECA Анализ", 2.9, 4.3, 2.2, 2.2, "#FEE2E2", "#991B1B",
         ["• Виды отказов", "• Критичность последствий", "• Вероятность сбоя", "• Индекс RPN (риска)"]),
        ("3. Расчет ресурса", 5.2, 4.3, 2.2, 2.2, "#FEF3C7", "#92400E",
         ["• P-F интервал отказа", "• Остаточный ресурс", "• Скорость деградации", "• Оценка надежности"]),
        ("4. ТО по состоянию", 7.5, 4.3, 2.4, 2.2, "#DCFCE7", "#166534",
         ["• Авто-наряд на ТОиР", "• Резервирование ЗИП", "• Предотвращение аварии", "• Минимизация простоя"]),
    ]

    for title, x, y, w, h, bg_col, text_col, points in steps:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor=bg_col, edgecolor='#94A3B8', linewidth=1.2)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h - 0.28, title, fontsize=9.0, fontweight='bold', ha='center', color=text_col, fontfamily='sans-serif')
        curr_y = y + h - 0.65
        for pt in points:
            ax.text(x + 0.12, curr_y, pt, fontsize=7.5, color='#334155', fontfamily='sans-serif', va='top')
            curr_y -= 0.35

    arrow_props = dict(facecolor='#475569', edgecolor='#475569', width=1.5, headwidth=6, headlength=6)
    ax.annotate('', xy=(2.9, 5.4), xytext=(2.7, 5.4), arrowprops=arrow_props)
    ax.annotate('', xy=(5.2, 5.4), xytext=(5.0, 5.4), arrowprops=arrow_props)
    ax.annotate('', xy=(7.5, 5.4), xytext=(7.3, 5.4), arrowprops=arrow_props)

    # Bottom comparison container
    bot_rect = patches.FancyBboxPatch((0.6, 0.8), 9.8, 2.8, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor='#FFFFFF', edgecolor='#CBD5E1', linewidth=1.2)
    ax.add_patch(bot_rect)
    ax.text(0.9, 3.2, "Сравнение стратегий эксплуатации критических агрегатов в спецсистемах:", fontsize=9.5, fontweight='bold', color='#0F172A')

    ax.text(0.9, 2.7, "1. Реактивный ремонт (по отказу):", fontsize=8.5, fontweight='bold', color='#B91C1C')
    ax.text(0.9, 2.4, "   Категорически недопустим на атомных объектах и спецсудах — ведет к потере хода и авариям 1-й категории.", fontsize=7.8, color='#334155')

    ax.text(0.9, 2.0, "2. Планово-предупредительный ремонт (ППР по регламенту часов):", fontsize=8.5, fontweight='bold', color='#D97706')
    ax.text(0.9, 1.7, "   Замена исправных узлов «по часам» приводит к перерасходу дефицитных ЗИП и внесению дефектов при сборке.", fontsize=7.8, color='#334155')

    ax.text(0.9, 1.3, "3. Обслуживание по техническому состоянию (RCM / CBM в ПК ТРИМ):", fontsize=8.5, fontweight='bold', color='#15803D')
    ax.text(0.9, 1.0, "   Ремонт выполняется точно до перехода дефекта в отказ. Экономия ЗИП 25–40% при 100% коэффициенте готовности.", fontsize=7.8, color='#334155')

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_trim_rcm_workflow.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

# 3. Structure of Electronic Logbook in TRIM
def make_logbook_diagram():
    fig, ax = plt.subplots(figsize=(11, 6.8), dpi=300)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 7.5)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(5.5, 7.1, "Структура электронного паспорта и формуляра оборудования в ТРИМ",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    cards = [
        ("ИЕРАРХИЯ ОБОРУДОВАНИЯ", 0.6, 3.8, 4.6, 2.8, "#E0F2FE", "#0369A1",
         ["Древовидная структура спецобъекта:", "• Главная энергетическая установка (ГЭУ)", "• Вспомогательные турбогенераторы", "• Насосные агрегаты 1-го и 2-го контуров", "• Система управления движением"]),
        ("ЖУРНАЛ НАРАБОТКИ И СОБЫТИЙ", 5.8, 3.8, 4.6, 2.8, "#DCFCE7", "#166534",
         ["Юридически значимый цифровой след:", "• Наработка в моточасах от АСУ ТП", "• Журнал дефектов и замечаний вахты", "• Акты освидетельствования РМРС / РКО", "• История ремонтов и замен узлов"]),
        ("ТЕХНОЛОГИЧЕСКИЕ КАРТЫ", 0.6, 0.7, 4.6, 2.7, "#FEF3C7", "#92400E",
         ["Регламенты и нормы ТОиР:", "• Карты смазки и дефектации узлов", "• Нормативы трудоемкости операций", "• Требования безопасности и допуски", "• Инструкции по консервации агрегатов"]),
        ("СКЛАДСКОЙ УЧЕТ ЗИП", 5.8, 0.7, 4.6, 2.7, "#F3E8FF", "#7E22CE",
         ["Управление судовым запасом:", "• Неснижаемый аварийный запас ЗИП", "• Контроль сроков годности РТИ и масел", "• Местоположение в судовых выгородках", "• Авто-заказ при подходе к минимальной квоте"]),
    ]

    for title, x, y, w, h, bg_col, text_col, points in cards:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor=bg_col, edgecolor='#94A3B8', linewidth=1.2)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h - 0.28, title, fontsize=9.0, fontweight='bold', ha='center', color=text_col, fontfamily='sans-serif')
        curr_y = y + h - 0.65
        for pt in points:
            ax.text(x + 0.15, curr_y, pt, fontsize=8.0, color='#1E293B', fontfamily='sans-serif', va='top')
            curr_y -= 0.4

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_trim_electronic_logbook.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

# 4. Fault Tolerance in Isolated Maritime Environment
def make_fault_tolerance_diagram():
    fig, ax = plt.subplots(figsize=(11, 6.8), dpi=300)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 7.5)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(5.5, 7.1, "Обеспечение живучести информационной системы TRIM на спецобъектах",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    boxes = [
        ("Аварийное обесточивание (Blackout)", 0.6, 4.2, 4.6, 2.4, "#FEE2E2", "#991B1B",
         ["• Автономные ИБП двойного преобразования", "• Форсированный сброс буферов транзакций WAL", "• Защита от повреждения табличных пространств", "• Авто-старт служб при пуске аварийного дизеля"]),
        ("Виброустойчивость и ударопрочность", 5.8, 4.2, 4.6, 2.4, "#FEF3C7", "#92400E",
         ["• Промышленные серверы в морском исполнении", "• Отказ от вращающихся HDD в пользу NVMe SSD", "• Амортизационные серверные шкафы", "• Работа в условиях арктических температур"]),
        ("Локальный кластер высокой доступности", 0.6, 1.2, 4.6, 2.4, "#DCFCE7", "#166534",
         ["• Дублирование сервера: Основной <-> Резервный", "• Синхронная репликация в пределах судовой сети", "• Автоматический Failover за < 10 секунд", "• Размещение серверов в разных отсеках объекта"]),
        ("Защита от несанкционированного доступа", 5.8, 1.2, 4.6, 2.4, "#DBEAFE", "#1E40AF",
         ["• Ролевая модель доступа по служебным обязанностям", "• Электронная подпись записей судового журнала", "• Невозможность задним числом скрыть аварию", "• Полное соответствие требованиям РМРС и ФСТЭК"]),
    ]

    for title, x, y, w, h, bg_col, text_col, points in boxes:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor=bg_col, edgecolor='#94A3B8', linewidth=1.2)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h - 0.28, title, fontsize=9.0, fontweight='bold', ha='center', color=text_col, fontfamily='sans-serif')
        curr_y = y + h - 0.65
        for pt in points:
            ax.text(x + 0.15, curr_y, pt, fontsize=8.0, color='#1E293B', fontfamily='sans-serif', va='top')
            curr_y -= 0.4

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_trim_fault_tolerance.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

if __name__ == '__main__':
    make_replication_diagram()
    make_rcm_workflow_diagram()
    make_logbook_diagram()
    make_fault_tolerance_diagram()
    print("All TRIM diagrams generated successfully!")

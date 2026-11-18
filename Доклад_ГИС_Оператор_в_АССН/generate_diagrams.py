# -*- coding: utf-8 -*-
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches

fig_dir = r"c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\Доклад_ГИС_Оператор_в_АССН\figures"
os.makedirs(fig_dir, exist_ok=True)

# 1. End-to-End Architecture of GIS Operator in Secure Military / Critical Network
def make_architecture_diagram():
    fig, ax = plt.subplots(figsize=(12, 7.5), dpi=300)
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 8.5)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(6, 8.1, "Сквозная архитектура защищенной геоинформационной системы ГИС «Оператор»",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    layers = [
        ("УРОВЕНЬ ХРАНЕНИЯ ГЕОДАННЫХ (GIS Server & Spatial DB)", 6.7, 1.0, "#E2E8F0", "#1E293B",
         ["Пространственная СУБД (Postgres Pro / Линтер-ВС)", "ГИС Сервер (КБ Панорама)", "Банк цифровых карт (SXF / SITX / TIFF)", "Матрицы высот и рельефа (MTW / MTL)"]),
        ("УРОВЕНЬ БЕЗОПАСНОСТИ И МАНДАТНОЙ ФИЛЬТРАЦИИ (Security Gateway)", 5.1, 1.0, "#FEE2E2", "#991B1B",
         ["Мандатный контроль доступа (МКД Astra Linux)", "Фильтр слоев по грифу секретности", "Контроль целостности картосновы (ГОСТ)", "Защищенный шлюз межсетевого экрана"]),
        ("УРОВЕНЬ ТРАНСПОРТА И РАСПРЕДЕЛЕНИЯ (Secure Field Comms)", 3.5, 1.0, "#FEF3C7", "#92400E",
         ["КАС ТР (Комплексы аппаратных связи)", "Низкоскоростные радиомодемы передачи данных", "Шифрование каналов (ГОСТ криптография)", "Протоколы сжатия и докачки фрагментов"]),
        ("КЛИЕНТСКИЙ УРОВЕНЬ ОТОБРАЖЕНИЯ (Operator Workstations)", 1.9, 1.0, "#DCFCE7", "#166534",
         ["АРМ Командира / Оперативного дежурного", "Защищенные полевые планшеты (Astra Linux)", "Бортовые навигационные ЭВМ спецтехники", "Слой оперативной обстановки (КЭО условные знаки)"]),
    ]

    for title, y, h, bg_col, text_col, items in layers:
        rect = patches.FancyBboxPatch((0.5, y), 11, h, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor=bg_col, edgecolor='#94A3B8', linewidth=1.2)
        ax.add_patch(rect)
        ax.text(0.7, y + h - 0.25, title, fontsize=9.5, fontweight='bold', color=text_col, fontfamily='sans-serif')

        dx = 10.6 / len(items)
        for i, item in enumerate(items):
            bx = 0.7 + i * dx
            b_rect = patches.FancyBboxPatch((bx, y + 0.12), dx - 0.2, h - 0.45,
                                           boxstyle="round,pad=0.08,rounding_size=0.1",
                                           facecolor='#FFFFFF', edgecolor='#CBD5E1', linewidth=1.0)
            ax.add_patch(b_rect)
            ax.text(bx + (dx - 0.2)/2, y + 0.38, item, fontsize=8.0, ha='center', va='center',
                    color='#334155', fontweight='bold', fontfamily='sans-serif', wrap=True)

    arrow_props = dict(facecolor='#0284C7', edgecolor='#0284C7', width=1.5, headwidth=6, headlength=6)
    for y_arrow in [6.7, 5.1, 3.5]:
        ax.annotate('', xy=(6.0, y_arrow), xytext=(6.0, y_arrow + 0.05), arrowprops=arrow_props)

    ax.text(6, 0.4, "Сертификация МО РФ по 2-му уровню контроля отсутствия НДВ; функционирование в изолированных сетях",
            fontsize=8.5, style='italic', ha='center', color='#64748B')

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_gis_operator_architecture.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

# 2. Tile Pyramid & Offline Cache Architecture
def make_tile_pyramid_diagram():
    fig, ax = plt.subplots(figsize=(11, 6.8), dpi=300)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 7.5)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(5.5, 7.1, "Пирамида пространственных тайлов и схема локального кэширования карт",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    # Visual Pyramid
    levels = [
        ("Масштаб 1:1 000 000 (Обзорный уровень)", 4.0, 5.5, 3.0, 0.7, "#E0F2FE", "#0369A1", "1 тайл = 256×256 px"),
        ("Масштаб 1:200 000 (Оперативный уровень)", 3.0, 4.4, 5.0, 0.7, "#BAE6FD", "#0284C7", "Сетка квадратов карты"),
        ("Масштаб 1:50 000 (Тактический уровень)", 2.0, 3.3, 7.0, 0.7, "#7DD3FC", "#0369A1", "Детализация рельефа, высот, дорог"),
        ("Масштаб 1:10 000 (Детальный городской / план)", 1.0, 2.2, 9.0, 0.7, "#38BDF8", "#075985", "Здания, ориентиры, инфраструктура"),
    ]

    for title, x, y, w, h, bg_col, text_col, desc in levels:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.08,rounding_size=0.1",
                                      facecolor=bg_col, edgecolor='#0284C7', linewidth=1.2)
        ax.add_patch(rect)
        ax.text(5.5, y + 0.38, title, fontsize=9.0, fontweight='bold', ha='center', color='#0F172A')
        ax.text(5.5, y + 0.15, desc, fontsize=7.5, ha='center', color='#334155')

    # Bottom explanation
    bot_rect = patches.FancyBboxPatch((0.6, 0.6), 9.8, 1.3, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor='#FFFFFF', edgecolor='#CBD5E1', linewidth=1.2)
    ax.add_patch(bot_rect)
    ax.text(0.9, 1.5, "Особенности кэширования для работы полевых АРМ без связи (Off-Line Cache):", fontsize=9.5, fontweight='bold', color='#0F172A')
    ax.text(0.9, 1.15, "• Вся пространственная база района боевого дежурства предварительно зашивается в зашифрованный контейнер на SSD.", fontsize=8.0, color='#334155')
    ax.text(0.9, 0.85, "• При масштабировании ГИС не запрашивает сервер, а мгновенно подтягивает нужный тайловый слой из ОЗУ/кэша.", fontsize=8.0, color='#334155')

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_gis_tile_pyramid_cache.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

# 3. Real-Time Tactical Picture Pipeline
def make_tactical_realtime_diagram():
    fig, ax = plt.subplots(figsize=(11, 6.8), dpi=300)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 7.5)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(5.5, 7.1, "Пайплайн отображения динамической оперативно-тактической обстановки",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    steps = [
        ("1. Источники телеметрии", 0.6, 4.3, 2.2, 2.2, "#DBEAFE", "#1E40AF",
         ["• Спецтехника (ГЛОНАСС)", "• БПЛА / Оптико-электроника", "• Радиолокационные посты", "• Донесения с передовых АРМ"]),
        ("2. Шлюз приема данных", 2.9, 4.3, 2.2, 2.2, "#FEF3C7", "#92400E",
         ["• Радиолинии и модемы", "• Контроль подлинности", "• Очередь сообщений", "• Декодер форматов"]),
        ("3. Ядро ГИС «Оператор»", 5.2, 4.3, 2.2, 2.2, "#DCFCE7", "#166534",
         ["• Классификатор КЭО", "• Условные знаки ГОСТ", "• Пространственная привязка", "• Рендеринг траекторий"]),
        ("4. Экран командующего", 7.5, 4.3, 2.4, 2.2, "#F3E8FF", "#7E22CE",
         ["• Тактический слой 24/7", "• Время задержки < 1 сек", "• Расчет зон поражения", "• Поддержка решений"]),
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

    # Bottom details container
    bot_rect = patches.FancyBboxPatch((0.6, 0.8), 9.8, 2.8, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor='#FFFFFF', edgecolor='#CBD5E1', linewidth=1.2)
    ax.add_patch(bot_rect)
    ax.text(0.9, 3.2, "Требования к производительности картографического ядра в реальном времени:", fontsize=9.5, fontweight='bold', color='#0F172A')
    ax.text(0.9, 2.7, "• Высокая частота обновления: отображение до 5 000 динамических подвижных меток с частотой перерисовки от 25 кадров/с.", fontsize=8.0, color='#334155')
    ax.text(0.9, 2.2, "• Отсутствие графических артефактов: аппаратное ускорение OpenGL на отечественных графических процессорах под Astra Linux.", fontsize=8.0, color='#334155')
    ax.text(0.9, 1.7, "• Векторная топологическая корректность: автоматический контроль самопересечений зон поражения и разграничительных линий.", fontsize=8.0, color='#334155')
    ax.text(0.9, 1.2, "• Устойчивость к перерывам связи: экстраполяция координат объектов по курсу и скорости при временной потере радиосигнала.", fontsize=8.0, color='#334155')

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_gis_tactical_realtime.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

# 4. Mandatory Access Control (MAC) in GIS Operator
def make_security_levels_diagram():
    fig, ax = plt.subplots(figsize=(11, 6.8), dpi=300)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 7.5)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(5.5, 7.1, "Мандатное разграничение доступа (МКД) к картографическим слоям",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    boxes = [
        ("Уровень 0: Несекретно / Открыто", 0.6, 4.3, 4.6, 2.3, "#DCFCE7", "#166534",
         ["• Общегеографическая цифровая основа", "• Базовый рельеф и гидрография", "• Доступно всем категориям пользователей", "• Нулевой гриф секретности"]),
        ("Уровень 1: ДСП (Для служебного пользования)", 5.8, 4.3, 4.6, 2.3, "#DBEAFE", "#1E40AF",
         ["• Дорожная сеть с мостами и грузоподъемностью", "• Объекты критической инфраструктуры", "• Служебные позывные и частоты связи", "• Доступ по служебному удостоверению"]),
        ("Уровень 2: Секретно", 0.6, 1.2, 4.6, 2.4, "#FEF3C7", "#92400E",
         ["• Слой дислокации подразделений и складов", "• Зоны досягаемости средств ПВО и РЭБ", "• Тактические замыслы операций", "• Требуется форма допуска №2"]),
        ("Уровень 3: Совершенно секретно", 5.8, 1.2, 4.6, 2.4, "#FEE2E2", "#991B1B",
         ["• Стратегические целеуказания и маршруты", "• Координаты защищенных пунктов управления", "• Специальные объекты высшей категории", "• Требуется форма допуска №1 + МКД ячейки"]),
    ]

    for title, x, y, w, h, bg_col, text_col, points in boxes:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor=bg_col, edgecolor='#94A3B8', linewidth=1.2)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h - 0.28, title, fontsize=9.0, fontweight='bold', ha='center', color=text_col, fontfamily='sans-serif')
        curr_y = y + h - 0.65
        for pt in points:
            ax.text(x + 0.15, curr_y, pt, fontsize=8.0, color='#1E293B', fontfamily='sans-serif', va='top')
            curr_y -= 0.38

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_gis_security_levels.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

if __name__ == '__main__':
    make_architecture_diagram()
    make_tile_pyramid_diagram()
    make_tactical_realtime_diagram()
    make_security_levels_diagram()
    print("All GIS Operator diagrams generated successfully!")

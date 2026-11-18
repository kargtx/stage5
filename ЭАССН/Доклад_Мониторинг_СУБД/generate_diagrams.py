# -*- coding: utf-8 -*-
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches

fig_dir = r"c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\Доклад_Мониторинг_СУБД\figures"
os.makedirs(fig_dir, exist_ok=True)

# 1. End-to-End Monitoring & Audit Architecture Diagram
def make_architecture_diagram():
    fig, ax = plt.subplots(figsize=(12, 7.5), dpi=300)
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 8.5)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(6, 8.1, "Комплексная архитектура мониторинга, аудита и диагностики СУБД реального времени",
            fontsize=13.5, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    layers = [
        ("УРОВЕНЬ ЯДРА СУБД И ИНТРОСПЕКЦИИ (Database Kernel & Telemetry)", 6.7, 1.0, "#E2E8F0", "#1E293B",
         ["SGA (Buffer Cache, Redo, Shared Pool)", "ASH (Active Session History Buffer)", "Словарь V$ и AWR Snapshot", "Unified Audit / FGA Engine"]),
        ("УРОВЕНЬ СБОРА И ЭКСПОРТА (Telemetry Ingestion & Agents)", 5.1, 1.0, "#DBEAFE", "#1E40AF",
         ["DB Exporter (Prometheus SQL Collector)", "Log Shipper (Vector / Filebeat)", "Audit Streamer (Syslog / FGA)", "OS Node Exporter (CPU, Disk I/O)"]),
        ("УРОВЕНЬ ХРАНЕНИЯ И ОБРАБОТКИ (Time-Series & Event Storage)", 3.5, 1.0, "#DCFCE7", "#166534",
         ["TSDB (Prometheus / VictoriaMetrics)", "Log Storage (OpenSearch / ClickHouse)", "AWR Historical Warehouse", "Репозиторий инцидентов ИБ"]),
        ("УРОВЕНЬ ОТОБРАЖЕНИЯ И РЕАГИРОВАНИЯ (Observability, Security & Action)", 1.9, 1.0, "#FEF3C7", "#92400E",
         ["Grafana Real-Time Dashboards", "Alertmanager (Предиктивные оповещения)", "SIEM (Контроль безопасности и прав)", "APM & Автоматические сценарии"]),
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
            ax.text(bx + (dx - 0.2)/2, y + 0.38, item, fontsize=8.5, ha='center', va='center',
                    color='#334155', fontweight='bold', fontfamily='sans-serif', wrap=True)

    # Inter-layer connectors (Arrows)
    arrow_props = dict(facecolor='#0284C7', edgecolor='#0284C7', width=1.5, headwidth=6, headlength=6)
    for y_arrow in [6.7, 5.1, 3.5]:
        ax.annotate('', xy=(6.0, y_arrow), xytext=(6.0, y_arrow + 0.05),
                    arrowprops=arrow_props)

    # Bottom notice
    ax.text(6, 0.4, "Изолированный контур АССН: отсутствие внешних зависимостей, накладные расходы < 2-3% CPU",
            fontsize=8.5, style='italic', ha='center', color='#64748B')

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_monitoring_architecture.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

# 2. Wait Model & Time Decomposition Diagram
def make_wait_model_diagram():
    fig, ax = plt.subplots(figsize=(11, 6.5), dpi=300)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 7)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(5.5, 6.5, "Модель времени базы данных (DB Time) и классификация ожиданий",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    # Main formula box
    formula_rect = patches.FancyBboxPatch((1.0, 5.2), 9.0, 0.9, boxstyle="round,pad=0.1,rounding_size=0.15",
                                         facecolor='#0F172A', edgecolor='#38BDF8', linewidth=1.5)
    ax.add_patch(formula_rect)
    ax.text(5.5, 5.65, "DB Time = CPU Time + Non-Idle Wait Time",
            fontsize=13, fontweight='bold', ha='center', color='#38BDF8', fontfamily='sans-serif')

    # Categories of DB Time
    cards = [
        ("CPU Time (Вычисления)", 0.6, 2.5, 2.3, 2.2, "#DCFCE7", "#166534",
         ["Разбор SQL (Parse)", "Выполнение кода PL/SQL", "Обработка данных в RAM", "Сортировка в PGA"]),
        ("User I/O (Дисковая подсистема)", 3.1, 2.5, 2.3, 2.2, "#DBEAFE", "#1E40AF",
         ["db file sequential read\n(чтение по индексу)", "db file scattered read\n(сканирование таблицы)", "direct path read/write\n(временные файлы)"]),
        ("Concurrency (Блокировки)", 5.6, 2.5, 2.3, 2.2, "#FEE2E2", "#991B1B",
         ["enq: TX - row lock\n(конфликты строк)", "buffer busy waits\n(горячие блоки)", "latch / mutex free\n(конкуренция памяти)"]),
        ("Commit / Log (Фиксация)", 8.1, 2.5, 2.3, 2.2, "#FEF3C7", "#92400E",
         ["log file sync\n(ожидание LGWR)", "log file parallel write", "Задержки подсистемы\nжурналирования Redo"])
    ]

    for title, x, y, w, h, bg_col, text_col, points in cards:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor=bg_col, edgecolor='#94A3B8', linewidth=1.2)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h - 0.25, title, fontsize=9.0, fontweight='bold', ha='center', color=text_col, fontfamily='sans-serif')
        
        curr_y = y + h - 0.65
        for pt in points:
            ax.text(x + 0.12, curr_y, "• " + pt, fontsize=7.5, color='#1E293B', fontfamily='sans-serif', va='top')
            curr_y -= 0.42

    # Bottom insight box
    bot_rect = patches.FancyBboxPatch((0.6, 0.7), 9.8, 1.3, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor='#F1F5F9', edgecolor='#64748B', linewidth=1.0)
    ax.add_patch(bot_rect)
    ax.text(0.9, 1.6, "Диагностический вывод в реальном времени:", fontsize=9.5, fontweight='bold', color='#0F172A')
    ax.text(0.9, 1.25, "1. Если преобладает CPU Time — неоптимальные планы запросов, избыточная логика или нехватка ядер.", fontsize=8.0, color='#334155')
    ax.text(0.9, 0.95, "2. Если преобладает User I/O — отсутствуют необходимые индексы, устарела статистика CBO или деградация дисков.", fontsize=8.0, color='#334155')

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_oracle_wait_model.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

# 3. Fine-Grained Auditing (FGA) Architecture Diagram
def make_fga_diagram():
    fig, ax = plt.subplots(figsize=(11, 6.5), dpi=300)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 7)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(5.5, 6.5, "Архитектура контекстного аудита безопасности (Fine-Grained Auditing)",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    steps = [
        ("1. Инициация запроса", 0.6, 4.0, 2.2, 1.8, "#F1F5F9", "#0F172A",
         ["Пользователь / АРМ", "Выполнение SELECT", "Обращение к закрытым", "секторам данных"]),
        ("2. FGA Policy Engine", 3.2, 4.0, 2.3, 1.8, "#DBEAFE", "#1E40AF",
         ["Проверка политики:", "• Столбцы (SALARY...)", "• Условие (VALUE > N)", "• Контекст сессии"]),
        ("3. Ядро СУБД", 5.9, 4.0, 2.2, 1.8, "#DCFCE7", "#166534",
         ["Выполнение SQL", "Возврат выборки", "Асинхронная фиксация", "инцидента"]),
        ("4. Журнал и SIEM", 8.4, 4.0, 2.2, 1.8, "#FEE2E2", "#991B1B",
         ["FGA_LOG$ / Unified", "Запись: IP, OS_USER,", "Текст SQL, Значения", "Оповещение ИБ"]),
    ]

    for title, x, y, w, h, bg_col, text_col, points in steps:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor=bg_col, edgecolor='#94A3B8', linewidth=1.2)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h - 0.25, title, fontsize=9.0, fontweight='bold', ha='center', color=text_col, fontfamily='sans-serif')
        curr_y = y + h - 0.6
        for pt in points:
            ax.text(x + 0.15, curr_y, pt, fontsize=7.5, color='#334155', fontfamily='sans-serif', va='top')
            curr_y -= 0.3

    # Arrows between steps
    arrow_props = dict(facecolor='#0284C7', edgecolor='#0284C7', width=1.5, headwidth=6, headlength=6)
    for x_arrow in [2.9, 5.6, 8.2]:
        ax.annotate('', xy=(x_arrow + 0.2, 4.9), xytext=(x_arrow - 0.05, 4.9), arrowprops=arrow_props)

    # Comparison box: Standard Audit vs FGA
    comp_rect = patches.FancyBboxPatch((0.6, 1.0), 9.8, 2.4, boxstyle="round,pad=0.1,rounding_size=0.15",
                                       facecolor='#FFFFFF', edgecolor='#CBD5E1', linewidth=1.2)
    ax.add_patch(comp_rect)
    ax.text(0.9, 3.0, "Сравнение традиционного аудита и FGA в АС специального назначения", fontsize=9.5, fontweight='bold', color='#0F172A')

    ax.text(0.9, 2.5, "Традиционный аудит (Standard Audit):", fontsize=8.5, fontweight='bold', color='#B91C1C')
    ax.text(0.9, 2.2, "• Фиксирует только факт доступа (КТО и КОГДА открыл таблицу), без привязки к условиям.", fontsize=8.0, color='#334155')
    ax.text(0.9, 1.9, "• Огромный объем избыточных логов, невозможно выделить компрометацию конкретной записи.", fontsize=8.0, color='#334155')

    ax.text(0.9, 1.5, "Тонко гранулированный аудит (FGA):", fontsize=8.5, fontweight='bold', color='#15803D')
    ax.text(0.9, 1.2, "• Срабатывает только при чтении критических атрибутов при выполнении условия (WHERE dep = 10).", fontsize=8.0, color='#334155')
    ax.text(0.9, 0.9, "• Сохраняет точный текст SQL, значения привязанных переменных (bind variables) и контекст среды.", fontsize=8.0, color='#334155')

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_audit_fga_workflow.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

# 4. Performance Tuning Cycle Diagram
def make_tuning_cycle_diagram():
    fig, ax = plt.subplots(figsize=(11, 6.5), dpi=300)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 7)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(5.5, 6.5, "Замкнутый цикл диагностики и устранения деградации производительности",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    nodes = [
        ("1. Непрерывный сбор метрик", 0.6, 3.8, 2.2, 1.8, "#E0F2FE", "#0369A1", "TPS, задержки, CPU, I/O,\nочереди блокировок"),
        ("2. Детекция аномалии", 2.9, 3.8, 2.2, 1.8, "#FEE2E2", "#B91C1C", "Срабатывание триггеров\nPrometheus / Alertmanager"),
        ("3. Экспресс-анализ ASH", 5.2, 3.8, 2.2, 1.8, "#FEF3C7", "#B45309", "Изоляция сессии,\nSQL_ID и событий ожидания"),
        ("4. Профилирование CBO", 7.5, 3.8, 2.2, 1.8, "#F3E8FF", "#7E22CE", "Анализ плана выполнения\n(EXPLAIN PLAN, Rows, Cost)"),
    ]

    for title, x, y, w, h, bg_col, text_col, desc in nodes:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor=bg_col, edgecolor='#94A3B8', linewidth=1.2)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h - 0.25, title, fontsize=8.5, fontweight='bold', ha='center', color=text_col, fontfamily='sans-serif')
        ax.text(x + w/2, y + 0.65, desc, fontsize=7.5, ha='center', va='center', color='#334155', fontfamily='sans-serif')

    # Arrow 1->2, 2->3, 3->4
    arrow_props = dict(facecolor='#475569', edgecolor='#475569', width=1.5, headwidth=6, headlength=6)
    ax.annotate('', xy=(2.9, 4.7), xytext=(2.7, 4.7), arrowprops=arrow_props)
    ax.annotate('', xy=(5.2, 4.7), xytext=(5.0, 4.7), arrowprops=arrow_props)
    ax.annotate('', xy=(7.5, 4.7), xytext=(7.3, 4.7), arrowprops=arrow_props)

    # Resolution step below
    res_rect = patches.FancyBboxPatch((2.0, 1.2), 6.5, 1.8, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor='#DCFCE7', edgecolor='#16A34A', linewidth=1.5)
    ax.add_patch(res_rect)
    ax.text(5.25, 2.65, "5. Корректирующее действие и фиксация стабильности", fontsize=9.5, fontweight='bold', ha='center', color='#166534')
    ax.text(5.25, 2.2, "• Пересчет статистики DBMS_STATS / Создание покрывающего индекса", fontsize=8.0, ha='center', color='#14532D')
    ax.text(5.25, 1.85, "• Фиксация стабильного плана (SQL Plan Baselines / Stored Outlines)", fontsize=8.0, ha='center', color='#14532D')
    ax.text(5.25, 1.5, "• Верификация возврата в рамки SLA без перезапуска инстанса СУБД", fontsize=8.0, ha='center', color='#14532D')

    # Connectors to Resolution
    ax.annotate('', xy=(7.5, 2.2), xytext=(8.6, 3.8), arrowprops=arrow_props)
    ax.annotate('', xy=(1.7, 3.8), xytext=(2.5, 2.2), arrowprops=arrow_props)

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_performance_tuning_cycle.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

if __name__ == '__main__':
    make_architecture_diagram()
    make_wait_model_diagram()
    make_fga_diagram()
    make_tuning_cycle_diagram()
    print("All diagrams generated successfully!")

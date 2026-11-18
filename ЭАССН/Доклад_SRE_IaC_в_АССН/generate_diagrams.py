# -*- coding: utf-8 -*-
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches

fig_dir = r"c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\Доклад_SRE_IaC_в_АССН\figures"
os.makedirs(fig_dir, exist_ok=True)

# 1. SRE Model: SLI, SLO and Error Budget
def make_sre_budget_diagram():
    fig, ax = plt.subplots(figsize=(11, 6.8), dpi=300)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 7.5)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(5.5, 7.1, "Концептуальная модель SRE: SLI, SLO и управление Бюджетом Ошибок",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    # Top summary formula box
    form_rect = patches.FancyBboxPatch((0.8, 5.7), 9.4, 0.95, boxstyle="round,pad=0.1,rounding_size=0.15",
                                       facecolor='#0F172A', edgecolor='#38BDF8', linewidth=1.5)
    ax.add_patch(form_rect)
    ax.text(5.5, 6.25, "Error Budget = 100% – SLO   |   При SLO = 99.95% допуск простоя = 21.6 мин / месяц",
            fontsize=11.5, fontweight='bold', ha='center', color='#38BDF8', fontfamily='sans-serif')
    ax.text(5.5, 5.85, "Бюджет ошибок — это математический инструмент балансировки надежности и скорости изменений",
            fontsize=9.0, ha='center', color='#94A3B8', fontfamily='sans-serif')

    # 3 Pillars
    pillars = [
        ("SLI (Service Level Indicator)", 0.6, 2.7, 3.0, 2.6, "#DBEAFE", "#1E40AF",
         ["Фактическая метрика работы:", "• Доля успешных транзакций", "• Задержка отклика (P99 < 50мс)", "• Доступность узлов СУБД", "• Точность данных телеметрии"]),
        ("SLO (Service Level Objective)", 3.9, 2.7, 3.2, 2.6, "#DCFCE7", "#166534",
         ["Целевой уровень надежности:", "• Формальное соглашение SLA", "• Цель: 99.95% доступности", "• Измеримый временной интервал", "• Граница допустимого сбоя"]),
        ("Error Budget (Бюджет ошибок)", 7.4, 2.7, 3.0, 2.6, "#FEF3C7", "#92400E",
         ["Ресурс для модернизации:", "• Расходуется при авариях", "• Расходуется при релизах", "• Если исчерпан (100%):", "  БЛОКИРОВКА обновлений,", "  переход на стабилизацию"]),
    ]

    for title, x, y, w, h, bg_col, text_col, points in pillars:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor=bg_col, edgecolor='#94A3B8', linewidth=1.2)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h - 0.28, title, fontsize=9.5, fontweight='bold', ha='center', color=text_col, fontfamily='sans-serif')
        curr_y = y + h - 0.65
        for pt in points:
            bold_prefix = pt.startswith("•") or pt.startswith("  БЛОКИРОВКА")
            ax.text(x + 0.15, curr_y, pt, fontsize=8.0, color='#1E293B' if not "БЛОКИРОВКА" in pt else '#B91C1C',
                    fontweight='bold' if "БЛОКИРОВКА" in pt else 'normal', fontfamily='sans-serif', va='top')
            curr_y -= 0.35

    # Bottom policy container
    bot_rect = patches.FancyBboxPatch((0.6, 0.7), 9.8, 1.6, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor='#FFFFFF', edgecolor='#CBD5E1', linewidth=1.2)
    ax.add_patch(bot_rect)
    ax.text(0.9, 1.95, "Политика расходования бюджета ошибок в АССН (Error Budget Policy):", fontsize=9.5, fontweight='bold', color='#0F172A')
    ax.text(0.9, 1.6, "1. Бюджет > 50% — нормальный режим: плановая модернизация, внедрение новых алгоритмов управления.", fontsize=8.0, color='#15803D')
    ax.text(0.9, 1.25, "2. Бюджет < 20% — предупреждение: ужесточение тестирования, развертывание только критических исправлений.", fontsize=8.0, color='#D97706')
    ax.text(0.9, 0.9, "3. Бюджет = 0% — заморозка (Freeze): запрет любых функциональных изменений; все силы SRE — на устранение причин сбоев.", fontsize=8.0, color='#DC2626', fontweight='bold')

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_sre_sli_slo_error_budget.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

# 2. Air-Gapped IaC Pipeline Diagram
def make_iac_pipeline_diagram():
    fig, ax = plt.subplots(figsize=(12, 7.2), dpi=300)
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 8.0)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(6, 7.6, "Пайплайн Infrastructure as Code (IaC) в изолированном контуре АССН (Air-Gapped)",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    stages = [
        ("1. Версионирование (Git)", 0.6, 4.3, 2.4, 2.6, "#E0F2FE", "#0369A1",
         ["Декларативные манифесты:", "• Terraform / OpenTofu (ВМ, сеть)", "• Ansible Playbooks (ОС, СУБД)", "• Dockerfiles / Compose", "• Подписание коммитов ЭЦП"]),
        ("2. Контроль качества (Gate)", 3.4, 4.3, 2.5, 2.6, "#FEE2E2", "#B91C1C",
         ["DevSecOps проверки:", "• tflint, ansible-lint", "• Сканирование уязвимостей (Trivy)", "• Контроль требований ФСТЭК", "• Валидация идемпотентности"]),
        ("3. Локальные реестры", 6.3, 4.3, 2.4, 2.6, "#FEF3C7", "#92400E",
         ["Изолированный контур:", "• Доверенный Nexus / Harbor", "• Локальные mirror-пакеты", "• Golden Images (Packer)", "• Нулевая внешняя зависимость"]),
        ("4. Целевой флот АССН", 9.1, 4.3, 2.4, 2.6, "#DCFCE7", "#166534",
         ["Автоматизированная среда:", "• Astra Linux Special Edition", "• Кластеры СУБД (Postgres Pro)", "• Узлы телеметрии и связи", "• Серверы прикладного ПО"]),
    ]

    for title, x, y, w, h, bg_col, text_col, points in stages:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor=bg_col, edgecolor='#94A3B8', linewidth=1.2)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h - 0.28, title, fontsize=9.0, fontweight='bold', ha='center', color=text_col, fontfamily='sans-serif')
        curr_y = y + h - 0.65
        for pt in points:
            ax.text(x + 0.12, curr_y, pt, fontsize=7.5, color='#334155', fontfamily='sans-serif', va='top')
            curr_y -= 0.38

    # Arrows
    arrow_props = dict(facecolor='#0284C7', edgecolor='#0284C7', width=1.5, headwidth=6, headlength=6)
    for x_arr in [3.0, 5.9, 8.7]:
        ax.annotate('', xy=(x_arr + 0.35, 5.6), xytext=(x_arr, 5.6), arrowprops=arrow_props)

    # Key Principles Container
    prin_rect = patches.FancyBboxPatch((0.6, 0.8), 10.9, 3.0, boxstyle="round,pad=0.1,rounding_size=0.15",
                                       facecolor='#FFFFFF', edgecolor='#CBD5E1', linewidth=1.2)
    ax.add_patch(prin_rect)
    ax.text(0.9, 3.4, "Ключевые принципы надежности IaC при эксплуатации спецсистем:", fontsize=10, fontweight='bold', color='#0F172A')

    ax.text(0.9, 2.9, "• Идемпотентность (Idempotency): многократное повторение конфигурации приводит к строго одинаковому целевому состоянию.", fontsize=8.5, color='#334155')
    ax.text(0.9, 2.4, "• Устранение дрейфа конфигураций (Configuration Drift): полный запрет ручных правок на серверах через SSH (ручные правки затираются кодом).", fontsize=8.5, color='#334155')
    ax.text(0.9, 1.9, "• Иммутабельная инфраструктура (Immutable Infrastructure): при обновлении узла разворачивается новый эталонный сервер взамен старого.", fontsize=8.5, color='#334155')
    ax.text(0.9, 1.4, "• Самодокументируемость: репозиторий Git является единственным достоверным источником правды (Single Source of Truth) об архитектуре.", fontsize=8.5, color='#334155')

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_iac_airgap_pipeline.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

# 3. Zero-Downtime Deployment Diagram
def make_zero_downtime_diagram():
    fig, ax = plt.subplots(figsize=(11, 6.8), dpi=300)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 7.5)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(5.5, 7.1, "Схема безопасного развертывания без простоя (Zero-Downtime Blue-Green / Canary)",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    # Traffic Router / Balancer
    router_rect = patches.FancyBboxPatch((3.5, 5.6), 4.0, 0.9, boxstyle="round,pad=0.1,rounding_size=0.15",
                                        facecolor='#0F172A', edgecolor='#38BDF8', linewidth=1.5)
    ax.add_patch(router_rect)
    ax.text(5.5, 6.15, "Маршрутизатор трафика / Балансировщик", fontsize=9.5, fontweight='bold', ha='center', color='#38BDF8')
    ax.text(5.5, 5.8, "Плавное переключение трафика (90% -> 10% Canary -> 100%)", fontsize=8.0, ha='center', color='#94A3B8')

    # Blue and Green environments
    # Blue (Active)
    blue_rect = patches.FancyBboxPatch((0.8, 2.5), 4.3, 2.5, boxstyle="round,pad=0.1,rounding_size=0.15",
                                       facecolor='#DBEAFE', edgecolor='#1D4ED8', linewidth=1.5)
    ax.add_patch(blue_rect)
    ax.text(2.95, 4.65, "Контур BLUE (Текущая версия v1.2)", fontsize=9.5, fontweight='bold', ha='center', color='#1E40AF')
    ax.text(2.95, 4.15, "• Обрабатывает 100% боевого трафика\n• Полная доступность сервисов 24/7\n• Мгновенный откат при сбое Canary",
            fontsize=8.0, ha='center', color='#1E293B')

    # Green (Target)
    green_rect = patches.FancyBboxPatch((5.9, 2.5), 4.3, 2.5, boxstyle="round,pad=0.1,rounding_size=0.15",
                                        facecolor='#DCFCE7', edgecolor='#15803D', linewidth=1.5)
    ax.add_patch(green_rect)
    ax.text(8.05, 4.65, "Контур GREEN (Новая версия v1.3)", fontsize=9.5, fontweight='bold', ha='center', color='#166534')
    ax.text(8.05, 4.15, "• Развернут автоматически через IaC\n• Проходит синтетические smoke-тесты\n• Прием 5-10% нагрузки для валидации",
            fontsize=8.0, ha='center', color='#1E293B')

    # Arrows from Router
    arrow_blue = dict(facecolor='#1D4ED8', edgecolor='#1D4ED8', width=1.5, headwidth=6, headlength=6)
    arrow_green = dict(facecolor='#15803D', edgecolor='#15803D', width=1.5, headwidth=6, headlength=6)
    ax.annotate('', xy=(2.95, 5.0), xytext=(4.5, 5.6), arrowprops=arrow_blue)
    ax.annotate('', xy=(8.05, 5.0), xytext=(6.5, 5.6), arrowprops=arrow_green)

    # Database Migration Layer (Shared state)
    db_rect = patches.FancyBboxPatch((2.2, 0.6), 6.6, 1.4, boxstyle="round,pad=0.1,rounding_size=0.15",
                                     facecolor='#FEF3C7', edgecolor='#D97706', linewidth=1.5)
    ax.add_patch(db_rect)
    ax.text(5.5, 1.6, "СУБД: Паттерн Expand-Contract (Параллельная совместимость схемы)",
            fontsize=9.5, fontweight='bold', ha='center', color='#92400E')
    ax.text(5.5, 1.25, "1. Expand: добавление новых колонок/триггеров с обратной совместимостью для v1.2.", fontsize=8.0, ha='center', color='#334155')
    ax.text(5.5, 0.95, "2. Contract: удаление устаревших структур только ПОСЛЕ полного перевода трафика на v1.3.", fontsize=8.0, ha='center', color='#334155')

    # Connectors to DB
    arrow_db = dict(facecolor='#B45309', edgecolor='#B45309', width=1.2, headwidth=5, headlength=5)
    ax.annotate('', xy=(3.8, 2.0), xytext=(2.95, 2.5), arrowprops=arrow_db)
    ax.annotate('', xy=(7.2, 2.0), xytext=(8.05, 2.5), arrowprops=arrow_db)

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_zero_downtime_deployment.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

# 4. SRE Incident Lifecycle & Toil Reduction
def make_incident_lifecycle_diagram():
    fig, ax = plt.subplots(figsize=(11, 6.8), dpi=300)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 7.5)
    ax.axis('off')

    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    ax.text(5.5, 7.1, "Замкнутый жизненный цикл инцидента и борьба с рутиной (Toil Reduction)",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    steps = [
        ("1. Автоматическая детекция", 0.6, 4.3, 2.2, 2.2, "#DBEAFE", "#1E40AF",
         ["Срабатывание SLI/SLO", "Алерт в Alertmanager", "Отсутствие ручных", "проверок дежурным"]),
        ("2. Авто-смягчение (Mitigate)", 2.9, 4.3, 2.2, 2.2, "#FEE2E2", "#B91C1C",
         ["Авто-Failover кластера", "Изоляция сбойного узла", "Откат канареечного", "релиза назад"]),
        ("3. Локализация (RCA)", 5.2, 4.3, 2.2, 2.2, "#FEF3C7", "#92400E",
         ["Анализ метрик и логов", "Изучение трейсов", "Выявление первопричины", "(Root Cause)"]),
        ("4. Blameless Post-Mortem", 7.5, 4.3, 2.4, 2.2, "#F3E8FF", "#7E22CE",
         ["Разбор без чувства вины", "Фокус на системных", "недостатках, а не", "на ошибке человека"]),
    ]

    for title, x, y, w, h, bg_col, text_col, points in steps:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor=bg_col, edgecolor='#94A3B8', linewidth=1.2)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h - 0.28, title, fontsize=8.5, fontweight='bold', ha='center', color=text_col, fontfamily='sans-serif')
        curr_y = y + h - 0.65
        for pt in points:
            ax.text(x + 0.12, curr_y, pt, fontsize=7.5, color='#334155', fontfamily='sans-serif', va='top')
            curr_y -= 0.35

    arrow_props = dict(facecolor='#475569', edgecolor='#475569', width=1.5, headwidth=6, headlength=6)
    ax.annotate('', xy=(2.9, 5.4), xytext=(2.7, 5.4), arrowprops=arrow_props)
    ax.annotate('', xy=(5.2, 5.4), xytext=(5.0, 5.4), arrowprops=arrow_props)
    ax.annotate('', xy=(7.5, 5.4), xytext=(7.3, 5.4), arrowprops=arrow_props)

    # 5th Toil Reduction block below
    res_rect = patches.FancyBboxPatch((1.5, 1.1), 7.8, 2.4, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor='#DCFCE7', edgecolor='#16A34A', linewidth=1.5)
    ax.add_patch(res_rect)
    ax.text(5.4, 3.1, "5. Автоматизация и сокращение рутины (Eliminating Toil via Code)",
            fontsize=10, fontweight='bold', ha='center', color='#166534')
    ax.text(5.4, 2.65, "• Правило SRE: рутинные механические действия (Toil) должны составлять не более 50% времени инженера.", fontsize=8.0, ha='center', color='#14532D')
    ax.text(5.4, 2.25, "• Каждая ручная операция инцидента оформляется в виде автоматизированного сценария Ansible / Terraform.", fontsize=8.0, ha='center', color='#14532D')
    ax.text(5.4, 1.85, "• Включение автотестов в CI/CD для гарантированного предотвращения рецидива сбоя.", fontsize=8.0, ha='center', color='#14532D')
    ax.text(5.4, 1.45, "• Постоянное повышение надежности АССН как инженерная дисциплина, а не «тушение пожаров».", fontsize=8.0, ha='center', color='#14532D')

    # Connectors
    ax.annotate('', xy=(8.7, 3.5), xytext=(8.7, 4.3), arrowprops=arrow_props)
    ax.annotate('', xy=(1.7, 4.3), xytext=(1.7, 3.5), arrowprops=arrow_props)

    plt.tight_layout()
    out_file = os.path.join(fig_dir, "diagram_incident_lifecycle_sre.png")
    plt.savefig(out_file, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print("Saved:", out_file)

if __name__ == '__main__':
    make_sre_budget_diagram()
    make_iac_pipeline_diagram()
    make_zero_downtime_diagram()
    make_incident_lifecycle_diagram()
    print("All SRE diagrams generated successfully!")

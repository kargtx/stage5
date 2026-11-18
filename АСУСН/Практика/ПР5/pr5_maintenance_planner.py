# -*- coding: utf-8 -*-
"""
pr5_maintenance_planner.py
Комплексный расчетно-аналитический модуль планирования технического обслуживания (ТО)
и регламентных работ для АС СН Нефтеперекачивающей станции (НПС).

Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С
УУНиТ, Кафедра АСУ, 2026 г.
"""

import os
import sys
import math
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.gridspec import GridSpec
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Настройка шрифтов для Matplotlib
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 10
plt.rcParams['axes.titlesize'] = 12
plt.rcParams['axes.labelsize'] = 10
plt.rcParams['xtick.labelsize'] = 9
plt.rcParams['ytick.labelsize'] = 9
plt.rcParams['legend.fontsize'] = 9
plt.rcParams['figure.titlesize'] = 14

# =============================================================================
# 1. СТРУКТУРА ДАННЫХ ОБОРУДОВАНИЯ НПС
# =============================================================================

EQUIPMENT_DATA = [
    {
        "id": "PT-101",
        "name": "Датчик давления на входе НПС",
        "model": "Метран-150-TG",
        "category": "КИПиА (Давление)",
        "qty": 3,
        "location": "Узел подключения / входной коллектор",
        "criticality": "Высокая (ПАЗ/РСУ)",
        "maint": {"ТО-1": (0.5, 24, "Слесарь КИПиА"), "ТО-2": (1.0, 12, "Слесарь КИПиА"), "ТО-4": (2.5, 1, "Инженер-метролог")},
        "mtbf_h": 120000,
        "zip_stock": 2
    },
    {
        "id": "PT-102",
        "name": "Датчик давления на выкиде МНА",
        "model": "Метран-150-TG",
        "category": "КИПиА (Давление)",
        "qty": 4,
        "location": "Насосный цех МНА-1..4",
        "criticality": "Критическая (ПАЗ)",
        "maint": {"ТО-1": (0.5, 24, "Слесарь КИПиА"), "ТО-2": (1.0, 12, "Слесарь КИПиА"), "ТО-4": (2.5, 1, "Инженер-метролог")},
        "mtbf_h": 120000,
        "zip_stock": 3
    },
    {
        "id": "TT-201",
        "name": "Датчики температуры подшипников МНА",
        "model": "Метран-280 / ТСП Метран-206",
        "category": "КИПиА (Температура)",
        "qty": 16,
        "location": "Подшипниковые узлы МНА и электродвигателей",
        "criticality": "Критическая (ПАЗ)",
        "maint": {"ТО-1": (0.3, 24, "Слесарь КИПиА"), "ТО-2": (0.8, 12, "Слесарь КИПиА"), "ТО-4": (1.5, 1, "Инженер-метролог")},
        "mtbf_h": 150000,
        "zip_stock": 4
    },
    {
        "id": "VT-301",
        "name": "Датчики виброскорости и виброперемещения",
        "model": "ВК-310В / ИВ-Д-ТР",
        "category": "КИПиА (Вибрация)",
        "qty": 8,
        "location": "Корпуса насосов МНА и опоры двигателей",
        "criticality": "Критическая (ПАЗ)",
        "maint": {"ТО-1": (0.4, 24, "Слесарь КИПиА"), "ТО-2": (1.0, 12, "Инженер АСУ ТП"), "ТО-4": (2.0, 1, "Инженер-метролог")},
        "mtbf_h": 90000,
        "zip_stock": 2
    },
    {
        "id": "FT-401",
        "name": "Ультразвуковой расходомер нефти",
        "model": "Взлет МР / Emerson Daniel SeniorSonic",
        "category": "КИПиА (Расход)",
        "qty": 2,
        "location": "Узел коммерческого/оперативного учета (СИКН)",
        "criticality": "Высокая (РСУ)",
        "maint": {"ТО-1": (0.5, 24, "Слесарь КИПиА"), "ТО-2": (1.5, 12, "Инженер АСУ ТП"), "ТО-4": (4.0, 1, "Инженер-метролог")},
        "mtbf_h": 100000,
        "zip_stock": 1
    },
    {
        "id": "LT-501",
        "name": "Радарный уровнемер резервуаров РВС",
        "model": "VEGAPULS 64 / Ризур-М",
        "category": "КИПиА (Уровень)",
        "qty": 4,
        "location": "Резервуарный парк (РВС-20000)",
        "criticality": "Высокая (ПАЗ/РСУ)",
        "maint": {"ТО-1": (0.5, 24, "Слесарь КИПиА"), "ТО-2": (1.2, 12, "Слесарь КИПиА"), "ТО-4": (3.0, 1, "Инженер-метролог")},
        "mtbf_h": 110000,
        "zip_stock": 1
    },
    {
        "id": "GT-601",
        "name": "Оптический датчик довзрывных концентраций газа",
        "model": "СГОЭС-М11 / Сигма-03",
        "category": "КИПиА (Загазованность)",
        "qty": 12,
        "location": "Насосный зал, резервуарный парк, узел фильтров",
        "criticality": "Критическая (ПАЗ)",
        "maint": {"ТО-1": (0.3, 24, "Слесарь КИПиА"), "ТО-2": (0.8, 12, "Слесарь КИПиА"), "ТО-3": (1.5, 4, "Слесарь КИПиА"), "ТО-4": (2.0, 1, "Инженер-метролог")},
        "mtbf_h": 80000,
        "zip_stock": 3
    },
    {
        "id": "PLC-01",
        "name": "Основной контроллер РСУ НПС",
        "model": "Regul R500 (дублированные CPU 01)",
        "category": "Контроллеры (ПЛК)",
        "qty": 1,
        "location": "Шкаф управления ШУ-01 (операторная)",
        "criticality": "Критическая (РСУ)",
        "maint": {"ТО-2": (1.0, 12, "Инженер АСУ ТП"), "ТО-3": (3.0, 4, "Инженер АСУ ТП"), "ТО-4": (6.0, 1, "Инженер АСУ ТП")},
        "mtbf_h": 200000,
        "zip_stock": 1
    },
    {
        "id": "PLC-02",
        "name": "Контроллер системы ПАЗ (SIL3)",
        "model": "Regul R500S (отказоустойчивый)",
        "category": "Контроллеры (ПЛК)",
        "qty": 1,
        "location": "Шкаф противоаварийной защиты ШПАЗ-01",
        "criticality": "Критическая (ПАЗ)",
        "maint": {"ТО-2": (1.0, 12, "Инженер АСУ ТП"), "ТО-3": (3.5, 4, "Инженер АСУ ТП"), "ТО-4": (8.0, 1, "Инженер АСУ ТП")},
        "mtbf_h": 250000,
        "zip_stock": 1
    },
    {
        "id": "IO-MOD",
        "name": "Модули ввода/вывода (AI, DI, DO, AO, TC)",
        "model": "Серия Regul R500 (AI-08-H, DI-16, DO-16)",
        "category": "Контроллеры (ПЛК)",
        "qty": 24,
        "location": "Шкафы контроллерные и кроссовые",
        "criticality": "Высокая",
        "maint": {"ТО-2": (0.3, 12, "Инженер АСУ ТП"), "ТО-3": (0.5, 4, "Инженер АСУ ТП"), "ТО-4": (1.0, 1, "Инженер АСУ ТП")},
        "mtbf_h": 180000,
        "zip_stock": 5
    },
    {
        "id": "ACT-101",
        "name": "Электроприводная задвижка магистральная",
        "model": "Auma SA 14.2 / блок управления AC 01.2",
        "category": "Исполнительные механизмы",
        "qty": 6,
        "location": "Вход/выход НПС, гребенка коллектора",
        "criticality": "Критическая (ПАЗ/РСУ)",
        "maint": {"ТО-1": (0.5, 24, "Слесарь КИПиА"), "ТО-2": (1.5, 12, "Электромонтер"), "ТО-3": (3.0, 4, "Электромонтер"), "ТО-4": (5.0, 1, "Электромонтер")},
        "mtbf_h": 70000,
        "zip_stock": 1
    },
    {
        "id": "ACT-201",
        "name": "Регулирующий клапан узла гашения ударных волн",
        "model": "Mokveld Axial Control Valve",
        "category": "Исполнительные механизмы",
        "qty": 2,
        "location": "Узел СРД (сброс волны давления)",
        "criticality": "Критическая (ПАЗ)",
        "maint": {"ТО-1": (0.5, 24, "Слесарь КИПиА"), "ТО-2": (2.0, 12, "Слесарь КИПиА"), "ТО-3": (4.0, 4, "Слесарь КИПиА"), "ТО-4": (6.0, 1, "Слесарь КИПиА")},
        "mtbf_h": 85000,
        "zip_stock": 1
    },
    {
        "id": "CAB-01",
        "name": "Шкафы автоматики, кроссовые и интерфейсные",
        "model": "Rittal TS8 (со степенью защиты IP54)",
        "category": "Шкафы автоматики",
        "qty": 4,
        "location": "Серверная / операторный зал НПС",
        "criticality": "Высокая",
        "maint": {"ТО-1": (0.2, 24, "Электромонтер"), "ТО-2": (0.5, 12, "Электромонтер"), "ТО-3": (2.0, 4, "Электромонтер"), "ТО-4": (4.0, 1, "Электромонтер")},
        "mtbf_h": 300000,
        "zip_stock": 0
    },
    {
        "id": "UPS-01",
        "name": "Источник бесперебойного питания (ИБП)",
        "model": "Штиль / APC Smart-UPS RT 10kVA On-Line",
        "category": "Источники питания",
        "qty": 2,
        "location": "Шкаф гарантированного питания ШГП",
        "criticality": "Критическая",
        "maint": {"ТО-1": (0.3, 24, "Электромонтер"), "ТО-2": (1.0, 12, "Электромонтер"), "ТО-3": (2.5, 4, "Электромонтер"), "ТО-4": (4.0, 1, "Электромонтер")},
        "mtbf_h": 100000,
        "zip_stock": 1
    },
    {
        "id": "BAT-01",
        "name": "Аккумуляторный блок ИБП (16x12В AGM)",
        "model": "Delta HRL 12-55 (55 А·ч, герметичные)",
        "category": "Источники питания",
        "qty": 2,
        "location": "Батарейные полки ШГП",
        "criticality": "Критическая",
        "maint": {"ТО-2": (1.0, 12, "Электромонтер"), "ТО-3": (2.0, 4, "Электромонтер"), "ТО-4": (4.0, 1, "Электромонтер")},
        "mtbf_h": 60000,
        "zip_stock": 4
    },
    {
        "id": "NET-01",
        "name": "Промышленные коммутаторы Ethernet Turbo Ring",
        "model": "MOXA EDS-510E-3SFP",
        "category": "Каналы связи",
        "qty": 4,
        "location": "Шкафы ШУ-01, ШПАЗ-01, ШС-01",
        "criticality": "Критическая",
        "maint": {"ТО-2": (0.5, 12, "Инженер АСУ ТП"), "ТО-3": (1.5, 4, "Инженер АСУ ТП"), "ТО-4": (3.0, 1, "Инженер АСУ ТП")},
        "mtbf_h": 350000,
        "zip_stock": 1
    },
    {
        "id": "SCADA-01",
        "name": "Серверы диспетчерского контроля и АРМ",
        "model": "Серверы HP ProLiant / ПО MasterSCADA 4D",
        "category": "SCADA-система",
        "qty": 4,
        "location": "Серверные стойки, рабочие места операторов",
        "criticality": "Критическая",
        "maint": {"ТО-1": (0.5, 24, "Инженер АСУ ТП"), "ТО-2": (1.5, 12, "Инженер АСУ ТП"), "ТО-3": (3.0, 4, "Инженер АСУ ТП"), "ТО-4": (8.0, 1, "Инженер АСУ ТП")},
        "mtbf_h": 80000,
        "zip_stock": 1
    }
]

# =============================================================================
# 2. РАСЧЕТ БАЛАНСА ТРУДОЕМКОСТИ И РАСПРЕДЕЛЕНИЯ ПО МЕСЯЦАМ
# =============================================================================

def calculate_labor_statistics():
    """Рассчитывает общую трудоемкость по видам ТО, должностям и месяцам."""
    roles = ["Слесарь КИПиА", "Инженер АСУ ТП", "Инженер-метролог", "Электромонтер"]
    total_by_role = {r: 0.0 for r in roles}
    total_by_maint = {"ТО-1": 0.0, "ТО-2": 0.0, "ТО-3": 0.0, "ТО-4": 0.0}
    total_by_cat = {}

    table_rows = []

    for eq in EQUIPMENT_DATA:
        cat = eq["category"]
        if cat not in total_by_cat:
            total_by_cat[cat] = 0.0

        for m_type, (hours, freq_per_year, role) in eq["maint"].items():
            annual_item_hours = hours * freq_per_year * eq["qty"]
            total_by_role[role] += annual_item_hours
            total_by_maint[m_type] += annual_item_hours
            total_by_cat[cat] += annual_item_hours

            table_rows.append({
                "id": eq["id"],
                "name": eq["name"],
                "model": eq["model"],
                "qty": eq["qty"],
                "m_type": m_type,
                "labor_single": hours,
                "freq_per_year": freq_per_year,
                "freq_label": "2 раза в месяц" if freq_per_year == 24 else ("1 раз в месяц" if freq_per_year == 12 else ("1 раз в квартал" if freq_per_year == 4 else "1 раз в год")),
                "labor_annual": annual_item_hours,
                "role": role
            })

    # Распределение по 12 месяцам года с учетом плановых остановов НПС в Мае (мес 5) и Октябре (мес 10)
    # В эти месяцы проводятся крупные регламенты ТО-4 и поверки
    monthly_by_role = {r: np.zeros(12) for r in roles}

    for row in table_rows:
        r = row["role"]
        m_type = row["m_type"]
        qty = row["qty"]
        h = row["labor_single"]

        if m_type == "ТО-1":
            # 24 раза в год -> по 2 раза в месяц равномерно
            monthly_hours = h * 2 * qty
            monthly_by_role[r] += monthly_hours
        elif m_type == "ТО-2":
            # 12 раз в год -> каждый месяц 1 раз
            monthly_hours = h * 1 * qty
            monthly_by_role[r] += monthly_hours
        elif m_type == "ТО-3":
            # 4 раза в год -> март, июнь, сентябрь, декабрь
            quarters = [2, 5, 8, 11]  # 0-indexed: Мар, Июн, Сен, Дек
            for q_m in quarters:
                monthly_by_role[r][q_m] += h * qty
        elif m_type == "ТО-4":
            # 1 раз в год -> распределяем на плановые остановы (Май - индекс 4, Октябрь - индекс 9)
            if "метролог" in r.lower():
                # Поверки датчиков делим пополам: май и октябрь
                monthly_by_role[r][4] += (h * qty) * 0.55
                monthly_by_role[r][9] += (h * qty) * 0.45
            elif "Инженер АСУ" in r:
                # ПЛК и SCADA: часть в мае, часть в октябре
                monthly_by_role[r][4] += (h * qty) * 0.5
                monthly_by_role[r][9] += (h * qty) * 0.5
            else:
                # Силовое оборудование и задвижки
                monthly_by_role[r][4] += (h * qty) * 0.5
                monthly_by_role[r][9] += (h * qty) * 0.5

    return {
        "table_rows": table_rows,
        "total_by_role": total_by_role,
        "total_by_maint": total_by_maint,
        "total_by_cat": total_by_cat,
        "monthly_by_role": monthly_by_role,
        "grand_total": sum(total_by_role.values())
    }

# =============================================================================
# 3. ПОСТРОЕНИЕ 4 ГРАФИЧЕСКИХ ДИАГРАММ ВЫСОКОГО РАЗРЕШЕНИЯ (DPI 300)
# =============================================================================

def generate_figure_1_gantt(stats):
    """Диаграмма Ганта: Годовой график регламентного ТО АС СН НПС."""
    fig, ax = plt.subplots(figsize=(14, 8), dpi=300)

    subsystems = [
        ("Датчики давления PT-101/102", ["ТО-1", "ТО-2", "ТО-4"]),
        ("Датчики темп. подшипников TT-201", ["ТО-1", "ТО-2", "ТО-4"]),
        ("Датчики виброконтроля VT-301", ["ТО-1", "ТО-2", "ТО-4"]),
        ("УЗ расходомер нефти FT-401", ["ТО-1", "ТО-2", "ТО-4"]),
        ("Уровнемеры РВС LT-501", ["ТО-1", "ТО-2", "ТО-4"]),
        ("Газоанализаторы загазованности GT-601", ["ТО-1", "ТО-2", "ТО-3", "ТО-4"]),
        ("Контроллеры РСУ Regul R500 (PLC-01)", ["ТО-2", "ТО-3", "ТО-4"]),
        ("Контроллеры ПАЗ Regul R500S (PLC-02)", ["ТО-2", "ТО-3", "ТО-4"]),
        ("Модули ввода/вывода (IO-MOD)", ["ТО-2", "ТО-3", "ТО-4"]),
        ("Электроприводы задвижек Auma (ACT-101)", ["ТО-1", "ТО-2", "ТО-3", "ТО-4"]),
        ("Регулирующий клапан Mokveld (ACT-201)", ["ТО-1", "ТО-2", "ТО-3", "ТО-4"]),
        ("Шкафы ШУ/ШПАЗ/КШ (CAB-01)", ["ТО-1", "ТО-2", "ТО-3", "ТО-4"]),
        ("ИБП и батарейный блок (UPS/BAT)", ["ТО-1", "ТО-2", "ТО-3", "ТО-4"]),
        ("Сетевые коммутаторы MOXA (NET-01)", ["ТО-2", "ТО-3", "ТО-4"]),
        ("Серверы SCADA и АРМ (SCADA-01)", ["ТО-1", "ТО-2", "ТО-3", "ТО-4"]),
    ]

    months = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]
    n_sub = len(subsystems)
    y_pos = np.arange(n_sub)

    colors = {
        "ТО-1": "#10b981",  # Изумрудный
        "ТО-2": "#3b82f6",  # Синий
        "ТО-3": "#f59e0b",  # Янтарный
        "ТО-4": "#ef4444",  # Красный (поверка/капремонт)
    }

    # Подсветка окон останова НПС в Мае и Октябре
    ax.axvspan(4.0, 5.0, color="#fef3c7", alpha=0.6, label="Плановый останов НПС №1 (весна)")
    ax.axvspan(9.0, 10.0, color="#ffedd5", alpha=0.6, label="Плановый останов НПС №2 (осень)")

    # Сетка месяцев
    for m in range(13):
        ax.axvline(m, color="#cbd5e1", linestyle="--", linewidth=0.8, alpha=0.7)

    bar_height = 0.55

    for idx, (name, m_list) in enumerate(subsystems):
        y = n_sub - 1 - idx

        # Базовый фон полосы
        ax.barh(y, 12, left=0, height=bar_height, color="#f8fafc", edgecolor="#e2e8f0", linewidth=1.0)

        # Разметка ТО-1 и ТО-2 (регулярные фоновые)
        if "ТО-1" in m_list:
            for m in range(12):
                ax.barh(y, 0.4, left=m + 0.05, height=bar_height*0.4, color=colors["ТО-1"], alpha=0.9)
                ax.barh(y, 0.4, left=m + 0.55, height=bar_height*0.4, color=colors["ТО-1"], alpha=0.9)

        if "ТО-2" in m_list:
            for m in range(12):
                ax.barh(y, 0.8, left=m + 0.1, height=bar_height*0.35, color=colors["ТО-2"], alpha=0.85)

        # ТО-3 (ежеквартально: Мар, Июн, Сен, Дек)
        if "ТО-3" in m_list:
            for q_m in [2, 5, 8, 11]:
                ax.barh(y, 0.85, left=q_m + 0.08, height=bar_height*0.7, color=colors["ТО-3"], alpha=0.95, edgecolor="#b45309", linewidth=1)

        # ТО-4 / Поверка (Май или Октябрь)
        if "ТО-4" in m_list:
            # Чередуем май и октябрь для баланса
            target_m = 4 if idx % 2 == 0 else 9
            ax.barh(y, 0.9, left=target_m + 0.05, height=bar_height*0.95, color=colors["ТО-4"], alpha=0.95, edgecolor="#991b1b", linewidth=1.5)

    ax.set_yticks(y_pos)
    ax.set_yticklabels([s[0] for s in reversed(subsystems)], fontsize=9, fontweight='medium')
    ax.set_xticks(np.arange(12) + 0.5)
    ax.set_xticklabels(months, fontsize=10, fontweight='bold')
    ax.set_xlim(0, 12)
    ax.set_ylim(-0.6, n_sub - 0.4)
    ax.grid(axis='y', linestyle=':', alpha=0.5)

    # Кастомная легенда
    leg_patches = [
        patches.Patch(color=colors["ТО-1"], label="ТО-1 (Еженедельное / 2 раза в месяц)"),
        patches.Patch(color=colors["ТО-2"], label="ТО-2 (Ежемесячное профилактическое)"),
        patches.Patch(color=colors["ТО-3"], label="ТО-3 (Ежеквартальное расширенное)"),
        patches.Patch(color=colors["ТО-4"], label="ТО-4 / Поверка (Годовое регламентное)"),
        patches.Patch(facecolor="#fef3c7", edgecolor="#f59e0b", label="Окно останова НПС (весеннее/осеннее)")
    ]
    ax.legend(handles=leg_patches, loc="upper center", bbox_to_anchor=(0.5, 1.12), ncol=3, frameon=True, facecolor="#ffffff", edgecolor="#cbd5e1")

    ax.set_title("Годовой календарный план-график технического обслуживания АС СН НПС", fontsize=13, fontweight='bold', pad=35)
    plt.tight_layout()
    out_path = os.path.join(SCRIPT_DIR, "pr5_gantt_schedule.png")
    plt.savefig(out_path, dpi=300)
    plt.close()
    print(f"  + Сохранен: {out_path}")
    return out_path

def generate_figure_2_workload(stats):
    """Гистограмма и круговая диаграмма распределения трудоемкости ТО."""
    fig = plt.figure(figsize=(15, 7), dpi=300)
    gs = GridSpec(1, 2, width_ratios=[1.3, 1.0], wspace=0.25)

    # --- Подграфик 1: Ежемесячная загрузка по должностям (Stacked Bar) ---
    ax1 = fig.add_subplot(gs[0])
    months = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]
    x = np.arange(12)

    roles = ["Слесарь КИПиА", "Инженер АСУ ТП", "Инженер-метролог", "Электромонтер"]
    role_colors = ["#0ea5e9", "#6366f1", "#ec4899", "#f59e0b"]

    bottom = np.zeros(12)
    for r, col in zip(roles, role_colors):
        y_vals = stats["monthly_by_role"][r]
        ax1.bar(x, y_vals, bottom=bottom, label=r, color=col, edgecolor="#ffffff", linewidth=0.7, width=0.65)
        bottom += y_vals

    # Линия номинального фонда рабочего времени ремонтной службы (2 человека по 160 ч = 320 ч)
    # На ТО выделяется около 30-40% фонда смены (110 чел-ч в месяц)
    ax1.axhline(110, color="#dc2626", linestyle="--", linewidth=1.5, label="Нормативный лимит фонда ТО (110 чел·ч)")
    ax1.scatter([4, 9], [bottom[4], bottom[9]], color="#b91c1c", s=70, zorder=5)
    ax1.annotate("Останов НПС №1\n(май)", xy=(4, bottom[4]), xytext=(4, bottom[4]+12),
                 ha='center', fontsize=8, fontweight='bold', arrowprops=dict(arrowstyle="->", color="#b91c1c"))
    ax1.annotate("Останов НПС №2\n(октябрь)", xy=(9, bottom[9]), xytext=(9, bottom[9]+12),
                 ha='center', fontsize=8, fontweight='bold', arrowprops=dict(arrowstyle="->", color="#b91c1c"))

    ax1.set_xticks(x)
    ax1.set_xticklabels(months, fontweight='bold')
    ax1.set_ylabel("Трудоемкость ТО, чел·ч", fontweight='bold')
    ax1.set_title("Ежемесячное распределение трудоемкости ТО по специальностям", fontweight='bold')
    ax1.legend(loc="upper left", frameon=True, facecolor="#ffffff", edgecolor="#cbd5e1", fontsize=8.5)
    ax1.grid(axis='y', linestyle=':', alpha=0.6)
    ax1.set_ylim(0, max(bottom)*1.18)

    # --- Подграфик 2: Структура годовой трудоемкости (Donut Chart) ---
    ax2 = fig.add_subplot(gs[1])
    m_labels = ["ТО-1 (Еженедельное)", "ТО-2 (Ежемесячное)", "ТО-3 (Ежеквартальное)", "ТО-4 / Поверка"]
    m_values = [stats["total_by_maint"]["ТО-1"], stats["total_by_maint"]["ТО-2"],
                stats["total_by_maint"]["ТО-3"], stats["total_by_maint"]["ТО-4"]]
    m_colors = ["#10b981", "#3b82f6", "#f59e0b", "#ef4444"]

    wedges, texts, autotexts = ax2.pie(
        m_values, labels=m_labels, autopct='%1.1f%%', startangle=140,
        colors=m_colors, pctdistance=0.75, wedgeprops=dict(width=0.45, edgecolor='#ffffff', linewidth=2)
    )
    for at in autotexts:
        at.set_fontsize(9)
        at.set_fontweight('bold')
        at.set_color('#ffffff')

    ax2.text(0, 0, f"Всего:\n{stats['grand_total']:.0f} чел·ч", ha='center', va='center',
             fontsize=11, fontweight='bold', color="#1e293b")
    ax2.set_title("Структура годовой трудоемкости по видам ТО", fontweight='bold')

    plt.tight_layout()
    out_path = os.path.join(SCRIPT_DIR, "pr5_workload_distribution.png")
    plt.savefig(out_path, dpi=300)
    plt.close()
    print(f"  + Сохранен: {out_path}")
    return out_path

def generate_figure_3_incident_flowchart():
    """Блок-схема регламента действий персонала в нештатных ситуациях на НПС."""
    fig, ax = plt.subplots(figsize=(13, 9), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    def draw_box(x, y, w, h, text, bg_color="#e0f2fe", border_color="#0284c7", shape="rect", text_color="#0f172a", font_size=8.5, bold=False):
        if shape == "rect":
            box = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.6,rounding_size=2.0",
                                         facecolor=bg_color, edgecolor=border_color, linewidth=1.5)
            ax.add_patch(box)
        elif shape == "diamond":
            diamond = patches.Polygon([
                [x + w/2, y], [x + w, y + h/2], [x + w/2, y + h], [x, y + h/2]
            ], facecolor=bg_color, edgecolor=border_color, linewidth=1.5)
            ax.add_patch(diamond)
        elif shape == "oval":
            box = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.8,rounding_size=6.0",
                                         facecolor=bg_color, edgecolor=border_color, linewidth=2.0)
            ax.add_patch(box)

        fw = 'bold' if bold else 'normal'
        ax.text(x + w/2, y + h/2, text, ha='center', va='center', fontsize=font_size,
                color=text_color, fontweight=fw, wrap=True)

    def draw_arrow(x1, y1, x2, y2, label=""):
        ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="->", color="#334155", lw=1.5, mutation_scale=14))
        if label:
            mx, my = (x1 + x2)/2, (y1 + y2)/2
            ax.text(mx + 1.5, my, label, fontsize=8, fontweight='bold', color="#0f766e")

    # Уровень 1: Исходное событие
    draw_box(30, 91, 40, 7, "НАЧАЛО: Возникновение нештатной ситуации в АС СН НПС\n(Отказ датчика / Обрыв связи / Сбой ПЛК / Пропадание ввода 220В)",
             bg_color="#f1f5f9", border_color="#475569", shape="oval", font_size=8.5, bold=True)
    draw_arrow(50, 91, 50, 84)

    # Уровень 2: Первичная реакция АСУ ТП и сигнализация
    draw_box(22, 77, 56, 7, "Автоматическая фиксация аварии ПЛК / SCADA:\nСрабатывание сигнализации (звук, мигание мнемосхемы, аварийный баннер)",
             bg_color="#fee2e2", border_color="#ef4444", shape="rect", bold=True)
    draw_arrow(50, 77, 50, 70)

    # Уровень 3: Действия дежурного оператора
    draw_box(18, 62, 64, 8, "Действия дежурного оператора НПС:\n1. Квитирование звуковой сигнализации (до 10 с)\n2. Оценка критичности параметра (давление на выкиде МНА, вибрация, загазованность)\n3. Проверка срабатывания технологической защиты (ПАЗ) и статуса агрегата",
             bg_color="#fef3c7", border_color="#f59e0b", shape="rect")
    draw_arrow(50, 62, 50, 55)

    # Уровень 4: Условие критичности
    draw_box(28, 46, 44, 9, "Угроза аварии или выхода за\nкритические уставки ПАЗ?\n(Давление > 6.4 МПа / Вибрация > 7.1 мм/с)",
             bg_color="#ffedd5", border_color="#ea580c", shape="diamond", bold=True, font_size=8)

    # Ветка ДА (Аварийный останов)
    draw_arrow(72, 50.5, 82, 50.5, "ДА")
    draw_box(80, 46, 18, 9, "Аварийный останов МНА:\nПеревод узла в\nSafe-State",
             bg_color="#fecdd3", border_color="#e11d48", shape="rect", bold=True, font_size=7.5)
    draw_arrow(89, 46, 89, 36)

    # Ветка НЕТ (Переход на резерв)
    draw_arrow(50, 46, 50, 39, "НЕТ")

    # Уровень 5: Вызов ремонтного персонала
    draw_box(15, 30, 70, 9, "Оповещение и вызов дежурного персонала:\n1. Фиксация в оперативном журнале времени и кода события\n2. Вызов инженера АСУ ТП и слесаря КИПиА (нормативное время прибытия <= 15 мин)\n3. При угрозе остановки МНА — доклад начальнику смены и диспетчеру РНУ",
             bg_color="#e0e7ff", border_color="#6366f1", shape="rect")
    draw_arrow(89, 36, 85, 36)
    draw_arrow(50, 30, 50, 23)

    # Уровень 6: Диагностика и устранение отказа
    draw_box(12, 14, 76, 9, "Проведение аварийно-восстановительных работ (АВР):\n- Локализация причины: датчик / кабельная линия / модуль ПЛК / сетевой свитч / питание\n- Замена дефектного компонента из фонда неснижаемого ЗИП\n- Калибровка, тестирование сигналов в SCADA и проверка контура ПАЗ",
             bg_color="#dcfce7", border_color="#16a34a", shape="rect")
    draw_arrow(50, 14, 50, 8)

    # Уровень 7: Завершение и документирование
    draw_box(25, 1, 50, 7, "КОНЕЦ: Ввод в эксплуатацию, оформление Акта расследования отказа,\nсписание ЗИП и закрытие наряда-допуска",
             bg_color="#f1f5f9", border_color="#475569", shape="oval", font_size=8.5, bold=True)

    ax.set_title("Блок-схема алгоритма действий оперативного персонала в нештатных ситуациях",
                 fontsize=13, fontweight='bold', pad=15)
    plt.tight_layout()
    out_path = os.path.join(SCRIPT_DIR, "pr5_incident_flowchart.png")
    plt.savefig(out_path, dpi=300)
    plt.close()
    print(f"  + Сохранен: {out_path}")
    return out_path

def generate_figure_4_equipment_tree():
    """Иерархическая структура комплекса технических средств (КТС) АС СН НПС."""
    fig, ax = plt.subplots(figsize=(14, 8), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    def draw_node(x, y, w, h, title, items, bg_color="#ffffff", border_color="#0284c7"):
        box = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.5,rounding_size=1.5",
                                     facecolor=bg_color, edgecolor=border_color, linewidth=1.5)
        ax.add_patch(box)
        ax.text(x + w/2, y + h - 2.5, title, ha='center', va='top', fontsize=9, fontweight='bold', color="#0f172a")

        item_str = "\n".join([f"• {it}" for it in items])
        ax.text(x + 1.5, y + h - 5.5, item_str, ha='left', va='top', fontsize=7.5, color="#334155")

    # Корневой узел
    root = patches.FancyBboxPatch((30, 88), 40, 9, boxstyle="round,pad=0.6,rounding_size=2.0",
                                  facecolor="#0f172a", edgecolor="#38bdf8", linewidth=2.0)
    ax.add_patch(root)
    ax.text(50, 92.5, "КОМПЛЕКС ТЕХНИЧЕСКИХ СРЕДСТВ АС СН НПС\n(Объект технического обслуживания и метрологического контроля)",
            ha='center', va='center', fontsize=9.5, fontweight='bold', color="#ffffff")

    # 4 ветви
    coords = [
        (2, 42, 22, 38, "1. ПОЛЕВОЙ УРОВЕНЬ\n(КИПиА и датчики)",
         ["Датчики давления: Метран-150", "Датчики температуры: Метран-280", "Датчики вибрации: ВК-310В", "УЗ расходомер: SeniorSonic", "Уровнемеры РВС: VEGAPULS", "Газоанализаторы: СГОЭС-М11"],
         "#f0fdf4", "#22c55e"),
        (26.5, 42, 22, 38, "2. ИСПОЛНИТЕЛЬНЫЙ\nИ СИЛОВОЙ УРОВЕНЬ",
         ["Э/приводы Auma SA 14.2", "Блоки управления AC 01.2", "Клапаны сброса Mokveld", "Отсечные задвижки ПАЗ", "Насосы откачки утечек", "Регулирующие заслонки"],
         "#eff6ff", "#3b82f6"),
        (51.5, 42, 22, 38, "3. КОНТРОЛЛЕРНЫЙ\nИ КРОССОВЫЙ УРОВЕНЬ",
         ["ПЛК РСУ: Regul R500 (CPU 01)", "ПЛК ПАЗ: Regul R500S (SIL3)", "Модули AI-08-H, DI-16, DO-16", "Шкафы автоматики Rittal IP54", "Кроссовые шкафы КШ", "Барьеры искрозащиты"],
         "#faf5ff", "#a855f7"),
        (76, 42, 22, 38, "4. СЕТЕВОЙ, ПИТАЮЩИЙ\nИ ВЕРХНИЙ УРОВЕНЬ",
         ["Серверы MasterSCADA 4D", "АРМ оператора и инженера", "Коммутаторы MOXA EDS-510E", "Оптоволокно кольцо Turbo Ring", "ИБП Штиль/APC 10кВА On-Line", "АКБ AGM 16x12В 55А·ч"],
         "#fff7ed", "#f97316"),
    ]

    # Линии от корня к ветвям
    ax.plot([50, 50], [88, 83], color="#64748b", lw=1.5)
    ax.plot([13, 87], [83, 83], color="#64748b", lw=1.5)
    for cx, cy, cw, ch, title, items, bg, border in coords:
        top_x = cx + cw/2
        ax.plot([top_x, top_x], [83, 80], color="#64748b", lw=1.5)
        draw_node(cx, cy, cw, ch, title, items, bg_color=bg, border_color=border)

    # Нижний блок регламентов и зон ответственности
    recs = [
        (2, 6, 22, 30, "РЕГЛАМЕНТ КИПиА",
         ["ТО-1: раз в 2 нед. (осмотр)", "ТО-2: раз в месяц (ноль)", "ТО-4: раз в год (поверка)", "Отв.: Слесарь КИПиА,", "Инженер-метролог"],
         "#dcfce7", "#16a34a"),
        (26.5, 6, 22, 30, "РЕГЛАМЕНТ ПРИВОДОВ",
         ["ТО-1: раз в 2 нед. (смазка)", "ТО-2: раз в мес. (тест хода)", "ТО-3: квартал (моменты)", "ТО-4: год (капревизия)", "Отв.: Электромонтер"],
         "#dbeafe", "#2563eb"),
        (51.5, 6, 22, 30, "РЕГЛАМЕНТ ПЛК",
         ["ТО-2: раз в месяц (диагн.)", "ТО-3: квартал (продувка)", "ТО-4: год (бекап, софт)", "Тестирование схем ПАЗ", "Отв.: Инженер АСУ ТП"],
         "#f3e8ff", "#9333ea"),
        (76, 6, 22, 30, "РЕГЛАМЕНТ ИТ И СЕТЕЙ",
         ["ТО-2: раз в мес. (тест АКБ)", "ТО-3: квартал (коммутаторы)", "ТО-4: год (серверы, БД)", "Тест перехода на ИБП", "Отв.: Инженер АСУ / Эл."],
         "#ffedd5", "#ea580c"),
    ]
    for idx, (rcx, rcy, rcw, rch, rtitle, ritems, rbg, rborder) in enumerate(recs):
        top_x = rcx + rcw/2
        ax.plot([top_x, top_x], [42, 36], color="#94a3b8", lw=1.2, linestyle=":")
        draw_node(rcx, rcy, rcw, rch, rtitle, ritems, bg_color=rbg, border_color=rborder)

    ax.set_title("Иерархическая структура КТС АС СН НПС и закрепление регламентов ТО",
                 fontsize=13, fontweight='bold', pad=15)
    plt.tight_layout()
    out_path = os.path.join(SCRIPT_DIR, "pr5_equipment_tree.png")
    plt.savefig(out_path, dpi=300)
    plt.close()
    print(f"  + Сохранен: {out_path}")
    return out_path

# =============================================================================
# 4. СОЗДАНИЕ РАСЧЕТНОЙ EXCEL-КНИГИ СО ВСЕМИ ТАБЛИЦАМИ
# =============================================================================

def build_excel_workbook(stats):
    """Формирует книгу Excel: ПР5_план_и_регламент_ТО_НПС.xlsx."""
    wb = openpyxl.Workbook()
    # Удаляем дефолтный лист
    wb.remove(wb.active)

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    sub_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    alt_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # -------------------------------------------------------------------------
    # Лист 1: Перечень_Оборудования
    # -------------------------------------------------------------------------
    ws1 = wb.create_sheet(title="Перечень_Оборудования")
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:G1")
    ws1["A1"] = "ПЕРЕЧЕНЬ ОБОРУДОВАНИЯ И ТЕХНИЧЕСКИХ СРЕДСТВ АС СН НПС (ИНВЕНТАРИЗАЦИЯ)"
    ws1["A1"].font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 30

    headers1 = ["Поз. ID", "Наименование оборудования", "Тип / Модель", "Категория КТС", "Кол-во, шт.", "Место установки", "Критичность в АС СН"]
    ws1.append([])
    ws1.append(headers1)
    ws1.row_dimensions[3].height = 25

    for col_idx in range(1, len(headers1) + 1):
        cell = ws1.cell(row=3, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for eq in EQUIPMENT_DATA:
        row_vals = [
            eq["id"], eq["name"], eq["model"], eq["category"], eq["qty"], eq["location"], eq["criticality"]
        ]
        ws1.append(row_vals)
        curr_row = ws1.max_row
        ws1.row_dimensions[curr_row].height = 20
        for col_idx in range(1, len(row_vals) + 1):
            c = ws1.cell(row=curr_row, column=col_idx)
            c.font = Font(name="Calibri", size=10)
            c.border = thin_border
            if col_idx in [1, 5]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
            if curr_row % 2 == 1:
                c.fill = alt_fill

    # -------------------------------------------------------------------------
    # Лист 2: Матрица_ТО
    # -------------------------------------------------------------------------
    ws2 = wb.create_sheet(title="Матрица_ТО")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:H1")
    ws2["A1"] = "ВЕДОМОСТЬ ВИДОВ, ПЕРИОДИЧНОСТИ И ТРУДОЕМКОСТИ ТЕХНИЧЕСКОГО ОБСЛУЖИВАНИЯ"
    ws2["A1"].font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 30

    headers2 = ["Поз. ID", "Наименование оборудования", "Вид ТО", "Периодичность", "Кол-во ед.", "Трудоемкость ед., чел·ч", "Годовая трудоемкость, чел·ч", "Исполнитель"]
    ws2.append([])
    ws2.append(headers2)
    ws2.row_dimensions[3].height = 25

    for col_idx in range(1, len(headers2) + 1):
        cell = ws2.cell(row=3, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row in stats["table_rows"]:
        vals = [
            row["id"], row["name"], row["m_type"], row["freq_label"],
            row["qty"], row["labor_single"], row["labor_annual"], row["role"]
        ]
        ws2.append(vals)
        curr_row = ws2.max_row
        ws2.row_dimensions[curr_row].height = 19
        for col_idx in range(1, len(vals) + 1):
            c = ws2.cell(row=curr_row, column=col_idx)
            c.font = Font(name="Calibri", size=10)
            c.border = thin_border
            if col_idx in [1, 3, 4, 5]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [6, 7]:
                c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
            if curr_row % 2 == 1:
                c.fill = alt_fill

    # Итоговая строка
    tot_row = ws2.max_row + 1
    ws2.cell(row=tot_row, column=2, value="ИТОГО ГОДОВАЯ ТРУДОЕМКОСТЬ ТО:").font = Font(name="Calibri", size=11, bold=True)
    sum_cell = ws2.cell(row=tot_row, column=7, value=f"=SUM(G4:G{tot_row-1})")
    sum_cell.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    sum_cell.border = thin_border

    # -------------------------------------------------------------------------
    # Лист 3: Годовой_График_ТО
    # -------------------------------------------------------------------------
    ws3 = wb.create_sheet(title="Годовой_График_ТО")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:O1")
    ws3["A1"] = "ГОДОВОЙ КАЛЕНДАРНЫЙ ПЛАН-ГРАФИК ВЫПОЛНЕНИЯ ТО ПО МЕСЯЦАМ (В ЧЕЛ·Ч)"
    ws3["A1"].font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 30

    headers3 = ["Специальность / Должность", "Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек", "ИТОГО ГОД", "Штат, чел."]
    ws3.append([])
    ws3.append(headers3)
    ws3.row_dimensions[3].height = 25

    for col_idx in range(1, len(headers3) + 1):
        cell = ws3.cell(row=3, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    staff_norm = {"Слесарь КИПиА": 2, "Инженер АСУ ТП": 2, "Инженер-метролог": 1, "Электромонтер": 1}

    for role, m_hours in stats["monthly_by_role"].items():
        row_vals = [role] + [round(h, 1) for h in m_hours] + [round(sum(m_hours), 1), staff_norm.get(role, 1)]
        ws3.append(row_vals)
        curr_row = ws3.max_row
        ws3.row_dimensions[curr_row].height = 20
        for col_idx in range(1, len(row_vals) + 1):
            c = ws3.cell(row=curr_row, column=col_idx)
            c.font = Font(name="Calibri", size=10)
            c.border = thin_border
            if col_idx == 1:
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.font = Font(name="Calibri", size=10, bold=True)
            else:
                c.alignment = Alignment(horizontal="right", vertical="center")
            if curr_row % 2 == 1:
                c.fill = alt_fill

    # Итоговая строка по месяцам
    tot_row3 = ws3.max_row + 1
    ws3.cell(row=tot_row3, column=1, value="СУММАРНАЯ НАГРУЗКА (чел·ч):").font = Font(name="Calibri", size=11, bold=True)
    for col_idx in range(2, 14):
        col_letter = get_column_letter(col_idx)
        cell = ws3.cell(row=tot_row3, column=col_idx, value=f"=SUM({col_letter}4:{col_letter}{tot_row3-1})")
        cell.font = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="right", vertical="center")

    gt_cell = ws3.cell(row=tot_row3, column=14, value=f"=SUM(N4:N{tot_row3-1})")
    gt_cell.font = Font(name="Calibri", size=11, bold=True, color="B91C1C")
    gt_cell.border = thin_border
    gt_cell.alignment = Alignment(horizontal="right", vertical="center")

    # -------------------------------------------------------------------------
    # Лист 4: Нормы_ЗИП_и_Приборы
    # -------------------------------------------------------------------------
    ws4 = wb.create_sheet(title="Нормы_ЗИП_и_Приборы")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:G1")
    ws4["A1"] = "ВЕДОМОСТЬ НЕСНИЖАЕМОГО ЗАПАСА ЗИП И ОБРАЗЦОВЫХ СРЕДСТВ ИЗМЕРЕНИЙ"
    ws4["A1"].font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    ws4["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 30

    headers4 = ["Тип ресурса", "Наименование изделия / прибора", "Модель / Марка", "Назначение", "Норма запаса, шт.", "Срок пополнения, сут.", "Место хранения"]
    ws4.append([])
    ws4.append(headers4)
    ws4.row_dimensions[3].height = 25

    for col_idx in range(1, len(headers4) + 1):
        cell = ws4.cell(row=3, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    zip_items = [
        ("ЗИП датчики", "Датчик избыточного давления", "Метран-150-TG (0..6 МПа)", "Аварийная замена на входе/выходе МНА", 3, 5, "Склад КИПиА"),
        ("ЗИП датчики", "Термопреобразователь сопротивления", "Метран-280 ТСП Pt100", "Контроль подшипников МНА", 4, 3, "Склад КИПиА"),
        ("ЗИП датчики", "Датчик виброскорости", "ВК-310В (4..20 мА)", "Мониторинг вибрации МНА", 2, 7, "Склад КИПиА"),
        ("ЗИП датчики", "Оптический датчик метана", "СГОЭС-М11 (0..100% НКПР)", "Контроль загазованности насосной", 3, 5, "Склад КИПиА"),
        ("ЗИП ПЛК", "Процессорный модуль дублированный", "Regul CPU 01", "Резерв для ПЛК РСУ/ПАЗ", 1, 14, "Серверная АСУ ТП"),
        ("ЗИП ПЛК", "Модуль аналогового ввода 8xAI HART", "Regul AI-08-H", "Замена вышедших из строя каналов", 2, 7, "Серверная АСУ ТП"),
        ("ЗИП ПЛК", "Модуль дискретного ввода 16xDI", "Regul DI-16", "Замена каналов сигнализации", 2, 7, "Серверная АСУ ТП"),
        ("ЗИП Питание", "Импульсный блок питания 24В DC 10A", "Phoenix Contact Quint 24V/10A", "Питание цепей датчиков и модулей", 2, 3, "Склад КИПиА"),
        ("ЗИП Питание", "Аккумуляторная батарея 12В 55А·ч", "Delta HRL 12-55 AGM", "Замена деградировавших банок ИБП", 4, 5, "ШГП / Аккумуляторная"),
        ("ЗИП Кабели", "Кабель сигнальный экранированный", "МКЭШВнг(А)-FRLS 2x2x0.75", "Ремонт искробезопасных трасс", "200 м", 2, "Склад КИПиА"),
        ("ЗИП Защита", "Предохранители быстродействующие", "5x20 мм 0.5A, 1A, 2A, 5A", "Защита клеммных цепей питания", 50, 1, "Шкафы автоматики"),
        ("Эталон СИ", "Прецизионный калибратор давления", "Метран-517 / Fluke 718", "Поверка и калибровка датчиков давления", 1, "-", "Лаборатория метрологии"),
        ("Эталон СИ", "Калибратор токовой петли HART", "Fluke 709H / Метран-650", "Проверка аналоговых каналов 4-20 мА", 2, "-", "Лаборатория метрологии"),
        ("Эталон СИ", "Цифровой мегомметр 2500В", "MIC-2500 / Sonel", "Измерение сопротивления изоляции кабелей", 1, "-", "Мастерская КИПиА"),
        ("Эталон СИ", "Тепловизор диагностический", "Testo 875-1i", "Контроль перегрева клемм и плат ПЛК", 1, "-", "Служба АСУ ТП"),
    ]

    for item in zip_items:
        ws4.append(list(item))
        curr_row = ws4.max_row
        ws4.row_dimensions[curr_row].height = 19
        for col_idx in range(1, len(item) + 1):
            c = ws4.cell(row=curr_row, column=col_idx)
            c.font = Font(name="Calibri", size=10)
            c.border = thin_border
            if col_idx in [1, 5, 6]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
            if curr_row % 2 == 1:
                c.fill = alt_fill

    # -------------------------------------------------------------------------
    # Лист 5: Журнал_ТО_и_Аварий
    # -------------------------------------------------------------------------
    ws5 = wb.create_sheet(title="Журнал_ТО_и_Аварий")
    ws5.views.sheetView[0].showGridLines = True

    ws5.merge_cells("A1:H1")
    ws5["A1"] = "ТИПОВОЙ ОПЕРАТИВНЫЙ ЖУРНАЛ РЕГИСТРАЦИИ ТЕХНИЧЕСКОГО ОБСЛУЖИВАНИЯ И ОТКАЗОВ"
    ws5["A1"].font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    ws5["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 30

    headers5 = ["№ п/п", "Дата и время", "Поз. ID оборудования", "Вид ТО / Характер инцидента", "Выполненные работы / Принятые меры", "Заключение о пригодности", "ФИО исполнителя", "Подпись / Отметка"]
    ws5.append([])
    ws5.append(headers5)
    ws5.row_dimensions[3].height = 25

    for col_idx in range(1, len(headers5) + 1):
        cell = ws5.cell(row=3, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    sample_logs = [
        (1, "12.01.2026 09:30", "PT-101", "ТО-1 (Плановое)", "Внешний осмотр, проверка герметичности импульсных трубок, контроль нуля. Отклонение 0.05%", "Годен к эксплуатации", "Гайнанов Д.И.", "Выполнено"),
        (2, "19.01.2026 14:15", "VT-301", "Нештатная ситуация", "Ложное срабатывание предупредительной уставки (дребезг контактов клеммы). Протяжка клеммника", "Годен к эксплуатации", "Гарифуллин К.Р.", "Устранено"),
        (3, "02.02.2026 10:00", "PLC-01", "ТО-2 (Плановое)", "Диагностика дублированных CPU, проверка логов MasterSCADA, тестирование резервирования", "Годен к эксплуатации", "Гарифуллин К.Р.", "Выполнено"),
        (4, "15.02.2026 11:45", "ACT-101", "ТО-1 (Плановое)", "Проверка указателя положения задвижки Auma, контроль концевых микровыключателей", "Годен к эксплуатации", "Гайнанов Д.И.", "Выполнено"),
        (5, "18.05.2026 08:30", "PT-102", "ТО-4 (Поверка)", "Демонтаж, поверка на калибраторе Метран-517 в 5 точках шкалы, погрешность 0.08%. Монтаж", "Поверен (Свидетельство)", "Метролог В.А.", "Поверен"),
        (6, "19.05.2026 13:00", "UPS-01", "ТО-4 (Регламент)", "Тестирование емкости батарей под нагрузкой, протяжка силовых клемм, очистка пыли", "Годен к эксплуатации", "Гайнанов Д.И.", "Выполнено"),
    ]

    for log in sample_logs:
        ws5.append(list(log))
        curr_row = ws5.max_row
        ws5.row_dimensions[curr_row].height = 22
        for col_idx in range(1, len(log) + 1):
            c = ws5.cell(row=curr_row, column=col_idx)
            c.font = Font(name="Calibri", size=10)
            c.border = thin_border
            if col_idx in [1, 2, 3, 6, 8]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
            if curr_row % 2 == 1:
                c.fill = alt_fill

    # Автоподбор ширины столбцов во всех листах
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row == 1:  # заголовок объединения
                    continue
                val = str(cell.value or '')
                if len(val) > max_len:
                    max_len = len(val)
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 11)

    out_excel = os.path.join(SCRIPT_DIR, "ПР5_план_и_регламент_ТО_НПС.xlsx")
    wb.save(out_excel)
    print(f"  + Сохранена расчетная книга Excel: {out_excel}")
    return out_excel

# =============================================================================
# ТОЧКА ВХОДА
# =============================================================================

if __name__ == "__main__":
    print("===========================================================================")
    print("  РАСЧЕТНО-АНАЛИТИЧЕСКИЙ МОДУЛЬ ПР5 (АСУСН)")
    print("  План технического обслуживания и регламент АС СН НПС")
    print("  Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С")
    print("===========================================================================")

    print("\n[1/3] Расчет трудоемкости, фондов времени и распределения по месяцам...")
    stats = calculate_labor_statistics()
    print(f"  + Итоговая годовая трудоемкость: {stats['grand_total']:.1f} чел·ч")
    for r, h in stats["total_by_role"].items():
        print(f"    - {r}: {h:.1f} чел·ч ({h/stats['grand_total']*100:.1f}%)")

    print("\n[2/3] Построение 4 графических диаграмм высокого разрешения (DPI 300)...")
    generate_figure_1_gantt(stats)
    generate_figure_2_workload(stats)
    generate_figure_3_incident_flowchart()
    generate_figure_4_equipment_tree()

    print("\n[3/3] Генерация многостраничной расчетной книги Excel...")
    build_excel_workbook(stats)

    print("\n===========================================================================")
    print("  МОДЕЛИРОВАНИЕ И РАСЧЕТЫ ПР5 УСПЕШНО ЗАВЕРШЕНЫ!")
    print("===========================================================================")

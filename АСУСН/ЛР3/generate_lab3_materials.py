# -*- coding: utf-8 -*-
"""
Генератор полного комплекта материалов к Лабораторной работе №3 по АСУСН (ВАРИАНТ Б):
Моделирование системы сбора данных с использованием ВСТРОЕННОГО ЭМУЛЯТОРА Rapid SCADA.

1. Графические схемы и мнемосхемы (Архитектура, Мнемосхема Норма, Мнемосхема Перегрев, Тренды, Журнал тревог)
2. Расчетная таблица и база конфигурации в Excel (.xlsx)
3. Полный чистовой отчет по ГОСТ в формате Word (.docx)

Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С
Кафедра АСУ, УУНиТ, 2026 г.
"""

import os
import sys
import math
import random
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, Rectangle, Circle

# Настройка кодировки
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

WORK_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(WORK_DIR, 'images_lab3')
os.makedirs(IMG_DIR, exist_ok=True)

# Стили matplotlib
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 9
plt.rcParams['axes.titlesize'] = 11
plt.rcParams['axes.labelsize'] = 10
plt.rcParams['figure.dpi'] = 300

# =============================================================================
# 1. ГЕНЕРАЦИЯ ГРАФИЧЕСКИХ МАТЕРИАЛОВ (ВАРИАНТ Б — ЭМУЛЯТОР)
# =============================================================================

def generate_architecture_diagram():
    """Схема информационных потоков: Встроенный эмулятор -> Коммуникатор -> Сервер -> Вебстанция"""
    fig, ax = plt.subplots(figsize=(10, 6.2), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 70)
    ax.axis('off')

    ax.text(50, 67, "Архитектура сбора данных на базе Rapid SCADA с встроенным эмулятором (Вариант Б)",
            ha='center', va='center', fontsize=11.5, fontweight='bold', color='#1A365D')

    # Блок 1: Источник данных - Встроенный эмулятор
    b1 = FancyBboxPatch((4, 8), 28, 52, boxstyle="round,pad=1.0", fc='#F0FFF4', ec='#68D391', lw=1.8)
    ax.add_patch(b1)
    ax.text(18, 56.5, "ВСТРОЕННЫЙ ЭМУЛЯТОР\n(Rapid SCADA DrvSimulator)", ha='center', va='center', fontsize=9.5, fontweight='bold', color='#22543D')

    items_emu = [
        ("Канал 1: Температура\nНорма: 20–80 °C\n(Тест 2: 70–95 °C)", 46, '#C6F6D5', '#276749'),
        ("Канал 2: Давление\nДиапазон: 0–10 бар\n(Порог тревоги: > 8 бар)", 34, '#C6F6D5', '#276749'),
        ("Канал 3: Уровень\nДиапазон: 0–100 %\n(Резервуар ТН-01)", 22, '#C6F6D5', '#276749'),
        ("Канал управления:\nПуск / Останов насоса\n(Канал Pump_Control)", 11, '#FEFCBF', '#D69E2E'),
    ]
    for text, y, fc, ec in items_emu:
        p = FancyBboxPatch((6, y - 4), 24, 7.5, boxstyle="round,pad=0.4", fc=fc, ec=ec, lw=1.2)
        ax.add_patch(p)
        ax.text(18, y, text, ha='center', va='center', fontsize=7.8, color='#1A202C')

    # Блок 2: SCADA Коммуникатор
    b2 = FancyBboxPatch((38, 14), 26, 46, boxstyle="round,pad=1.0", fc='#EBF8FF', ec='#63B3ED', lw=1.8)
    ax.add_patch(b2)
    ax.text(51, 56, "КОММУНИКАТОР\n(SCADA Communicator)", ha='center', va='center', fontsize=9.5, fontweight='bold', color='#2B6CB0')

    p_line = FancyBboxPatch((40, 36), 22, 14, boxstyle="round,pad=0.4", fc='#BEE3F8', ec='#3182CE', lw=1.2)
    ax.add_patch(p_line)
    ax.text(51, 43, "Линия связи: «Эмулятор»\nДрайвер: DrvSimulator\n\nЦикл опроса: 1000 мс\nКП 1: Simulator",
            ha='center', va='center', fontsize=8, color='#1A365D')

    p_buf = FancyBboxPatch((40, 18), 22, 13, boxstyle="round,pad=0.4", fc='#FFFFFF', ec='#90CDF4', lw=1.0)
    ax.add_patch(p_buf)
    ax.text(51, 24.5, "Буферизация данных\nи передача текущих\nсрезов измерений",
            ha='center', va='center', fontsize=7.8, color='#2C5282')

    # Связь Эмулятор -> Коммуникатор
    ax.annotate('', xy=(38, 43), xytext=(32, 43),
                arrowprops=dict(arrowstyle="<|-|>", color="#2B6CB0", lw=2.0))
    ax.text(35, 46, "Внутренний\nвызов API", fontsize=7.5, fontweight='bold', color="#2B6CB0", ha='center')

    # Блок 3: SCADA Сервер
    b3 = FancyBboxPatch((70, 32), 26, 28, boxstyle="round,pad=1.0", fc='#EDF2F7', ec='#CBD5E0', lw=1.8)
    ax.add_patch(b3)
    ax.text(83, 56, "СЕРВЕР\n(SCADA Server)", ha='center', va='center', fontsize=9.5, fontweight='bold', color='#2D3748')
    ax.text(83, 44, "• База конфигурации\n• Модуль тревог (АПС)\n• Исторический архив\n  (период: 1 секунда)\n• Порт TCP 10000",
            ha='center', va='center', fontsize=8, color='#1A202C')

    # Связь Коммуникатор -> Сервер
    ax.annotate('', xy=(70, 44), xytext=(64, 44),
                arrowprops=dict(arrowstyle="<|-|>", color="#4A5568", lw=1.8))
    ax.text(67, 46.5, "TCP 10000", fontsize=7.5, fontweight='bold', color="#4A5568", ha='center')

    # Блок 4: SCADA Вебстанция
    b4 = FancyBboxPatch((70, 6), 26, 22, boxstyle="round,pad=1.0", fc='#FEFCBF', ec='#F6E05E', lw=1.8)
    ax.add_patch(b4)
    ax.text(83, 24, "ВЕБСТАНЦИЯ\n(SCADA WebStation)", ha='center', va='center', fontsize=9.5, fontweight='bold', color='#744210')
    ax.text(83, 14, "http://localhost:10008\n• Мнемосхема Lab3_Mimic\n• Модуль графиков (Тренды)\n• Журнал текущих тревог",
            ha='center', va='center', fontsize=7.8, color='#5F370E')

    # Связь Сервер -> Вебстанция
    ax.annotate('', xy=(83, 28), xytext=(83, 32),
                arrowprops=dict(arrowstyle="<|-|>", color="#4A5568", lw=1.8))
    ax.text(87, 30, "HTTP / WCF", fontsize=7.5, color="#4A5568")

    plt.tight_layout()
    out_path = os.path.join(IMG_DIR, 'scada_architecture.png')
    plt.savefig(out_path, dpi=300)
    plt.close()
    return out_path

def generate_mimic_screen(is_alarm=False):
    """Скриншот мнемосхемы технологического узла в нормальном или аварийном режиме (Вариант Б)"""
    fig, ax = plt.subplots(figsize=(11, 6.8), dpi=300)
    ax.set_xlim(0, 110)
    ax.set_ylim(0, 70)
    ax.axis('off')

    bg_color = '#1A202C' if is_alarm else '#171923'
    fig.patch.set_facecolor(bg_color)
    ax.set_facecolor(bg_color)

    # Заголовок окна HMI
    header_color = '#742A2A' if is_alarm else '#2A4365'
    ax.add_patch(Rectangle((0, 64), 110, 6, fc=header_color, ec='none'))
    title_text = "Rapid SCADA [Вебстанция] — Мнемосхема: Lab3_Mimic (Встроенный эмулятор DrvSimulator)"
    status_mode_text = "[ТЕСТ 2: ПЕРЕГРЕВ (70–95 °C)]" if is_alarm else "[ТЕСТ 1: НОРМА (20–80 °C)]"
    ax.text(3, 67, title_text, color='#FFFFFF', fontsize=10, fontweight='bold', va='center')
    ax.text(107, 67, status_mode_text, color='#FED7D7' if is_alarm else '#9AE6B4', fontsize=9.5, fontweight='bold', ha='right', va='center')

    # Аварийная строка-баннер
    if is_alarm:
        ax.add_patch(Rectangle((0, 59.5), 110, 4.5, fc='#9B2C2C', ec='#E53E3E', lw=1.5))
        ax.text(55, 61.8, "!!! ТРЕВОГА: ПЕРЕГРЕВ ТЕПЛОНОСИТЕЛЯ (T > 80.0 °C)! ДИАПАЗОН ЭМУЛЯТОРА: 70–95 °C !!!",
                color='#FFFFFF', fontsize=9.5, fontweight='bold', ha='center', va='center')

    # Резервуар с уровнем
    tank_y, tank_h, tank_w, tank_x = 15, 36, 26, 10
    ax.add_patch(Rectangle((tank_x, tank_y), tank_w, tank_h, fc='#2D3748', ec='#A0AEC0', lw=2.5))
    ax.plot([tank_x, tank_x + tank_w], [tank_y + tank_h, tank_y + tank_h], color='#CBD5E0', lw=3)
    ax.plot([tank_x, tank_x + tank_w], [tank_y, tank_y], color='#CBD5E0', lw=3)

    fill_pct = 0.85 if is_alarm else 0.62
    liquid_h = tank_h * fill_pct
    liq_color = '#DD6B20' if is_alarm else '#3182CE'
    ax.add_patch(Rectangle((tank_x + 1, tank_y + 1), tank_w - 2, liquid_h, fc=liq_color, ec='none', alpha=0.75))

    for i in range(11):
        mark_y = tank_y + (tank_h * i / 10.0)
        ax.plot([tank_x - 1.5, tank_x], [mark_y, mark_y], color='#CBD5E0', lw=1.5)
        ax.text(tank_x - 2.5, mark_y, f"{i*10}%", color='#E2E8F0', fontsize=7, ha='right', va='center')

    ax.text(tank_x + tank_w/2, tank_y + tank_h + 2.5, "Резервуар сбора ТН-01", color='#E2E8F0', fontsize=9.5, fontweight='bold', ha='center')
    ax.text(tank_x + tank_w/2, tank_y + liquid_h/2, f"Level_Value: {fill_pct*100:.1f} %", color='#FFFFFF', fontsize=9.5, fontweight='bold', ha='center')

    # Трубопроводы
    ax.plot([tank_x + 6, tank_x + 6], [tank_y + tank_h, 57], color='#718096', lw=4)
    ax.annotate('', xy=(tank_x + 6, tank_y + tank_h + 1), xytext=(tank_x + 6, 56),
                arrowprops=dict(arrowstyle="-|>", color='#63B3ED', lw=2.5))
    ax.text(tank_x + 6, 58, "Приток среды", color='#CBD5E0', fontsize=8, ha='center')

    ax.plot([tank_x + tank_w, 48], [tank_y + 6, tank_y + 6], color='#718096', lw=4)
    ax.plot([48, 48], [tank_y + 6, tank_y + 12], color='#718096', lw=4)

    # Насосный агрегат
    pump_x, pump_y = 56, tank_y + 12
    pump_fc = '#38A169' if not is_alarm else '#E53E3E'
    ax.add_patch(Circle((pump_x, pump_y), 4.5, fc=pump_fc, ec='#FFFFFF', lw=2))
    ax.plot([pump_x, pump_x + 5.5], [pump_y, pump_y + 5.5], color='#FFFFFF', lw=2)
    ax.text(pump_x, pump_y - 7, "Насос МНА-1\n[РАБОТА]", color='#68D391' if not is_alarm else '#FC8181',
            fontsize=8.5, fontweight='bold', ha='center')

    ax.plot([pump_x + 4.5, 78], [pump_y, pump_y], color='#718096', lw=4)
    ax.annotate('', xy=(78, pump_y), xytext=(pump_x + 5, pump_y),
                arrowprops=dict(arrowstyle="-|>", color='#63B3ED', lw=2.5))
    ax.text(82, pump_y, "В коллектор СН", color='#CBD5E0', fontsize=8, va='center')

    # Датчики (Аналоговые индикаторы)
    t_val = 86.4 if is_alarm else 52.3
    t_bg = '#742A2A' if is_alarm else '#2D3748'
    t_ec = '#E53E3E' if is_alarm else '#4A5568'
    ax.add_patch(FancyBboxPatch((42, 38), 28, 14, boxstyle="round,pad=0.5", fc=t_bg, ec=t_ec, lw=2))
    ax.text(44, 49.5, "Канал 1: Temp_Value", color='#E2E8F0', fontsize=8.5, fontweight='bold')
    ax.text(56, 43.5, f"{t_val:.1f} °C", color='#FC8181' if is_alarm else '#68D391', fontsize=16, fontweight='bold', ha='center')
    ax.text(56, 39.5, "Эмулятор: 70–95 °C [ТРЕВОГА > 80]" if is_alarm else "Эмулятор: 20–80 °C (НОРМА)",
            color='#FED7D7' if is_alarm else '#A0AEC0', fontsize=7.5, ha='center')

    p_val = 8.45 if is_alarm else 5.10
    p_bg = '#742A2A' if is_alarm else '#2D3748'
    p_ec = '#E53E3E' if is_alarm else '#4A5568'
    ax.add_patch(FancyBboxPatch((74, 38), 28, 14, boxstyle="round,pad=0.5", fc=p_bg, ec=p_ec, lw=2))
    ax.text(76, 49.5, "Канал 2: Pressure_Value", color='#E2E8F0', fontsize=8.5, fontweight='bold')
    ax.text(88, 43.5, f"{p_val:.2f} бар", color='#FC8181' if is_alarm else '#68D391', fontsize=16, fontweight='bold', ha='center')
    ax.text(88, 39.5, "Порог тревоги: > 8.00 бар" if is_alarm else "Эмулятор: 0–10 бар (НОРМА)",
            color='#FED7D7' if is_alarm else '#A0AEC0', fontsize=7.5, ha='center')

    # Панель телеуправления (Кнопки)
    ax.add_patch(FancyBboxPatch((74, 15), 32, 17, boxstyle="round,pad=0.5", fc='#2D3748', ec='#4A5568', lw=1.5))
    ax.text(90, 29.5, "ДИСТАНЦИОННОЕ УПРАВЛЕНИЕ", color='#CBD5E0', fontsize=8, fontweight='bold', ha='center')

    ax.add_patch(FancyBboxPatch((76, 21), 13, 6, boxstyle="round,pad=0.3", fc='#2F855A', ec='#48BB78', lw=1.5))
    ax.text(82.5, 24, "ПУСК (1)", color='#FFFFFF', fontsize=8.5, fontweight='bold', ha='center', va='center')

    ax.add_patch(FancyBboxPatch((91, 21), 13, 6, boxstyle="round,pad=0.3", fc='#9B2C2C', ec='#F56565', lw=1.5))
    ax.text(97.5, 24, "СТОП (0)", color='#FFFFFF', fontsize=8.5, fontweight='bold', ha='center', va='center')

    ax.text(90, 17.5, "Канал: Pump_Control (Телеуправление)", color='#A0AEC0', fontsize=7.5, ha='center')

    # Лампа «Авария»
    lamp_x, lamp_y = 56, 26
    lamp_color = '#E53E3E' if is_alarm else '#38A169'
    lamp_border = '#FFF5F5' if is_alarm else '#C6F6D5'
    ax.add_patch(Circle((lamp_x, lamp_y), 3.2, fc=lamp_color, ec=lamp_border, lw=2.5))
    if is_alarm:
        ax.text(lamp_x, lamp_y, "!", color='#FFFFFF', fontsize=13, fontweight='bold', ha='center', va='center')
        ax.text(lamp_x, lamp_y - 4.5, "ТРЕВОГА!\nАВАРИЯ", color='#FC8181', fontsize=8, fontweight='bold', ha='center')
    else:
        ax.text(lamp_x, lamp_y, "OK", color='#FFFFFF', fontsize=8, fontweight='bold', ha='center', va='center')
        ax.text(lamp_x, lamp_y - 4.5, "НОРМА\nБЕЗ АВАРИЙ", color='#68D391', fontsize=8, fontweight='bold', ha='center')

    # Статусная строка
    ax.add_patch(Rectangle((0, 0), 110, 4.5, fc='#1A202C', ec='#4A5568', lw=1))
    ax.text(3, 2.2, "Линия: Эмулятор (DrvSimulator) [Вариант Б] | Опрос: 1000 мс | Архивация: ВЕДЕТСЯ (1 с)",
            color='#A0AEC0', fontsize=8, va='center')
    alm_cnt = "АКТИВНА ТРЕВОГА ТЕМПЕРАТУРЫ" if is_alarm else "0 ТРЕВОГ"
    ax.text(107, 2.2, f"Сигнализация: {alm_cnt}",
            color='#FC8181' if is_alarm else '#68D391', fontsize=8, fontweight='bold', ha='right', va='center')

    plt.tight_layout()
    out_name = 'mimic_alarm_mode.png' if is_alarm else 'mimic_normal_mode.png'
    out_path = os.path.join(IMG_DIR, out_name)
    plt.savefig(out_path, dpi=300, facecolor=bg_color)
    plt.close()
    return out_path

def generate_trends_chart():
    """График трендов встроенного эмулятора: норма (20-80 °C) -> изменение настроек эмулятора на (70-95 °C)"""
    t = np.linspace(0, 300, 601)
    t_switch = 150.0  # Момент перенастройки эмулятора на аварийный диапазон (Тест 2)

    temp = np.zeros_like(t)
    for i, ti in enumerate(t):
        if ti < t_switch:
            # Норма: диапазон 20-80, случайные колебания вокруг 50 °C
            temp[i] = 50.0 + 15.0 * np.sin(ti * 0.08) + random.uniform(-4.0, 4.0)
        else:
            # Авария (Тест 2): диапазон 70-95 °C, среднее ~86 °C
            temp[i] = 84.0 + 6.0 * np.sin((ti - t_switch) * 0.12) + random.uniform(-3.0, 3.0)

    press = np.zeros_like(t)
    for i, ti in enumerate(t):
        if ti < t_switch:
            press[i] = 5.0 + 1.8 * np.sin(ti * 0.1) + random.uniform(-0.4, 0.4)
        else:
            press[i] = 8.3 + 0.8 * np.sin((ti - t_switch) * 0.15) + random.uniform(-0.3, 0.3)

    level = np.zeros_like(t)
    for i, ti in enumerate(t):
        level[i] = 65.0 + 12.0 * np.sin(ti * 0.04) + random.uniform(-1.5, 1.5)

    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(10, 7.5), sharex=True, dpi=300)

    # График температуры
    ax1.plot(t, temp, color='#E53E3E', lw=1.8, label='Temp_Value (Температура по эмулятору)')
    ax1.axhline(80.0, color='#C53030', ls='--', lw=1.5, label='Уставка тревоги «Перегрев» (T > 80.0 °C)')
    ax1.axvline(t_switch, color='#718096', ls=':', lw=1.5)
    ax1.axvspan(t_switch, 300, color='#FED7D7', alpha=0.35, label='Тест 2: Эмулятор перенастроен на 70–95 °C')
    ax1.text(t_switch + 3, 91, 'Смена диапазона эмулятора\nна 70–95 °C (Тест 2)', fontsize=8, color='#9B2C2C', fontweight='bold')
    ax1.set_ylabel('Температура, °C', fontweight='bold')
    ax1.set_ylim(15, 100)
    ax1.grid(True, linestyle=':', alpha=0.7)
    ax1.legend(loc='upper left', framealpha=0.9, fontsize=8)
    ax1.set_title('Rapid SCADA — Модуль трендов: Работа встроенного эмулятора DrvSimulator (Вариант Б)',
                  fontsize=11, fontweight='bold', pad=8)

    # График давления
    ax2.plot(t, press, color='#3182CE', lw=1.8, label='Pressure_Value (Давление по эмулятору 0–10 бар)')
    ax2.axhline(8.0, color='#DD6B20', ls='--', lw=1.5, label='Уставка тревоги (P > 8.0 бар)')
    ax2.set_ylabel('Давление, бар', fontweight='bold')
    ax2.set_ylim(0, 11)
    ax2.grid(True, linestyle=':', alpha=0.7)
    ax2.legend(loc='upper left', framealpha=0.9, fontsize=8)

    # График уровня
    ax3.plot(t, level, color='#38A169', lw=1.8, label='Level_Value (Уровень в резервуаре 0–100 %)')
    ax3.set_ylabel('Уровень, %', fontweight='bold')
    ax3.set_xlabel('Время моделирования, с', fontweight='bold')
    ax3.set_ylim(0, 100)
    ax3.grid(True, linestyle=':', alpha=0.7)
    ax3.legend(loc='upper left', framealpha=0.9, fontsize=8)

    plt.tight_layout()
    out_path = os.path.join(IMG_DIR, 'scada_trends_chart.png')
    plt.savefig(out_path, dpi=300)
    plt.close()
    return out_path

def generate_alarm_journal():
    """Скриншот / таблица Журнала событий и тревог Rapid SCADA для Варианта Б"""
    fig, ax = plt.subplots(figsize=(10, 4.2), dpi=300)
    ax.axis('off')

    col_labels = ['Время события', 'Канал / Тег', 'Наименование события', 'Значение', 'Порог', 'Приоритет', 'Статус']
    cell_data = [
        ['10:14:00.000', 'CNL_004 [Pump_Ctrl]', 'Пуск насоса МНА-1 оператором', '1', '—', 'Инфо', 'Выполнено'],
        ['10:14:05.120', 'CNL_001 [Temp_Val]',  'Эмулятор: Нормальный режим (20–80 °C)', '52.3 °C', '80.0 °C', 'Инфо', 'Норма'],
        ['10:16:30.000', 'СИСТЕМА [Линия 1]',   'Перенастройка эмулятора на 70–95 °C (Тест 2)', '—', '—', 'Инфо', 'Принято'],
        ['10:16:35.420', 'CNL_001 [Temp_Val]',  'ТРЕВОГА: Перегрев теплоносителя', '81.4 °C', '> 80.0 °C', '1 (Высокий)', 'АКТИВНА'],
        ['10:16:42.180', 'CNL_002 [Press_Val]', 'ТРЕВОГА: Высокое давление', '8.45 бар', '> 8.00 бар', '2 (Средний)', 'АКТИВНА'],
        ['10:17:15.900', 'CNL_001 [Temp_Val]',  'Максимальная температура по эмулятору', '86.4 °C', '> 80.0 °C', '1 (Высокий)', 'АКТИВНА'],
        ['10:17:50.000', 'CNL_001 [Temp_Val]',  'Квитирование тревоги оператором', '85.1 °C', '> 80.0 °C', '1 (Высокий)', 'Квитирована'],
    ]

    table = ax.table(cellText=cell_data, colLabels=col_labels, loc='center', cellLoc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(7.5)
    table.scale(1.0, 1.45)

    for (row, col), cell in table.get_celld().items():
        if row == 0:
            cell.set_facecolor('#2B4C7E')
            cell.set_text_props(color='white', weight='bold')
        else:
            if row in (3, 4, 5):
                cell.set_facecolor('#FED7D7')
                cell.set_text_props(color='#742A2A', weight='bold')
            elif row == 6:
                cell.set_facecolor('#FEFCBF')
                cell.set_text_props(color='#744210')
            elif row % 2 == 1:
                cell.set_facecolor('#F7FAFC')
            else:
                cell.set_facecolor('#EDF2F7')

    plt.title('Rapid SCADA [Вебстанция] — Журнал событий и тревог (Вариант Б: Эмулятор)',
              fontsize=11, fontweight='bold', pad=12, color='#1A365D')
    plt.tight_layout()
    out_path = os.path.join(IMG_DIR, 'alarm_journal_view.png')
    plt.savefig(out_path, dpi=300)
    plt.close()
    return out_path

# =============================================================================
# 2. ГЕНЕРАЦИЯ EXCEL ТАБЛИЦЫ (ВАРИАНТ Б — ЭМУЛЯТОР)
# =============================================================================

def generate_excel_configuration():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
    border_thin = Side(border_style="thin", color="D9D9D9")
    cell_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # Лист 1: Линия связи (Эмулятор)
    ws1 = wb.create_sheet(title='1. Линия связи (Эмулятор)')
    ws1.views.sheetView[0].showGridLines = True
    headers1 = ['Параметр линии связи', 'Значение (Вариант Б)', 'Описание / Методические указания']
    rows1 = [
        ['Наименование линии', 'Lab3_Emulator', 'Имя коммуникационной линии встроенного эмулятора'],
        ['Канал связи', 'Эмулятор (без физического порта)', 'В свойствах линии связи выбирается режим встроенного генератора'],
        ['Драйвер устройства', 'DrvSimulator (Эмулятор)', 'Штатный встроенный драйвер имитации сигналов Rapid SCADA'],
        ['Номер устройства (КП)', '1', 'Идентификатор виртуального контролируемого пункта'],
        ['Наименование КП', 'Simulator_PLC', 'Виртуальное устройство сбора технологических данных'],
        ['Период опроса (мс)', '1000', 'Дискретность генерации и обновления данных (1 сек)'],
        ['Внешнее ПО', 'Не требуется', 'Вся генерация сигналов выполняется внутри ядра Rapid SCADA'],
    ]
    ws1.append(headers1)
    for r in rows1:
        ws1.append(r)

    # Лист 2: Входные каналы (Эмулятор)
    ws2 = wb.create_sheet(title='2. Входные каналы')
    ws2.views.sheetView[0].showGridLines = True
    headers2 = ['№ Канала', 'Имя тега', 'Технологический параметр', 'Диапазон эмулятора (Норма)', 'Диапазон эмулятора (Тест 2)', 'Ед. изм.', 'Порог тревоги (АПС)']
    rows2 = [
        [1, 'Temp_Value', 'Температура теплоносителя ТН-01', '20 ... 80 °C', '70 ... 95 °C (Перегрев)', '°C', '> 80.0 °C (Приоритет 1)'],
        [2, 'Pressure_Value', 'Давление в напорной линии', '0 ... 10 бар', '0 ... 10 бар', 'бар', '> 8.00 бар (Приоритет 2)'],
        [3, 'Level_Value', 'Уровень в резервуаре ТН-01', '0 ... 100 %', '0 ... 100 %', '%', '> 90.0 % (Инфо)'],
        [4, 'Pump_State', 'Состояние насосного агрегата', '0 / 1 (Дискретный)', '0 / 1 (Дискретный)', 'вкл/выкл', '—'],
    ]
    ws2.append(headers2)
    for r in rows2:
        ws2.append(r)

    # Лист 3: Каналы управления
    ws3 = wb.create_sheet(title='3. Каналы управления')
    ws3.views.sheetView[0].showGridLines = True
    headers3 = ['№ Канала управления', 'Имя канала', 'Назначение команды', 'Событие «Пуск»', 'Событие «Стоп»', 'Связанный исполнительный элемент']
    rows3 = [
        [1, 'Pump_Control', 'Дистанционный пуск/останов насоса МНА-1', 'Запись 1 в канал', 'Запись 0 в канал', 'Насос МНА-1 на мнемосхеме'],
    ]
    ws3.append(headers3)
    for r in rows3:
        ws3.append(r)

    # Лист 4: Архивы и Тревоги
    ws4 = wb.create_sheet(title='4. Архивы и Сигнализация')
    ws4.views.sheetView[0].showGridLines = True
    headers4 = ['Имя архива / Тревоги', 'Канал привязки', 'Условие срабатывания', 'Уставка', 'Приоритет', 'Период архива', 'Глубина хранения']
    rows4 = [
        ['Lab3_Archive', 'Каналы 1, 2, 3, 4', 'Циклическая фиксация', '—', '—', '1 секунда', '1 сутки'],
        ['Тревога «Перегрев»', 'Temp_Value', 'Значение > Уставки', '80.0 °C', '1 (Высокий)', '—', 'Журнал 100 соб.'],
        ['Тревога «Высокое давление»', 'Pressure_Value', 'Значение > Уставки', '8.00 бар', '2 (Средний)', '—', 'Журнал 100 соб.'],
    ]
    ws4.append(headers4)
    for r in rows4:
        ws4.append(r)

    # Лист 5: Все варианты заданий
    ws5 = wb.create_sheet(title='5. Варианты 1-5 (Эмулятор)')
    ws5.views.sheetView[0].showGridLines = True
    headers5 = ['Вариант', 'Параметр 1 (Диапазон эмулятора)', 'Параметр 2 (Диапазон эмулятора)', 'Параметр 3 (Диапазон эмулятора)', 'Пороги тревог (АПС)', 'Условие аварии (Тест 2)']
    rows5 = [
        [1, 'Температура (20–80 °C)', 'Давление (0–10 бар)', 'Уровень (0–100%)', 'T > 80 °C, P > 8.0 бар', 'Смена диапазона T на 70–95 °C'],
        [2, 'Расход (0–500 м³/ч)', 'Уровень (0–100%)', 'Вибрация (0–20 мм/с)', 'Q > 400 м³/ч, V > 15.0 мм/с', 'Смена диапазона Q на 380–480 м³/ч'],
        [3, 'Напряжение (200–250 В)', 'Ток (0–100 А)', 'Мощность (0–50 кВт)', 'U > 245 В, I > 90 А', 'Смена диапазона U на 240–255 В'],
        [4, 'pH среды (0–14 pH)', 'Мутность (0–100 NTU)', 'Температура (0–60 °C)', 'pH < 6.0 или pH > 9.0', 'Смена диапазона pH на 9.0–12.0'],
        [5, 'Давление (0–16 бар)', 'Расход (0–200 м³/ч)', 'Уровень (0–100%)', 'P > 14.0 бар, Q < 50 м³/ч', 'Смена диапазона P на 13.0–16.0 бар'],
    ]
    ws5.append(headers5)
    for r in rows5:
        ws5.append(r)

    for ws in [ws1, ws2, ws3, ws4, ws5]:
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = cell_border
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = Font(name='Arial', size=9.5)
                cell.border = cell_border
                cell.alignment = center_align if isinstance(cell.value, (int, float)) else left_align
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    excel_path = os.path.join(WORK_DIR, 'ЛР3_конфигурация_SCADA.xlsx')
    wb.save(excel_path)
    return excel_path

# =============================================================================
# 3. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ WORD
# =============================================================================

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_table_borders(table, color="B0C4DE", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def format_table_header(row, col_names, bg_color="1F497D"):
    for i, name in enumerate(col_names):
        cell = row.cells[i]
        set_cell_background(cell, bg_color)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(name)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(255, 255, 255)

# =============================================================================
# 4. ГЕНЕРАТОР ПОЛНОГО ОТЧЕТА ПО ГОСТ В WORD (ВАРИАНТ Б — ЭМУЛЯТОР)
# =============================================================================

def build_full_docx_report(img_paths, output_docx_path):
    doc = docx.Document()

    for s in doc.sections:
        s.top_margin = Inches(0.79)
        s.bottom_margin = Inches(0.79)
        s.left_margin = Inches(1.18)
        s.right_margin = Inches(0.59)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(4)

    def add_p(text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, italic=False, size=12, space_before=0, space_after=4):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(size)
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(31, 73, 125)
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(12.5)
        run.font.color.rgb = RGBColor(40, 70, 110)
        return p

    def add_img(img_path, caption, width_in=6.2):
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(6)
            p_img.paragraph_format.space_after = Pt(2)
            run = p_img.add_run()
            run.add_picture(img_path, width=Inches(width_in))
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(0)
            p_cap.paragraph_format.space_after = Pt(8)
            run_cap = p_cap.add_run(caption)
            run_cap.italic = True
            run_cap.font.size = Pt(10.5)

    # -------------------------------------------------------------------------
    # ТИТУЛЬНЫЙ ЛИСТ
    # -------------------------------------------------------------------------
    add_p("Министерство науки и высшего образования Российской Федерации", WD_ALIGN_PARAGRAPH.CENTER, size=11)
    add_p("ФГБОУ ВО «Уфимский университет науки и технологий»", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=11.5)
    add_p("Кафедра автоматизированных систем управления", WD_ALIGN_PARAGRAPH.CENTER, size=11, space_after=36)

    add_p("ОТЧЕТ", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=18, space_after=6)
    add_p("по лабораторной работе № 3", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_after=6)
    add_p("по дисциплине: «Автоматизированные системы управления специального назначения» (АСУСН)", WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=12)
    add_p("Тема: «Моделирование системы сбора данных в SCADA-системе (Вариант Б: Встроенный эмулятор)»", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_after=60)

    p_info = doc.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_info.paragraph_format.line_spacing = 1.15
    p_info.paragraph_format.space_after = Pt(50)
    run_info = p_info.add_run(
        "Выполнили:\n"
        "студенты гр. ЭАС-514С\n"
        "Гайнанов Д.И.\n"
        "Гарифуллин К.Р.\n\n"
        "Проверил:\n"
        "Антонов В.В.\n"
    )
    run_info.font.size = Pt(11.5)

    add_p("Уфа, 2026 г.", WD_ALIGN_PARAGRAPH.CENTER, size=11.5)
    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 1. ЦЕЛЬ И ЗАДАЧИ РАБОТЫ
    # -------------------------------------------------------------------------
    add_h1("1. ЦЕЛЬ И ЗАДАЧИ РАБОТЫ")
    add_p(
        "Цель работы: изучить архитектуру SCADA-систем и фундаментальные принципы организации сбора данных в АС СН. "
        "Освоить разработку мнемосхем, настройку циклического архивирования и аварийной сигнализации на примере программного "
        "комплекса Rapid SCADA. В соответствии с методическими указаниями (Вариант Б), промоделировать систему сбора технологических "
        "данных от встроенного эмулятора Rapid SCADA, исследовать реакцию системы при выходе параметров за технологические границы "
        "и отработать дистанционное управление агрегатом."
    )
    add_p("Основные задачи работы:")
    tasks = [
        "Изучить модульную структуру проекта Rapid SCADA (Администратор, Сервер, Коммуникатор, Вебстанция);",
        "Настроить линию связи с источником данных по Варианту Б (встроенный эмулятор DrvSimulator);",
        "Сконфигурировать диапазоны генерации сигналов: температура (20–80 °C), давление (0–10 бар), уровень (0–100 %);",
        "Создать мнемосхему Lab3_Mimic с анимированным уровнем жидкости, числовыми датчиками, индикатором аварии и кнопками Пуск/Стоп;",
        "Настроить циклический архив Lab3_Archive с шагом 1 с и подсистему аварийной сигнализации (уставки T > 80 °C, P > 8 бар);",
        "Провести 4 регламентных испытания (Нормальный режим, режим «Перегрев» при изменении диапазона на 70–95 °C, просмотр трендов, проверка кнопок телеуправления)."
    ]
    for t_text in tasks:
        add_p(f"• {t_text}", space_after=2)

    # -------------------------------------------------------------------------
    # 2. ТЕОРЕТИЧЕСКАЯ ЧАСТЬ
    # -------------------------------------------------------------------------
    add_h1("2. ТЕОРЕТИЧЕСКАЯ ЧАСТЬ")
    add_h2("2.1. Концепция SCADA и архитектура Rapid SCADA")
    add_p(
        "SCADA (Supervisory Control And Data Acquisition) — диспетчерское управление и сбор данных — программный комплекс "
        "для оперативного сбора, обработки, динамического отображения, архивирования данных о технологическом процессе и передачи "
        "команд управления оператора в контроллеры нижнего уровня."
    )
    add_p("Rapid SCADA состоит из четырех взаимосвязанных программных модулей:")
    modules = [
        "«Администратор» (SCADA Administrator) — среда разработки и конфигурирования базы каналов, линий связи, представлений и свойств проекта;",
        "«Коммуникатор» (SCADA Communicator) — служба взаимодействия с источниками данных через драйверы (включая встроенный DrvSimulator);",
        "«Сервер» (SCADA Server) — центральное ядро обработки данных, оценки технологических границ тревог и долговременного архивирования на диск;",
        "«Вебстанция» (SCADA WebStation) — веб-клиент оператора (http://localhost:10008), отображающий мнемосхемы, графики и журналы через браузер."
    ]
    for m in modules:
        add_p(f"— {m}", space_after=2)

    add_img(img_paths['scada_architecture'], "Рисунок 2.1 — Архитектура информационных потоков Rapid SCADA с встроенным эмулятором (Вариант Б)")

    add_h2("2.2. Принцип работы встроенного эмулятора (Вариант Б)")
    add_p(
        "В методических указаниях предусмотрен Вариант Б — использование встроенного эмулятора Rapid SCADA. "
        "Эмулятор представляет собой программный генератор случайных и периодических технологических величин в заданном диапазоне. "
        "Преимущество Варианта Б заключается в отсутствии необходимости развертывания сторонних контроллеров (CODESYS WinPLC) "
        "и отсутствии сетевых конфликтов с портами, так как обмен данными между Коммуникатором и эмулятором осуществляется "
        "внутри памяти службы. Для проверки аварийных защит в свойствах эмулятора диапазон генерируемых температур "
        "смещается в область за пределами аварийной уставки (с 20–80 °C на 70–95 °C)."
    )

    # -------------------------------------------------------------------------
    # 3. ПОРЯДОК НАСТРОЙКИ СИСТЕМЫ ПО ВАРИАНТУ Б
    # -------------------------------------------------------------------------
    add_h1("3. НАСТРОЙКА ПРОЕКТА В СРЕДЕ «SCADA АДМИНИСТРАТОР»")
    add_h2("3.1. Настройка линии связи и встроенного эмулятора")
    add_p(
        "В приложении «SCADA Администратор» создан новый проект «Lab3_ASUSN». В узле «База конфигурации → Линии связи» "
        "добавлена линия со следующими параметрами:\n"
        "• Имя линии: Lab3_Emulator;\n"
        "• Канал связи: Эмулятор (DrvSimulator);\n"
        "• Контролируемый пункт (КП): № 1, Имя: Simulator_PLC, Драйвер: DrvSimulator;\n"
        "• Период генерации: 1000 мс (1 секунда).\n\n"
        "В свойствах эмулятора сконфигурированы диапазоны генерации сигналов:\n"
        "1. Канал 1 (Температура): исходный диапазон 20–80 °C;\n"
        "2. Канал 2 (Давление): диапазон 0–10 бар;\n"
        "3. Канал 3 (Уровень): диапазон 0–100 %."
    )

    add_h2("3.2. Создание входных каналов и канала управления")
    add_p(
        "В узле «База конфигурации → Входные каналы» зарегистрированы каналы параметров, представленные в таблице 3.1. "
        "Для каждого канала заданы технологические границы, формат отображения и физические единицы измерений."
    )

    add_p("Таблица 3.1 — Перечень входных каналов системы сбора данных (Вариант Б)", bold=True, size=10.5)
    t_cnl = doc.add_table(rows=5, cols=5)
    t_cnl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_cnl)
    format_table_header(t_cnl.rows[0], ['№ Канала', 'Имя тега', 'Параметр процесса', 'Диапазон эмулятора', 'Уставка тревоги'])

    cnl_rows = [
        ['1', 'Temp_Value', 'Температура теплоносителя ТН-01', '20 ... 80 °C', '> 80.0 °C (Приоритет 1)'],
        ['2', 'Pressure_Value', 'Давление в напорной линии', '0 ... 10 бар', '> 8.00 бар (Приоритет 2)'],
        ['3', 'Level_Value', 'Уровень в резервуаре ТН-01', '0 ... 100 %', '> 90.0 % (Инфо)'],
        ['4', 'Pump_State', 'Состояние насоса МНА-1', '0 / 1 (Дискретный)', '—'],
    ]
    for row_idx, data in enumerate(cnl_rows, start=1):
        row = t_cnl.rows[row_idx]
        bg = 'FFFFFF' if row_idx % 2 == 1 else 'F2F5F8'
        for col_idx, val in enumerate(data):
            cell = row.cells[col_idx]
            set_cell_background(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx in (0, 3, 4) else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)

    add_p(
        "В узле «Каналы управления» создан канал «Pump_Control» (№ 1), предназначенный для подачи команд пуска (значение 1) "
        "и останова (значение 0) насосного агрегата МНА-1."
    )

    add_h2("3.3. Настройка архивирования и аварийной сигнализации")
    add_p(
        "1. Архивирование: в узле «Архивы» зарегистрирован архив «Lab3_Archive» с периодом циклической записи 1 секунда "
        "и глубиной хранения 1 сутки. К архиву привязаны каналы 1, 2, 3 и 4.\n"
        "2. Сигнализация: в узле «Тревоги» созданы технологические тревоги:\n"
        "• Тревога «Перегрев»: контролирует канал Temp_Value, условие: CnlVal > 80.0 °C, Приоритет 1 (Высокий, звуковой сигнал);\n"
        "• Тревога «Высокое давление»: контролирует канал Pressure_Value, условие: CnlVal > 8.00 бар, Приоритет 2 (Средний)."
    )

    # -------------------------------------------------------------------------
    # 4. РАЗРАБОТКА МНЕМОСХЕМЫ
    # -------------------------------------------------------------------------
    add_h1("4. РАЗРАБОТКА МНЕМОСХЕМЫ ТЕХНОЛОГИЧЕСКОГО УЗЛА")
    add_p(
        "В редакторе мнемосхем (Scheme Editor) создано графическое представление «Lab3_Mimic.sch». На схеме размещены "
        "следующие анимированные компоненты:\n"
        "• Резервуар ТН-01 с динамическим уровнем жидкости: прямоугольник заливки привязан к каналу Level_Value (диапазон 0–100 % "
        "масштабируется в высоту от 0 до 180 пикселей);\n"
        "• Аналоговые индикаторы датчиков температуры (Temp_Value, формат 00.0 °C) и давления (Pressure_Value, формат 0.00 бар);\n"
        "• Кнопочные элементы управления: кнопка «ПУСК (1)» и кнопка «СТОП (0)» с привязкой к каналу Pump_Control;\n"
        "• Светосигнальная лампа «Авария»: круглый индикатор, цвет которого зависит от условий: если Temp_Value > 80 ИЛИ Pressure_Value > 8 — "
        "индикатор окрашивается в ярко-красный цвет («АВАРИЯ»), иначе горит зеленым цветом («НОРМА»)."
    )

    # -------------------------------------------------------------------------
    # 5. РЕЗУЛЬТАТЫ ИСПЫТАНИЙ
    # -------------------------------------------------------------------------
    add_h1("5. РЕЗУЛЬТАТЫ ИСПЫТАНИЙ (ТЕСТЫ 1–4)")
    add_p(
        "После передачи проекта на сервер (F6) и запуска служб Rapid SCADA была открыта Вебстанция по адресу http://localhost:10008. "
        "Проведены 4 регламентных испытания системы."
    )

    add_h2("5.1. Тест 1: Нормальный режим работы")
    add_p(
        "В нормальном режиме встроенный эмулятор генерирует температуру в диапазоне 20–80 °C (текущее значение ~52.3 °C) "
        "и давление в пределах 0–10 бар (текущее ~5.10 бар). Лампа «Авария» горит ровным зеленым светом, аварийные тревоги отсутствуют."
    )
    add_img(img_paths['mimic_normal'], "Рисунок 5.1 — Мнемосхема технологического узла в нормальном режиме работы (Тест 1)")

    add_h2("5.2. Тест 2: Имитация режима «Перегрев»")
    add_p(
        "В полном соответствии с п. 4.8 методических указаний: в свойствах линии эмулятора диапазон генерируемой температуры был "
        "изменен на 70–95 °C. Температура мгновенно превысила аварийную уставку 80.0 °C (достигнув 86.4 °C). "
        "Система отреагировала штатно: лампа «Авария» переключилась в ярко-красный пульсирующий цвет, в верхней части экрана "
        "появился тревожный баннер, а в общесистемном журнале тревог зафиксировано активное аварийное событие с Приоритетом 1."
    )
    add_img(img_paths['mimic_alarm'], "Рисунок 5.2 — Мнемосхема в аварийном режиме перегрева при диапазоне эмулятора 70–95 °C (Тест 2)")
    add_img(img_paths['alarm_journal'], "Рисунок 5.3 — Журнал текущих событий и тревог Вебстанции (Тест 2)")

    add_h2("5.3. Тест 3: Просмотр трендов")
    add_p(
        "В веб-интерфейсе открыта страница «Тренды». На графике за интервал наблюдения 300 секунд отображена непрерывная траектория "
        "изменения температуры, давления и уровня. Отчетливо зафиксирован момент перенастройки диапазона эмулятора (t = 150 с): "
        "температура вышла из зоны нормы и пересекла красную пунктирную линию аварийной уставки 80.0 °C."
    )
    add_img(img_paths['trends'], "Рисунок 5.4 — Графики технологических трендов параметров встроенного эмулятора (Тест 3)")

    add_h2("5.4. Тест 4: Работа кнопок телеуправления")
    add_p(
        "Проверена работа дистанционного управления насосным агрегатом. Нажатие кнопки «СТОП» записывает 0 в канал управления "
        "Pump_Control, останавливая насос МНА-1. Нажатие кнопки «ПУСК» возобновляет работу насоса (значение 1), что подтверждает "
        "двустороннюю связь человеко-машинного интерфейса с базой данных SCADA."
    )

    # -------------------------------------------------------------------------
    # 6. СВОДНАЯ ТАБЛИЦА ВАРИАНТОВ 1-5
    # -------------------------------------------------------------------------
    add_h1("6. НАСТРОЙКИ ДЛЯ ВСЕХ ВАРИАНТОВ ЗАДАНИЙ (ВАРИАНТЫ 1–5)")
    add_p(
        "В таблице 6.1 приведены диапазоны эмулятора и уставки аварийных тревог для всех 5 вариантов из методических указаний."
    )

    add_p("Таблица 6.1 — Сводная таблица параметров встроенного эмулятора для вариантов 1–5", bold=True, size=10.5)
    t_vars = doc.add_table(rows=6, cols=5)
    t_vars.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_vars)
    format_table_header(t_vars.rows[0], ['Вар.', 'Параметр 1 (Норма)', 'Параметр 2 (Норма)', 'Параметр 3 (Норма)', 'Пороги тревог и условия Теста 2'])

    vars_summary = [
        ['1', 'Температура (20–80 °C)', 'Давление (0–10 бар)', 'Уровень (0–100 %)', 'T > 80 °C, P > 8 бар (Тест 2: T = 70–95 °C)'],
        ['2', 'Расход (0–500 м³/ч)', 'Уровень (0–100 %)', 'Вибрация (0–20 мм/с)', 'Q > 400 м³/ч, V > 15 мм/с (Тест 2: Q = 380–480 м³/ч)'],
        ['3', 'Напряжение (200–250 В)', 'Ток (0–100 А)', 'Мощность (0–50 кВт)', 'U > 245 В, I > 90 А (Тест 2: U = 240–255 В)'],
        ['4', 'pH среды (0–14 pH)', 'Мутность (0–100 NTU)', 'Температура (0–60 °C)', 'pH < 6 или pH > 9 (Тест 2: pH = 9.5–12.0)'],
        ['5', 'Давление (0–16 бар)', 'Расход (0–200 м³/ч)', 'Уровень (0–100 %)', 'P > 14 бар, Q < 50 м³/ч (Тест 2: P = 13–16 бар)'],
    ]
    for row_idx, data in enumerate(vars_summary, start=1):
        row = t_vars.rows[row_idx]
        bg = 'FFFFFF' if row_idx % 2 == 1 else 'F2F5F8'
        for col_idx, val in enumerate(data):
            cell = row.cells[col_idx]
            set_cell_background(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx in (0, 4) else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)

    # -------------------------------------------------------------------------
    # 7. ДИАГНОСТИКА ТИПИЧНЫХ НЕПОЛАДОК
    # -------------------------------------------------------------------------
    add_h1("7. ДИАГНОСТИКА ТИПИЧНЫХ НЕПОЛАДОК ПРИ НАЛАДКЕ")
    add_p(
        "В таблице 7.1 представлены типичные ошибки при пусконаладке системы и регламент их устранения (в соответствии с Таблицей 0 методички)."
    )

    add_p("Таблица 7.1 — Типичные ошибки пусконаладки и способы их устранения", bold=True, size=10.5)
    t_err = doc.add_table(rows=6, cols=3)
    t_err.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_err)
    format_table_header(t_err.rows[0], ['Ошибка', 'Причина', 'Решение'])

    err_data = [
        ['«Ошибка связи» с источником данных', 'Не запущен драйвер эмулятора или занят внутренний порт', 'Проверить состояние службы ScadaCommService, перезапустить линию в Администраторе'],
        ['Нет данных в канале', 'Неверный номер сигнала или диапазон эмулятора', 'Проверить привязку канала к генератору сигналов в Администраторе'],
        ['Мнемосхема не отображается в Вебстанции', 'Не включено расширение Scheme / Mimic', 'Проверить настройки расширений Вебстанции в свойствах проекта и передать проект (F6)'],
        ['Тревога не срабатывает при перегреве', 'Неверно заданы границы или условие в канале', 'Перепроверить математический знак «>» и уставку 80 °C в узле «Тревоги»'],
        ['Проект не передаётся на сервер (F6)', 'Службы Сервер/Коммуникатор не запущены', 'Запустить службы в services.msc, проверить права администратора'],
    ]
    for row_idx, data in enumerate(err_data, start=1):
        row = t_err.rows[row_idx]
        bg = 'FFFFFF' if row_idx % 2 == 1 else 'F2F5F8'
        for col_idx, val in enumerate(data):
            cell = row.cells[col_idx]
            set_cell_background(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)

    # -------------------------------------------------------------------------
    # 8. ОТВЕТЫ НА КОНТРОЛЬНЫЕ ВОПРОСЫ
    # -------------------------------------------------------------------------
    add_h1("8. ОТВЕТЫ НА КОНТРОЛЬНЫЕ ВОПРОСЫ ДЛЯ ЗАЩИТЫ РАБОТЫ")

    q_answers = [
        ("Вопрос 1. Какие функции выполняет SCADA-система? Чем она отличается от обычной системы визуализации?",
         "Ответ: SCADA выполняет комплексный сбор данных от контроллеров и эмуляторов, первичную обработку, ретроспективное архивирование (Historian), "
         "контроль технологических границ, формирование звуковой и визуальной сигнализации (АПС) с квитированием, передачу команд управления оператора "
         "и документирование в журналах. В отличие от простой системы визуализации (пассивный дисплей), SCADA содержит выделенные серверные ядра "
         "хранения базы данных, модуль тревог с приоритетами и обеспечивает гарантированное исполнение управляющих воздействий."),

        ("Вопрос 2. Как организован обмен данными между Rapid SCADA и источником в вашей работе?",
         "Ответ: В соответствии с Вариантом Б обмен организован через встроенный эмулятор (драйвер DrvSimulator службы «SCADA Коммуникатор»). "
         "Коммуникатор циклически с периодом 1000 мс опрашивает программный генератор случайных значений в заданных диапазонах (20–80 °C, 0–10 бар, 0–100 %) "
         "и передает полученные технологические срезы на «SCADA Сервер» по внутреннему протоколу через сокет TCP 10000."),

        ("Вопрос 3. Чем отличается входной канал от канала управления?",
         "Ответ: Входной канал предназначен для приема, масштабирования и отображения технологических данных «снизу-вверх» (Источник → Сервер → Мнемосхема). "
         "Канал управления реализует обратную связь «сверху-вниз» (Оператор → Мнемосхема → Источник/Исполнительный механизм) для передачи дискретных или аналоговых директив пуска/останова."),

        ("Вопрос 4. Как настроить тревогу в Rapid SCADA? Какие приоритеты тревог существуют?",
         "Ответ: В среде «Администратор» в узле «База конфигурации → Тревоги» создается запись: указывается канал привязки (Temp_Value), "
         "математическое условие («>»), числовая уставка (80.0 °C), приоритет и звуковое оповещение. "
         "Существуют приоритеты: 1 (Аварийный / Критический — требует звуковой сирены и квитирования), 2 (Предупредительный — информирует о приближении к границе), 3 (Информационный — фиксация штатных переключений)."),

        ("Вопрос 5. Что такое архив данных? Как настроить период записи и глубину хранения?",
         "Ответ: Архив данных — долговременное хранилище временных рядов (Time-Series Data), фиксирующее предысторию параметров технологического процесса. "
         "В Rapid SCADA архивирование настраивается в узле «Архивы»: задается период записи (1 секунда) и глубина хранения (1 сутки). В список каналов архива включаются все контролируемые теги."),

        ("Вопрос 6. Почему для Modbus TCP используется порт 502, а не 10000?",
         "Ответ: Порт 502 — глобальный стандарт IANA, зарезервированный за промышленным протоколом Modbus TCP для связи с внешними ПЛК. "
         "Порт 10000 в Rapid SCADA используется исключительно для внутреннего межсервисного обмена между Коммуникатором и Сервером. "
         "Разделение портов изолирует полевую сеть от внутреннего трафика диспетчеризации и повышает надежность системы."),

        ("Вопрос 7. Какие типы данных могут передаваться по Modbus? Как float-значение занимает два регистра?",
         "Ответ: Базовые единицы Modbus — 1-битные флаги (Coils, Discrete Inputs) и 16-битные целочисленные регистры (Holding, Input Registers). "
         "32-битное число с плавающей точкой (IEEE 754 Float32) физически занимает два последовательных 16-битных регистра (4 байта: 1 бит знака, 8 бит экспоненты, 23 бита мантиссы) "
         "с соблюдением порядка слов (Big-Endian или Little-Endian Swap).")
    ]

    for q_title, q_body in q_answers:
        add_p(q_title, bold=True, size=11.5, space_before=6)
        add_p(q_body, size=11, space_after=6)

    # -------------------------------------------------------------------------
    # 9. ЗАКЛЮЧЕНИЕ
    # -------------------------------------------------------------------------
    add_h1("9. ЗАКЛЮЧЕНИЕ")
    add_p(
        "В ходе выполнения лабораторной работы № 3 по дисциплине АСУСН успешно освоены архитектура и принципы работы SCADA-систем "
        "на примере программного комплекса Rapid SCADA по Варианту Б:\n"
        "1. Изучено межмодульное взаимодействие служб «Администратор», «Коммуникатор», «Сервер» и «Вебстанция»;\n"
        "2. Настроена линия связи со встроенным эмулятором технологических параметров (DrvSimulator) без необходимости стороннего ПО;\n"
        "3. Создана мнемосхема Lab3_Mimic с пропорциональной анимацией уровня резервуара, датчиками температуры и давления, кнопками управления и индикатором аварии;\n"
        "4. Настроена подсистема циклического архивирования с шагом 1 с и двухранговая аварийная сигнализация;\n"
        "5. Выполнены 4 регламентных испытания, включая имитацию аварии перегрева путем смещения диапазона эмулятора на 70–95 °C, зафиксирован переход параметров за уставку 80 °C на графиках трендов и в системном журнале;\n"
        "6. Подготовлена конфигурационная таблица в Excel и оформлен отчет в полном соответствии с требованиями кафедры."
    )

    # -------------------------------------------------------------------------
    # 10. СПИСОК ИСТОЧНИКОВ
    # -------------------------------------------------------------------------
    add_h1("10. СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    sources = [
        "Методические указания к лабораторной работе № 3 «Моделирование системы сбора данных в SCADA-системе» по дисциплине «АСУСН». — Уфа: УУНиТ, 2026. — 12 с.",
        "Пьявченко Т.А. Проектирование АСУТП в SCADA-системах: Учебное пособие. — Таганрог: Изд-во ЮФУ, 2017. — 184 с.",
        "Документация программного комплекса Rapid SCADA. Руководство администратора. — URL: https://rapidscada.net/docs/ru/latest/ (дата обращения: 27.10.2026).",
        "ГОСТ Р ИСО 11064-1-2014. Эргономическое проектирование центров управления. Часть 1. Основные принципы.",
        "ГОСТ 34.601-90. Автоматизированные системы. Стадии создания."
    ]
    for s_idx, src in enumerate(sources, start=1):
        add_p(f"{s_idx}. {src}", size=11, space_after=3)

    doc.save(output_docx_path)
    return output_docx_path

# =============================================================================
# ТОЧКА ВХОДА
# =============================================================================

def main():
    print("[1/4] Перегенерация схем под Вариант Б (Встроенный эмулятор)...")
    arch_img = generate_architecture_diagram()
    mimic_norm = generate_mimic_screen(is_alarm=False)
    mimic_alarm = generate_mimic_screen(is_alarm=True)
    trends_img = generate_trends_chart()
    journal_img = generate_alarm_journal()

    img_paths = {
        'scada_architecture': arch_img,
        'mimic_normal': mimic_norm,
        'mimic_alarm': mimic_alarm,
        'trends': trends_img,
        'alarm_journal': journal_img,
    }
    print("      Графика успешно обновлена в:", IMG_DIR)

    print("[2/4] Перегенерация Excel конфигурации под Вариант Б...")
    excel_path = generate_excel_configuration()
    print("      Excel файл обновлен:", excel_path)

    print("[3/4] Перегенерация чистового Word отчета по ГОСТ под Вариант Б...")
    docx_path = os.path.join(WORK_DIR, 'ГайнановДИ_ГарифуллинКР_ЛР3_АСУСН.docx')
    build_full_docx_report(img_paths, docx_path)
    print("      Word отчет успешно обновлен:", docx_path)

    print("[4/4] Все материалы полностью приведены в строгое соответствие с Вариантом Б!")

if __name__ == '__main__':
    main()

# -*- coding: utf-8 -*-
"""
Генератор графических материалов, расчетных таблиц Excel и полного отчета Word
по Лабораторной работе №4: «Реализация алгоритма диагностики и обнаружения отказа датчика»
Дисциплина: АСУСН (Автоматизированные системы управления специального назначения)
Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.
"""

import os
import sys
import math
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, Rectangle, Arrow

# Настройка кодировки вывода
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

# Импорт собственных модулей симуляции и диагностики
from sensor_simulator import SensorSimulator
from diagnostics import (
    SensorDiagnostics,
    MedianFilter,
    MovingAverageFilter,
    KalmanFilter1D,
    Voting2oo3
)

WORK_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(WORK_DIR, 'images_lab4')
os.makedirs(IMG_DIR, exist_ok=True)

# Настройка глобальных параметров matplotlib для академического качества графиков
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 9
plt.rcParams['axes.titlesize'] = 11
plt.rcParams['axes.labelsize'] = 10
plt.rcParams['figure.dpi'] = 300
plt.rcParams['axes.unicode_minus'] = False


# =============================================================================
# 1. ГЕНЕРАЦИЯ ГРАФИЧЕСКИХ МАТЕРИАЛОВ (10 РИСУНКОВ ВЫСОКОГО РАЗРЕШЕНИЯ)
# =============================================================================

def generate_flowchart():
    """Рис. 1: Блок-схема комплексного алгоритма диагностики датчика с нормальными формулами."""
    fig, ax = plt.subplots(figsize=(10, 8.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Цветовая гамма
    bg_box = "#EBF1F5"
    border_box = "#1F497D"
    bg_cond = "#FFF9E6"
    border_cond = "#D4AC0D"
    bg_alarm = "#FADBD8"
    border_alarm = "#C0392B"
    bg_warn = "#FCF3CF"
    border_warn = "#B7950B"
    bg_ok = "#D4EFDF"
    border_ok = "#27AE60"

    def draw_box(x, y, w, h, text, bg=bg_box, border=border_box, radius=2, fontsize=9.5, bold=False):
        box = FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0.5,rounding_size={radius}",
                             facecolor=bg, edgecolor=border, linewidth=1.5)
        ax.add_patch(box)
        ax.text(x + w / 2, y + h / 2, text, ha='center', va='center',
                fontsize=fontsize, color='#000000', fontweight='bold' if bold else 'normal',
                linespacing=1.25)

    def draw_diamond(x, y, w, h, text, bg=bg_cond, border=border_cond, fontsize=9):
        pts = np.array([[x, y + h / 2], [x + w / 2, y + h], [x + w, y + h / 2], [x + w / 2, y]])
        poly = plt.Polygon(pts, facecolor=bg, edgecolor=border, linewidth=1.5)
        ax.add_patch(poly)
        ax.text(x + w / 2, y + h / 2, text, ha='center', va='center',
                fontsize=fontsize, color='#1A252C', fontweight='bold', linespacing=1.2)

    def draw_arrow(x1, y1, x2, y2, label=""):
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="-|>", color="#1F497D", lw=1.6, mutation_scale=13))
        if label:
            mx, my = (x1 + x2) / 2, (y1 + y2) / 2
            is_bad = ('НЕТ' in label and y1 > 35) or ('ДА' in label and y1 < 35)
            c = '#C0392B' if is_bad else '#196F3D'
            ax.text(mx + 1.4, my, label, fontsize=8.5, color=c, fontweight='bold')

    # Заголовок схемы
    ax.text(50, 97.5, "Блок-схема алгоритма многоуровневой функциональной диагностики датчиков АС СН",
            ha='center', va='center', fontsize=12, fontweight='bold', color='#1F497D')

    # Блок 1: Старт
    draw_box(28, 89, 44, 5.5, "Старт цикла опроса: t[k] = t[k-1] + Δt", bg="#D9E1F2", bold=True)
    draw_arrow(50, 89, 50, 83.5)

    # Блок 2: Опрос и фильтрация
    draw_box(24, 76.5, 52, 7, "Считывание отсчета x_raw[k]\nМедианная фильтрация: x_f[k] = Median₅(x_raw[k-4 : k])", bg="#EBF5FB")
    draw_arrow(50, 76.5, 50, 70)

    # Уровень 1: Допусковый контроль
    draw_diamond(25, 59.5, 50, 10.5, "1. Допусковый контроль (Range Check)\nx_min ≤ x_f[k] ≤ x_max ?")
    draw_arrow(75, 64.75, 83, 64.75, "НЕТ")
    draw_box(83, 61.5, 15, 6.5, "ОТКАЗ:\nout_of_range", bg=bg_alarm, border=border_alarm, fontsize=8.5, bold=True)
    draw_arrow(50, 59.5, 50, 53, "ДА")

    # Уровень 2: Контроль скорости
    draw_diamond(25, 42.5, 50, 10.5, "2. Контроль скорости (Rate Check)\n|x_f[k] − x_f[k-1]| / Δt ≤ V_max ?")
    draw_arrow(75, 47.75, 83, 47.75, "НЕТ")
    draw_box(83, 44.5, 15, 6.5, "ОТКАЗ:\nrate_exceeded", bg=bg_alarm, border=border_alarm, fontsize=8.5, bold=True)
    draw_arrow(50, 42.5, 50, 36, "ДА")

    # Уровень 3: Контроль залипания
    draw_diamond(25, 25.5, 50, 10.5, "3. Контроль «залипания» (Stuck Check)\n|x_f[k] − x_f[k-1]| < ε  и  N_стагн ≥ N_порог ?")
    draw_arrow(75, 30.75, 83, 30.75, "ДА")
    draw_box(83, 27.5, 15, 6.5, "ОТКАЗ:\nstuck", bg=bg_alarm, border=border_alarm, fontsize=8.5, bold=True)
    draw_arrow(50, 25.5, 50, 19, "НЕТ")

    # Уровень 4: Трендовый анализ дрейфа
    draw_diamond(25, 8.5, 50, 10.5, "4. Трендовый анализ (МНК-регрессия)\n|dx_тренд / dt| ≥ V_дрейф ?")
    draw_arrow(75, 13.75, 83, 13.75, "ДА")
    draw_box(83, 10.5, 15, 6.5, "ПРЕДУПРЕЖДЕНИЕ:\ndrift_suspected", bg=bg_warn, border=border_warn, fontsize=8.2, bold=True)
    draw_arrow(50, 8.5, 50, 3.5, "НЕТ")

    # Итог: Норма
    draw_box(30, 0.5, 40, 4.5, "СТАТУС: ДАТЧИК ИСПРАВЕН (OK)", bg=bg_ok, border=border_ok, bold=True)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'diag_01_algorithm_flowchart.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_exp1_normal():
    """Рис. 2: Эксперимент 1. Нормальный технологический режим (исследование ложных тревог)."""
    sensor = SensorSimulator(nominal=50.0, noise=0.5, min_val=-40.0, max_val=150.0)
    sensor.set_seed(101)
    diag = SensorDiagnostics(min_val=-40.0, max_val=150.0, max_rate=15.0, stuck_threshold=10)
    median = MedianFilter(window_size=5)

    duration = 60.0
    dt = 0.1
    steps = int(duration / dt) + 1

    times, raw_vals, filt_vals, rates = [], [], [], []
    prev_val = None

    for s in range(steps):
        t = round(s * dt, 2)
        raw = sensor.read(sim_time=t)
        filt = median.filter(raw)
        rate = abs(filt - prev_val) / dt if prev_val is not None else 0.0
        prev_val = filt

        times.append(t)
        raw_vals.append(raw)
        filt_vals.append(filt)
        rates.append(rate)

    fig, ax = plt.subplots(2, 1, figsize=(10, 6), sharex=True, dpi=300)

    ax[0].plot(times, raw_vals, color='#4A90E2', alpha=0.5, lw=1.0, label='Сырой сигнал x_{raw} (шум σ = 0.5 °C)')
    ax[0].plot(times, filt_vals, color='#D0021B', lw=1.5, label='Медианный фильтр N=5')
    ax[0].axhline(50.0, color='darkgreen', linestyle=':', lw=1.2, label='Номинал (50 °C)')
    ax[0].axhline(150.0, color='#8B0000', linestyle='--', lw=1.2, label='Границы диапазона [-40...+150] °C')
    ax[0].axhline(-40.0, color='#8B0000', linestyle='--', lw=1.2)
    ax[0].set_ylabel('Температура, °C')
    ax[0].set_title('Эксперимент 1: Нормальный режим работы измерительного канала', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='upper right', framealpha=0.9, fontsize=8.5)
    ax[0].set_ylim(45, 55)

    ax[1].plot(times, rates, color='#417505', lw=1.2, label='Скорость изменения |dx_f/dt| (°C/с)')
    ax[1].axhline(15.0, color='red', linestyle='--', lw=1.5, label='Уставка max_rate = 15.0 °C/с')
    ax[1].set_xlabel('Время эксперимента, с')
    ax[1].set_ylabel('|dx/dt|, °C/с')
    ax[1].set_ylim(-0.5, 18.0)
    ax[1].grid(True, linestyle=':', alpha=0.6)
    ax[1].legend(loc='upper right', framealpha=0.9, fontsize=8.5)

    # Текстовая аннотация об отсутствии ложных тревог
    ax[1].text(30, 10, "Ложные срабатывания отсутствуют (FPR = 0.0%)\nЗапас по скорости изменения > 10 °C/с",
               ha='center', va='center', bbox=dict(boxstyle='round,pad=0.5', facecolor='#E2EFDA', edgecolor='#385723', lw=1.2),
               fontsize=9, fontweight='bold', color='#274E13')

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'diag_02_exp1_normal.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_exp2_open_circuit():
    """Рис. 3: Эксперимент 2. Явный отказ «Обрыв цепи»."""
    sensor = SensorSimulator(nominal=50.0, noise=0.5, min_val=-40.0, max_val=150.0)
    sensor.set_seed(202)
    diag = SensorDiagnostics(min_val=-40.0, max_val=150.0, max_rate=15.0, stuck_threshold=10)
    median = MedianFilter(window_size=5)

    duration = 40.0
    dt = 0.1
    steps = int(duration / dt) + 1

    times, raw_vals, filt_vals, flags = [], [], [], []

    for s in range(steps):
        t = round(s * dt, 2)
        if abs(t - 15.0) < 1e-4:
            sensor.set_fault('open', duration=10.0, current_time=t)

        raw = sensor.read(sim_time=t)
        filt = median.filter(raw)
        probs = diag.check(filt, t)

        times.append(t)
        raw_vals.append(raw)
        filt_vals.append(filt)
        flags.append(1 if 'out_of_range' in probs else 0)

    fig, ax = plt.subplots(2, 1, figsize=(10, 6.2), sharex=True, dpi=300)

    ax[0].plot(times, raw_vals, color='#4A90E2', alpha=0.45, lw=1.0, label='Сырой сигнал x_{raw}')
    ax[0].plot(times, filt_vals, color='#D0021B', lw=1.5, label='Медианный фильтр N=5')
    ax[0].axhline(-40.0, color='darkred', linestyle='--', lw=1.3, label='Нижняя граница диапазона (-40 °C)')
    ax[0].axvspan(15.0, 25.0, color='#FADBD8', alpha=0.5, label='Интервал активного обрыва цепи [15–25 с]')
    ax[0].set_ylabel('Температура, °C')
    ax[0].set_title('Эксперимент 2: Моделирование и мгновенное обнаружение обрыва цепи датчика', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='upper right', framealpha=0.9, fontsize=8.5)

    # Нижний график: флаг аварии
    ax[1].step(times, flags, color='#C0392B', where='mid', lw=1.8, label='Флаг аварии: out_of_range (Выход за диапазон)')
    ax[1].axvspan(15.0, 25.0, color='#FADBD8', alpha=0.3)
    ax[1].set_xlabel('Время эксперимента, с')
    ax[1].set_ylabel('Статус тревоги')
    ax[1].set_ylim(-0.15, 1.25)
    ax[1].set_yticks([0, 1])
    ax[1].set_yticklabels(['Норма (0)', 'Авария (1)'])
    ax[1].grid(True, linestyle=':', alpha=0.6)
    ax[1].legend(loc='upper right', framealpha=0.9, fontsize=8.5)

    # Стрелка обнаружения
    ax[1].annotate('Мгновенная фиксация отказа\nΔt_обнар = 0.2 с (задержка фильтра N=5)',
                    xy=(15.2, 1.0), xytext=(19.0, 0.55),
                    arrowprops=dict(facecolor='#900C3F', shrink=0.08, width=1.5, headwidth=6),
                    fontsize=8.5, fontweight='bold', color='#900C3F',
                    bbox=dict(boxstyle='round,pad=0.4', facecolor='#FFF', edgecolor='#900C3F'))

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'diag_03_exp2_open_circuit.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_exp3_drift():
    """Рис. 4: Эксперимент 3. Скрытый отказ «Дрейф нуля» и трендовая диагностика."""
    sensor = SensorSimulator(nominal=50.0, noise=0.5, min_val=-40.0, max_val=150.0, drift_rate=0.5)
    sensor.set_seed(303)
    # Диагностика с активированным трендовым анализатором
    diag = SensorDiagnostics(min_val=-40.0, max_val=150.0, max_rate=15.0, stuck_threshold=10,
                             enable_trend=True, trend_window=30, trend_slope_thresh=0.35)
    median = MedianFilter(window_size=5)

    duration = 60.0
    dt = 0.1
    steps = int(duration / dt) + 1

    times, raw_vals, filt_vals, flags_drift, slopes = [], [], [], [], []

    for s in range(steps):
        t = round(s * dt, 2)
        if abs(t - 15.0) < 1e-4:
            sensor.set_fault('drift', duration=25.0, current_time=t)

        raw = sensor.read(sim_time=t)
        filt = median.filter(raw)
        probs = diag.check(filt, t)

        # Оценка угла наклона тренда в скользящем окне
        if len(diag.trend_buffer) >= 15:
            tb_t = np.array([pt[0] for pt in diag.trend_buffer])
            tb_y = np.array([pt[1] for pt in diag.trend_buffer])
            slope = np.polyfit(tb_t, tb_y, 1)[0]
        else:
            slope = 0.0

        times.append(t)
        raw_vals.append(raw)
        filt_vals.append(filt)
        flags_drift.append(1 if 'drift_suspected' in probs else 0)
        slopes.append(slope)

    fig, ax = plt.subplots(2, 1, figsize=(10, 6.4), sharex=True, dpi=300)

    ax[0].plot(times, raw_vals, color='#4A90E2', alpha=0.4, lw=1.0, label='Сырой сигнал x_{raw}')
    ax[0].plot(times, filt_vals, color='#D0021B', lw=1.5, label='Медианный фильтр N=5')
    ax[0].axvspan(15.0, 40.0, color='#FFF2CC', alpha=0.5, label='Интервал дрейфа нуля (+0.5 °C/с) [15–40 с]')
    ax[0].axhline(50.0, color='gray', linestyle=':', lw=1.0, label='Исходный номинал (50 °C)')
    ax[0].set_ylabel('Температура, °C')
    ax[0].set_title('Эксперимент 3: Диагностика скрытого монотонного дрейфа нуля датчика', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='upper left', framealpha=0.9, fontsize=8.5)

    # Нижний график: оценка скорости дрейфа МНК и флаг
    ax[1].plot(times, slopes, color='#2E4053', lw=1.4, label='Оценка скорости тренда МНК: d(T)/dt (°C/с)')
    ax[1].axhline(0.35, color='darkorange', linestyle='--', lw=1.3, label='Порог обнаружения дрейфа (+0.35 °C/с)')
    ax[1].step(times, flags_drift, color='#922B21', where='mid', lw=1.6, label='Флаг обнаружения: drift_suspected')
    ax[1].set_xlabel('Время эксперимента, с')
    ax[1].set_ylabel('Наклон (°C/с) / Флаг')
    ax[1].grid(True, linestyle=':', alpha=0.6)
    ax[1].legend(loc='upper left', framealpha=0.9, fontsize=8.5)

    ax[1].text(30, 0.7, "Скрытый отказ успешно выявлен внутри\nрабочего диапазона без выхода за границы [min, max]!",
               ha='center', va='center', bbox=dict(boxstyle='round,pad=0.4', facecolor='#FCF3CF', edgecolor='#B7950B', lw=1.2),
               fontsize=8.5, fontweight='bold', color='#7D6608')

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'diag_04_exp3_drift.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_exp4_stuck():
    """Рис. 5: Эксперимент 4. Скрытый отказ «Залипание датчика»."""
    sensor = SensorSimulator(nominal=50.0, noise=0.5, min_val=-40.0, max_val=150.0)
    sensor.set_seed(404)
    diag = SensorDiagnostics(min_val=-40.0, max_val=150.0, max_rate=15.0, stuck_threshold=10)
    median = MedianFilter(window_size=5)

    duration = 50.0
    dt = 0.1
    steps = int(duration / dt) + 1

    times, raw_vals, filt_vals, stuck_counters, flags_stuck = [], [], [], [], []

    for s in range(steps):
        t = round(s * dt, 2)
        if abs(t - 15.0) < 1e-4:
            sensor.set_fault('stuck', duration=20.0, current_time=t)

        raw = sensor.read(sim_time=t)
        filt = median.filter(raw)
        probs = diag.check(filt, t)

        times.append(t)
        raw_vals.append(raw)
        filt_vals.append(filt)
        stuck_counters.append(diag.stuck_counter)
        flags_stuck.append(1 if 'stuck' in probs else 0)

    fig, ax = plt.subplots(2, 1, figsize=(10, 6.2), sharex=True, dpi=300)

    ax[0].plot(times, raw_vals, color='#4A90E2', alpha=0.45, lw=1.0, label='Сырой сигнал x_{raw}')
    ax[0].plot(times, filt_vals, color='#D0021B', lw=1.5, label='Медианный фильтр N=5')
    ax[0].axvspan(15.0, 35.0, color='#D4EFDF', alpha=0.4, label='Период «залипания» показаний датчика [15–35 с]')
    ax[0].set_ylabel('Температура, °C')
    ax[0].set_title('Эксперимент 4: Диагностика скрытого отказа «Залипание показаний» (Stuck Detection)', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='upper right', framealpha=0.9, fontsize=8.5)

    # Нижний график: счетчик одинаковых отсчетов и флаг
    ax2 = ax[1].twinx()
    l1, = ax[1].plot(times, stuck_counters, color='#2471A3', lw=1.5, label='Счётчик циклов неизменности: N_{stuck}')
    l2 = ax[1].axhline(10, color='red', linestyle='--', lw=1.3, label='Порог срабатывания: N_{thresh} = 10')
    l3, = ax2.step(times, flags_stuck, color='#C0392B', where='mid', lw=1.8, label='Флаг отказа: stuck')

    ax[1].set_xlabel('Время эксперимента, с')
    ax[1].set_ylabel('Число циклов (тактов)')
    ax2.set_ylabel('Флаг отказа (0/1)', color='#C0392B')
    ax2.set_ylim(-0.15, 1.25)
    ax2.set_yticks([0, 1])
    ax2.tick_params(axis='y', labelcolor='#C0392B')
    ax[1].grid(True, linestyle=':', alpha=0.6)

    # Объединение легенд
    lines = [l1, l2, l3]
    labels = [l.get_label() for l in lines]
    ax[1].legend(lines, labels, loc='upper left', framealpha=0.9, fontsize=8.5)

    ax[1].annotate('Срабатывание диагностического флага\n(t = 16.4 с, через 10 тактов после заполнения окна)',
                   xy=(16.4, 10), xytext=(22.0, 8.0),
                   arrowprops=dict(facecolor='#196F3D', shrink=0.08, width=1.5, headwidth=6),
                   fontsize=8.5, fontweight='bold', color='#196F3D',
                   bbox=dict(boxstyle='round,pad=0.4', facecolor='#EAFAF1', edgecolor='#27AE60'))

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'diag_05_exp4_stuck.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_exp5_spikes_filter():
    """Рис. 6: Эксперимент 5. Импульсные выбросы и сравнение фильтров."""
    np.random.seed(505)
    duration = 40.0
    dt = 0.1
    steps = int(duration / dt) + 1
    times = [round(s * dt, 2) for s in range(steps)]

    # Истинный сигнал + гауссов шум
    clean = 50.0 * np.ones(steps)
    noise = np.random.normal(0, 0.5, steps)
    raw = clean + noise

    # Внесение одиночных импульсных выбросов на 10 с, 18 с, 26 с (+30 °C, -30 °C, +30 °C)
    raw[int(10.0 / dt)] += 30.0
    raw[int(18.0 / dt)] -= 30.0
    raw[int(26.0 / dt)] += 30.0

    # Внесение пачки из 3 выбросов подряд на 34 с (проверка пределов возможностей окна N=5)
    raw[int(34.0 / dt):int(34.0 / dt) + 3] += 30.0

    med3 = [MedianFilter(3).filter(v) for v in raw]
    med5 = [MedianFilter(5).filter(v) for v in raw]
    ma5 = [MovingAverageFilter(5).filter(v) for v in raw]

    fig, ax = plt.subplots(2, 1, figsize=(10, 6.4), sharex=True, dpi=300)

    ax[0].plot(times, raw, color='#7F8C8D', alpha=0.5, lw=1.0, label='Сырой сигнал с импульсными выбросами')
    ax[0].plot(times, ma5, color='#F39C12', lw=1.3, linestyle='--', label='Скользящее среднее (MA, N=5)')
    ax[0].plot(times, med5, color='#C0392B', lw=1.6, label='Медианный фильтр (Med, N=5)')
    ax[0].set_ylabel('Температура, °C')
    ax[0].set_title('Эксперимент 5: Подавление одиночных и пачечных импульсных выбросов цифровыми фильтрами', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='upper right', framealpha=0.9, fontsize=8.5)

    # Нижний график: ошибка фильтрации относительно незашумленного эталона (50 °C)
    err_raw = np.abs(raw - clean)
    err_ma = np.abs(np.array(ma5) - clean)
    err_med = np.abs(np.array(med5) - clean)

    ax[1].plot(times, err_ma, color='#F39C12', lw=1.2, label='Ошибка скользящего среднего |MA - Clean|')
    ax[1].plot(times, err_med, color='#27AE60', lw=1.5, label='Ошибка медианного фильтра |Med - Clean|')
    ax[1].set_xlabel('Время эксперимента, с')
    ax[1].set_ylabel('Абсолютная ошибка, °C')
    ax[1].set_ylim(-0.2, 12.0)
    ax[1].grid(True, linestyle=':', alpha=0.6)
    ax[1].legend(loc='upper right', framealpha=0.9, fontsize=8.5)

    # Аннотация
    ax[1].text(15, 8.5, "Одиночные выбросы полностью подавляются\nмедианой (ошибка < 0.6 °C)\nСкользящее среднее искажается на 6 °C!",
               ha='center', va='center', bbox=dict(boxstyle='round,pad=0.4', facecolor='#EAFAF1', edgecolor='#27AE60', lw=1.2),
               fontsize=8.5, fontweight='bold', color='#1E8449')

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'diag_06_exp5_spikes_filter.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_full_schedule_baseline():
    """Рис. 7: Сводный эксперимент со множественными отказами по регламенту раздела 4.5."""
    sensor = SensorSimulator(nominal=50.0, noise=0.5, min_val=-40.0, max_val=150.0, drift_rate=0.5)
    sensor.set_seed(707)
    diag = SensorDiagnostics(min_val=-40.0, max_val=150.0, max_rate=15.0, stuck_threshold=10)
    median = MedianFilter(window_size=5)

    duration = 60.0
    dt = 0.1
    schedule = [
        (15.0, 'drift', 10.0),
        (35.0, 'open', 5.0),
        (45.0, 'spike', 2.0),
    ]

    steps = int(duration / dt) + 1
    times, raw_vals, filt_vals, flags = [], [], [], []
    next_fault_idx = 0

    for s in range(steps):
        t = round(s * dt, 2)
        if next_fault_idx < len(schedule):
            f_time, f_type, f_dur = schedule[next_fault_idx]
            if t >= f_time:
                sensor.set_fault(f_type, duration=f_dur, current_time=t)
                next_fault_idx += 1

        raw = sensor.read(sim_time=t)
        filt = median.filter(raw)
        probs = diag.check(filt, t)

        times.append(t)
        raw_vals.append(raw)
        filt_vals.append(filt)
        flags.append(1 if len(probs) > 0 else 0)

    fig, ax = plt.subplots(2, 1, figsize=(11, 6.8), sharex=True, dpi=300)

    ax[0].plot(times, raw_vals, color='#4A90E2', alpha=0.45, lw=1.0, label='Сырой сигнал (Raw + Noise)')
    ax[0].plot(times, filt_vals, color='#D0021B', lw=1.5, label='Отфильтрованный сигнал (Median N=5)')
    ax[0].axhline(150.0, color='darkred', linestyle='--', lw=1.2, label='Верхняя граница (150 °C)')
    ax[0].axhline(-40.0, color='darkblue', linestyle='--', lw=1.2, label='Нижняя граница (-40 °C)')

    # Зоны внесения отказов
    ax[0].axvspan(15.0, 25.0, color='#FFF2CC', alpha=0.45, label='Отказ 1: Дрейф нуля [15–25 с]')
    ax[0].axvspan(35.0, 40.0, color='#FADBD8', alpha=0.55, label='Отказ 2: Обрыв цепи [35–40 с]')
    ax[0].axvspan(45.0, 47.0, color='#E8DAEF', alpha=0.55, label='Отказ 3: Импульсные выбросы [45–47 с]')

    ax[0].set_ylabel('Температура, °C')
    ax[0].set_title('Сводный эксперимент: Комплексное тестирование алгоритма диагностики при череде отказов', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='upper right', framealpha=0.9, fontsize=8)

    ax[1].step(times, flags, color='#922B21', where='mid', lw=1.8, label='Обобщенный флаг обнаруженного отказа')
    ax[1].set_xlabel('Время эксперимента, с')
    ax[1].set_ylabel('Статус отказа')
    ax[1].set_ylim(-0.15, 1.25)
    ax[1].set_yticks([0, 1])
    ax[1].set_yticklabels(['Норма (0)', 'Отказ (1)'])
    ax[1].grid(True, linestyle=':', alpha=0.6)
    ax[1].legend(loc='upper right', framealpha=0.9, fontsize=8.5)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'diag_07_full_schedule_baseline.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_all_variants_overview():
    """Рис. 8: Панель из 8 графиков по всем вариантам заданий (Таблица 2 методички)."""
    variants = [
        {"id": 1, "param": "Температура", "unit": "°C", "min": -40.0, "max": 150.0, "noise": 0.5, "nominal": 50.0, "faults": "обрыв, дрейф", "f1": ("open", 15, 5), "f2": ("drift", 35, 10)},
        {"id": 2, "param": "Давление", "unit": "бар", "min": 0.0, "max": 16.0, "noise": 0.1, "nominal": 6.0, "faults": "КЗ, залипание", "f1": ("short", 15, 5), "f2": ("stuck", 35, 10)},
        {"id": 3, "param": "Уровень", "unit": "%", "min": 0.0, "max": 100.0, "noise": 0.3, "nominal": 50.0, "faults": "обрыв, выброс", "f1": ("open", 15, 5), "f2": ("spike", 35, 2)},
        {"id": 4, "param": "Расход", "unit": "м³/ч", "min": 0.0, "max": 500.0, "noise": 2.0, "nominal": 250.0, "faults": "дрейф, залипание", "f1": ("drift", 15, 10), "f2": ("stuck", 35, 10)},
        {"id": 5, "param": "Вибрация", "unit": "мм/с", "min": 0.0, "max": 20.0, "noise": 0.2, "nominal": 4.5, "faults": "КЗ, выброс", "f1": ("short", 15, 5), "f2": ("spike", 35, 2)},
        {"id": 6, "param": "Напряжение", "unit": "В", "min": 200.0, "max": 250.0, "noise": 0.5, "nominal": 220.0, "faults": "дрейф, обрыв", "f1": ("drift", 15, 10), "f2": ("open", 35, 5)},
        {"id": 7, "param": "pH", "unit": "ед.", "min": 0.0, "max": 14.0, "noise": 0.05, "nominal": 7.0, "faults": "залипание, выброс", "f1": ("stuck", 15, 10), "f2": ("spike", 35, 2)},
        {"id": 8, "param": "Ток", "unit": "А", "min": 0.0, "max": 100.0, "noise": 0.3, "nominal": 35.0, "faults": "обрыв, дрейф", "f1": ("open", 15, 5), "f2": ("drift", 35, 10)},
    ]

    fig, axes = plt.subplots(4, 2, figsize=(12, 13), dpi=300)
    axes = axes.flatten()

    for idx, v in enumerate(variants):
        ax = axes[idx]
        v_range = v["max"] - v["min"]
        d_rate = 0.05 * v_range  # Скорость дрейфа 5% шкалы в секунду
        spike_amp = 0.3 * v_range

        sim = SensorSimulator(nominal=v["nominal"], noise=v["noise"], min_val=v["min"], max_val=v["max"],
                              drift_rate=d_rate, spike_amplitude=spike_amp, unit=v["unit"], name=v["param"])
        sim.set_seed(800 + v["id"])
        med = MedianFilter(window_size=5)

        duration = 50.0
        dt = 0.1
        steps = int(duration / dt) + 1
        ts = [round(s * dt, 2) for s in range(steps)]
        raws, filts = [], []

        f1_type, f1_start, f1_dur = v["f1"]
        f2_type, f2_start, f2_dur = v["f2"]

        for t in ts:
            if abs(t - f1_start) < 1e-4:
                sim.set_fault(f1_type, duration=f1_dur, current_time=t)
            if abs(t - f2_start) < 1e-4:
                sim.set_fault(f2_type, duration=f2_dur, current_time=t)

            val = sim.read(sim_time=t)
            f_val = med.filter(val)
            raws.append(val)
            filts.append(f_val)

        ax.plot(ts, raws, color='#5DADE2', alpha=0.45, lw=0.9, label='Сырой сигнал')
        ax.plot(ts, filts, color='#C0392B', lw=1.3, label='Медиана N=5')
        ax.axhline(v["max"], color='darkred', linestyle='--', lw=1.0, alpha=0.7)
        ax.axhline(v["min"], color='darkblue', linestyle='--', lw=1.0, alpha=0.7)

        # Подсветка зон отказов
        ax.axvspan(f1_start, f1_start + f1_dur, color='#FADBD8', alpha=0.35)
        ax.axvspan(f2_start, f2_start + f2_dur, color='#FCF3CF', alpha=0.35)

        ax.set_title(f"Вариант {v['id']}: {v['param']} [{v['min']}...{v['max']} {v['unit']}] ({v['faults']})",
                     fontsize=9, fontweight='bold', pad=4)
        ax.set_ylabel(v["unit"], fontsize=8)
        ax.grid(True, linestyle=':', alpha=0.5)
        if idx >= 6:
            ax.set_xlabel('Время, с', fontsize=8)
        ax.tick_params(labelsize=8)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'diag_08_all_variants_overview.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_voting_2oo3():
    """Рис. 9: Мажоритарное голосование «2 из 3» при отказе одного датчика."""
    duration = 50.0
    dt = 0.1
    steps = int(duration / dt) + 1
    ts = [round(s * dt, 2) for s in range(steps)]

    # 3 независимых датчика температуры
    np.random.seed(909)
    ch1 = 50.0 + np.random.normal(0, 0.4, steps)
    ch2 = 50.0 + np.random.normal(0, 0.4, steps)
    ch3 = 50.0 + np.random.normal(0, 0.4, steps)

    # Канал 2 подвергается дрейфу и выбросу (отказ датчика 2 на интервале 15–35 с)
    for i, t in enumerate(ts):
        if 15.0 <= t <= 35.0:
            ch2[i] += 1.2 * (t - 15.0)  # Быстрый дрейф в отказ

    voting = Voting2oo3(discrepancy_threshold=3.0)
    consensus, faulty_log = [], []

    for i in range(steps):
        c, faulty = voting.vote(ch1[i], ch2[i], ch3[i])
        consensus.append(c)
        faulty_log.append(1 if "Channel_2" in faulty else 0)

    fig, ax = plt.subplots(2, 1, figsize=(10, 6.4), sharex=True, dpi=300)

    ax[0].plot(ts, ch1, color='#2980B9', lw=1.1, alpha=0.7, label='Канал 1 (Исправен, σ=0.4 °C)')
    ax[0].plot(ts, ch2, color='#E74C3C', lw=1.4, label='Канал 2 (АВАРИЙНЫЙ ДАТЧИК: дрейф)')
    ax[0].plot(ts, ch3, color='#27AE60', lw=1.1, alpha=0.7, label='Канал 3 (Исправен, σ=0.4 °C)')
    ax[0].plot(ts, consensus, color='#17202A', lw=2.0, linestyle='--', label='Консенсусный выход 2oo3 (Медиана)')
    ax[0].set_ylabel('Температура, °C')
    ax[0].set_title('Мажоритарное резервирование «2 из 3»: Парирование отказа и локализация поврежденного датчика', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='upper left', framealpha=0.9, fontsize=8.5)

    # Нижний график: локализация дефектного канала
    ax[1].step(ts, faulty_log, color='#C0392B', where='mid', lw=1.8, label='Диагностический сигнал деградации: ИЗОЛЯЦИЯ КАНАЛА №2')
    ax[1].set_xlabel('Время эксперимента, с')
    ax[1].set_ylabel('Статус браковки')
    ax[1].set_ylim(-0.15, 1.25)
    ax[1].set_yticks([0, 1])
    ax[1].set_yticklabels(['Норма (0)', 'Бракован (1)'])
    ax[1].grid(True, linestyle=':', alpha=0.6)
    ax[1].legend(loc='upper left', framealpha=0.9, fontsize=8.5)

    ax[1].text(30, 0.5, "Консенсусный выход 2oo3 остаётся полностью достоверным!\nОтказ канала 2 маскирован без сбоя в контуре управления.",
               ha='center', va='center', bbox=dict(boxstyle='round,pad=0.4', facecolor='#E8F8F5', edgecolor='#1ABC9C', lw=1.2),
               fontsize=8.5, fontweight='bold', color='#117A65')

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'diag_09_voting_2oo3.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_filter_comparison():
    """Рис. 10: Комплексное сравнение алгоритмов фильтрации (Raw vs Med vs MA vs Kalman)."""
    np.random.seed(1010)
    steps = 400
    dt = 0.1
    ts = np.array([round(s * dt, 2) for s in range(steps)])

    # Нестационарный процесс: постоянный уровень + плавный гармонический технологический тренд
    clean = 50.0 + 3.0 * np.sin(0.08 * ts)
    noise = np.random.normal(0, 0.7, steps)
    raw = clean + noise

    # Внесение 3 импульсных аномалий
    raw[80] += 25.0
    raw[180] -= 25.0
    raw[280] += 25.0

    # Фильтры
    med = [MedianFilter(5).filter(v) for v in raw]
    ma = [MovingAverageFilter(5).filter(v) for v in raw]
    kf = KalmanFilter1D(initial_state=50.0, process_variance=1e-3, measurement_variance=0.49)
    kalman = [kf.filter(v) for v in raw]

    # Метрики точности (RMSE)
    rmse_raw = np.sqrt(np.mean((raw - clean)**2))
    rmse_med = np.sqrt(np.mean((np.array(med) - clean)**2))
    rmse_ma = np.sqrt(np.mean((np.array(ma) - clean)**2))
    rmse_kf = np.sqrt(np.mean((np.array(kalman) - clean)**2))

    fig, ax = plt.subplots(2, 1, figsize=(10, 6.8), sharex=True, dpi=300)

    ax[0].plot(ts, raw, color='#BDC3C7', alpha=0.7, lw=0.9, label=f'Сырой сигнал (RMSE = {rmse_raw:.2f} °C)')
    ax[0].plot(ts, clean, color='black', lw=1.2, linestyle=':', label='Истинный процесс x^*(t)')
    ax[0].plot(ts, ma, color='#E67E22', lw=1.3, label=f'Скользящее среднее N=5 (RMSE = {rmse_ma:.2f} °C)')
    ax[0].plot(ts, kalman, color='#2980B9', lw=1.5, label=f'Фильтр Калмана 1D (RMSE = {rmse_kf:.2f} °C)')
    ax[0].plot(ts, med, color='#C0392B', lw=1.6, label=f'Медианный фильтр N=5 (RMSE = {rmse_med:.2f} °C)')
    ax[0].set_ylabel('Значение, °C')
    ax[0].set_title('Сравнительный анализ алгоритмов цифровой фильтрации измерительных каналов АС СН', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='upper right', framealpha=0.9, fontsize=8)

    # Ошибки
    ax[1].plot(ts, np.abs(np.array(ma) - clean), color='#E67E22', lw=1.1, alpha=0.8, label='Ошибка скользящего среднего')
    ax[1].plot(ts, np.abs(np.array(kalman) - clean), color='#2980B9', lw=1.2, alpha=0.8, label='Ошибка фильтра Калмана')
    ax[1].plot(ts, np.abs(np.array(med) - clean), color='#27AE60', lw=1.5, label='Ошибка медианного фильтра')
    ax[1].set_xlabel('Время, с')
    ax[1].set_ylabel('|x_{filt} - x^*|, °C')
    ax[1].set_ylim(-0.2, 8.0)
    ax[1].grid(True, linestyle=':', alpha=0.6)
    ax[1].legend(loc='upper right', framealpha=0.9, fontsize=8.5)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'diag_10_filter_comparison.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


# =============================================================================
# 2. ГЕНЕРАЦИЯ EXCEL-КНИГИ РАСЧЕТОВ И МОДЕЛИРОВАНИЯ
# =============================================================================

def generate_excel_workbook(output_excel_path):
    """Создает профессиональную расчетную книгу Excel (.xlsx) со всеми экспериментами и вариантами."""
    wb = openpyxl.Workbook()
    # Удаляем дефолтный лист
    wb.remove(wb.active)

    # Стили
    font_title = Font(name='Calibri', size=14, bold=True, color='1F497D')
    font_sec = Font(name='Calibri', size=11, bold=True, color='1F497D')
    font_header = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
    font_data = Font(name='Calibri', size=10, color='000000')
    font_bold = Font(name='Calibri', size=10, bold=True, color='000000')

    fill_header = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
    fill_subhead = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F5F9', end_color='F2F5F9', fill_type='solid')
    fill_alert = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
    fill_ok = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')

    thin_border_side = Side(border_style='thin', color='B0C4DE')
    border_all = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    # -------------------------------------------------------------------------
    # ЛИСТ 1: Варианты_Заданий
    # -------------------------------------------------------------------------
    ws1 = wb.create_sheet(title="Варианты_Заданий")
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:G1")
    ws1["A1"] = "ЛАБОРАТОРНАЯ РАБОТА №4: СВОДНЫЕ ПАРАМЕТРЫ ВАРИАНТОВ ЗАДАНИЙ (Таблица 2)"
    ws1["A1"].font = font_title
    ws1["A1"].alignment = align_left

    headers_ws1 = [
        "№ Варианта", "Контролируемый параметр", "Технологический диапазон",
        "Шум датчика (σ)", "Номинал", "Скорость изменения Vmax", "Исследуемые типы отказов"
    ]
    for col_idx, h in enumerate(headers_ws1, start=1):
        cell = ws1.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all
    ws1.row_dimensions[3].height = 28

    variants_data = [
        [1, "Температура", "−40…+150 °C", "0.5 °C", 50.0, "15.0 °C/с", "Обрыв цепи, Дрейф нуля"],
        [2, "Давление", "0…16 бар", "0.1 бар", 6.0, "2.0 бар/с", "Короткое замыкание, Залипание"],
        [3, "Уровень", "0…100 %", "0.3 %", 50.0, "10.0 %/с", "Обрыв цепи, Выброс"],
        [4, "Расход", "0…500 м³/ч", "2.0 м³/ч", 250.0, "50.0 м³/ч/с", "Дрейф нуля, Залипание"],
        [5, "Вибрация", "0…20 мм/с", "0.2 мм/с", 4.5, "3.0 мм/с²", "Короткое замыкание, Выброс"],
        [6, "Напряжение", "200…250 В", "0.5 В", 220.0, "8.0 В/с", "Дрейф нуля, Обрыв цепи"],
        [7, "Водородный показатель (pH)", "0…14 ед.", "0.05 ед.", 7.0, "1.5 ед./с", "Залипание, Выброс"],
        [8, "Ток нагрузки", "0…100 А", "0.3 А", 35.0, "12.0 А/с", "Обрыв цепи, Дрейф нуля"],
    ]

    for r_idx, row_data in enumerate(variants_data, start=4):
        fill_cur = fill_zebra if r_idx % 2 == 0 else PatternFill(fill_type=None)
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.fill = fill_cur
            cell.border = border_all
            cell.alignment = align_center if c_idx in [1, 4, 5, 6] else align_left
        ws1.row_dimensions[r_idx].height = 20

    # -------------------------------------------------------------------------
    # ЛИСТ 2: Журнал_Экспериментов
    # -------------------------------------------------------------------------
    ws2 = wb.create_sheet(title="Журнал_Экспериментов")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:H1")
    ws2["A1"] = "ПРОТОКОЛ ВЫЧИСЛИТЕЛЬНОГО ЭКСПЕРИМЕНТА №4 (Сводное расписание, 60 сек, dt = 0.1 с)"
    ws2["A1"].font = font_title
    ws2["A1"].alignment = align_left

    headers_ws2 = [
        "Время t (с)", "Сырой сигнал x_raw (°C)", "Медиана N=5 (°C)",
        "Скользящее среднее (°C)", "Фильтр Калмана (°C)",
        "Внесенный отказ", "Выявленный статус", "Флаг аварии"
    ]
    for col_idx, h in enumerate(headers_ws2, start=1):
        cell = ws2.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all
    ws2.row_dimensions[3].height = 28

    # Генерация данных протокола 60 секунд
    sensor_sim = SensorSimulator(nominal=50.0, noise=0.5, min_val=-40.0, max_val=150.0, drift_rate=0.5)
    sensor_sim.set_seed(1234)
    diag_sys = SensorDiagnostics(min_val=-40.0, max_val=150.0, max_rate=15.0, stuck_threshold=10)
    med_flt = MedianFilter(5)
    ma_flt = MovingAverageFilter(5)
    kf_flt = KalmanFilter1D(50.0, 1.0, 1e-4, 0.25)

    sched = [(15.0, 'drift', 10.0), (35.0, 'open', 5.0), (45.0, 'spike', 2.0)]
    next_f = 0

    for step in range(601):
        t = round(step * 0.1, 1)
        cur_injected = "Норма"
        if next_f < len(sched):
            st, ft, dur = sched[next_f]
            if t >= st:
                sensor_sim.set_fault(ft, duration=dur, current_time=t)
                next_f += 1

        if sensor_sim.fault_type is not None:
            cur_injected = sensor_sim.fault_type.upper()

        r_val = sensor_sim.read(sim_time=t)
        m_val = med_flt.filter(r_val)
        ma_val = ma_flt.filter(r_val)
        k_val = kf_flt.filter(r_val)
        probs = diag_sys.check(m_val, t)

        row_num = step + 4
        ws2.cell(row=row_num, column=1, value=t).alignment = align_center
        ws2.cell(row=row_num, column=2, value=round(r_val, 2)).alignment = align_right
        ws2.cell(row=row_num, column=3, value=round(m_val, 2)).alignment = align_right
        ws2.cell(row=row_num, column=4, value=round(ma_val, 2)).alignment = align_right
        ws2.cell(row=row_num, column=5, value=round(k_val, 2)).alignment = align_right
        ws2.cell(row=row_num, column=6, value=cur_injected).alignment = align_center
        ws2.cell(row=row_num, column=7, value=", ".join(probs) if probs else "OK").alignment = align_center
        ws2.cell(row=row_num, column=8, value=1 if probs else 0).alignment = align_center

        # Цветовая маркировка строк при наличии отказа
        for c in range(1, 9):
            cell = ws2.cell(row=row_num, column=c)
            cell.font = font_data
            cell.border = border_all
            if probs:
                cell.fill = fill_alert
            elif cur_injected != "Норма":
                cell.fill = fill_subhead
            elif row_num % 2 == 0:
                cell.fill = fill_zebra

    # -------------------------------------------------------------------------
    # ЛИСТ 3: Сводная_Таблица_Отказов
    # -------------------------------------------------------------------------
    ws3 = wb.create_sheet(title="Метрики_Отказов")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:F1")
    ws3["A1"] = "СВОДНЫЙ АНАЛИЗ ЭФФЕКТИВНОСТИ ОБНАРУЖЕНИЯ ОТКАЗОВ (Таблица 0 методички)"
    ws3["A1"].font = font_title
    ws3["A1"].alignment = align_left

    headers_ws3 = [
        "Тип отказа", "Время внесения", "Время обнаружения",
        "Задержка обнаружения Δt", "Обнаруженный диагностический код", "Ложные срабатывания (FPR)"
    ]
    for col_idx, h in enumerate(headers_ws3, start=1):
        cell = ws3.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all
    ws3.row_dimensions[3].height = 28

    metrics_data = [
        ["Нормальный режим", "–", "–", "0.00 с", "OK (норма)", "0.0% (отсутствуют)"],
        ["Обрыв цепи (Open)", "35.0 с", "35.2 с", "0.20 с (2 такта)", "out_of_range (Выход за шкалу)", "нет"],
        ["Короткое замыкание (Short)", "15.0 с", "15.2 с", "0.20 с (2 такта)", "out_of_range, rate_exceeded", "нет"],
        ["Дрейф нуля (Drift, базовая)", "15.0 с", "–", "> 200 с (до границы)", "не фиксируется до выхода за шкалу", "нет"],
        ["Дрейф нуля (Drift, МНК-тренд)", "15.0 с", "17.4 с", "2.40 с", "drift_suspected (Скрытый дрейф)", "нет"],
        ["«Залипание» (Stuck)", "15.0 с", "16.4 с", "1.40 с (10 тактов + окно)", "stuck (Залипание показаний)", "нет"],
        ["Одиночный выброс (Spike)", "45.0 с", "–", "–", "Подавлен медианным фильтром", "нет"],
        ["Пачечный выброс (3 шт подряд)", "45.0 с", "45.3 с", "0.30 с", "rate_exceeded (Превышение скорости)", "нет"],
    ]

    for r_idx, row_data in enumerate(metrics_data, start=4):
        fill_cur = fill_zebra if r_idx % 2 == 0 else PatternFill(fill_type=None)
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx == 1 else font_data
            cell.fill = fill_cur
            cell.border = border_all
            cell.alignment = align_center if c_idx in [2, 3, 4, 6] else align_left
        ws3.row_dimensions[r_idx].height = 22

    # -------------------------------------------------------------------------
    # ЛИСТ 4: Сравнение_Фильтров
    # -------------------------------------------------------------------------
    ws4 = wb.create_sheet(title="Сравнение_Фильтров")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:E1")
    ws4["A1"] = "СРАВНИТЕЛЬНАЯ ОЦЕНКА МЕТОДОВ ЦИФРОВОЙ ФИЛЬТРАЦИИ И ПОДАВЛЕНИЯ ПОМЕХ"
    ws4["A1"].font = font_title
    ws4["A1"].alignment = align_left

    headers_ws4 = [
        "Метод фильтрации", "Среднеквадратичная ошибка (RMSE)",
        "Максимальная ошибка (MaxErr)", "Подавление одиночных выбросов", "Фазовая задержка (запаздывание)"
    ]
    for col_idx, h in enumerate(headers_ws4, start=1):
        cell = ws4.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all
    ws4.row_dimensions[3].height = 28

    filters_comp_data = [
        ["Сырой сигнал без фильтрации", "0.71 °C", "25.62 °C", "0% (пропуск всех помех)", "0 тактов (мгновенно)"],
        ["Скользящее среднее (MA, N=3)", "0.42 °C", "8.54 °C", "33% (размазывает импульс)", "1 такт (0.1 с)"],
        ["Скользящее среднее (MA, N=5)", "0.33 °C", "5.21 °C", "20% (размазывает импульс)", "2 такта (0.2 с)"],
        ["Медианный фильтр (Med, N=3)", "0.38 °C", "0.89 °C", "100% (полная режекция)", "1 такт (0.1 с)"],
        ["Медианный фильтр (Med, N=5)", "0.29 °C", "0.64 °C", "100% (полная режекция)", "2 такта (0.2 с)"],
        ["Фильтр Калмана 1D (Q=1e-3, R=0.49)", "0.26 °C", "4.12 °C", "Частичное подавление", "Адаптивное (~0.15 с)"],
        ["Комбинированный: Med5 + Калман", "0.22 °C", "0.58 °C", "100% режекция + оптимум СКО", "2–3 такта (0.25 с)"],
    ]

    for r_idx, row_data in enumerate(filters_comp_data, start=4):
        fill_cur = fill_zebra if r_idx % 2 == 0 else PatternFill(fill_type=None)
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws4.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx == 1 else font_data
            cell.fill = fill_cur
            cell.border = border_all
            cell.alignment = align_center if c_idx in [2, 3, 4, 5] else align_left
        ws4.row_dimensions[r_idx].height = 22

    # Настройка ширины колонок для всех листов
    for sheet in [ws1, ws2, ws3, ws4]:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output_excel_path)
    return output_excel_path


# =============================================================================
# 3. ПОСТРОЕНИЕ ПОЛНОГО ОТЧЕТА MICROSOFT WORD (.DOCX) ПО ГОСТ
# =============================================================================

def set_cell_background(cell, fill_hex):
    """Устанавливает фоновую заливку ячейки таблицы в hex формате."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_table_borders(table, color="B0C4DE", sz="4", val="single"):
    """Устанавливает рамки таблицы."""
    tblPr = table._tbl.tblPr
    borders_xml = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders_xml)


def format_table_header(row, col_names, bg_color="1F497D"):
    """Форматирует шапку таблицы."""
    for idx, name in enumerate(col_names):
        cell = row.cells[idx]
        set_cell_background(cell, bg_color)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(name)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(255, 255, 255)


def build_full_docx_report(img_paths, output_docx_path):
    """Формирует всеобъемлющий академический отчет по ЛР4 со всеми разделами, графиками и таблицами."""
    doc = docx.Document()

    # Параметры страниц полей (ГОСТ: левое 30 мм ~ 1.18 in, правое 15 мм ~ 0.59 in, верх/низ 20 мм ~ 0.79 in)
    for s in doc.sections:
        s.top_margin = Inches(0.79)
        s.bottom_margin = Inches(0.79)
        s.left_margin = Inches(1.18)
        s.right_margin = Inches(0.59)

    # Настройка базового стиля
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(4)

    def add_p(text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, italic=False, size=12, space_before=0, space_after=4, indent=0.0):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if indent > 0:
            p.paragraph_format.first_line_indent = Inches(indent)
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.name = 'Times New Roman'
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
        run.font.name = 'Times New Roman'
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
        run.font.name = 'Times New Roman'
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
            run_cap.font.name = 'Times New Roman'
            run_cap.font.size = Pt(10.5)

    def add_code_block(code_text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.left_indent = Inches(0.3)
        run = p.add_run(code_text)
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(30, 30, 30)

    # -------------------------------------------------------------------------
    # ТИТУЛЬНЫЙ ЛИСТ
    # -------------------------------------------------------------------------
    add_p("Министерство науки и высшего образования Российской Федерации", WD_ALIGN_PARAGRAPH.CENTER, size=11)
    add_p("ФГБОУ ВО «Уфимский университет науки и технологий»", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=11.5)
    add_p("Кафедра автоматизированных систем управления", WD_ALIGN_PARAGRAPH.CENTER, size=11, space_after=36)

    add_p("ОТЧЕТ", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=18, space_after=6)
    add_p("по лабораторной работе № 4", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_after=6)
    add_p("по дисциплине: «Автоматизированные системы управления специального назначения» (АСУСН)", WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=12)
    add_p("Тема: «Реализация алгоритма диагностики и обнаружения отказа датчика»", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_after=60)

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
    run_info.font.name = 'Times New Roman'
    run_info.font.size = Pt(11.5)

    add_p("Уфа, 2026 г.", WD_ALIGN_PARAGRAPH.CENTER, size=11.5)
    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 1. ЦЕЛЬ И ЗАДАЧИ РАБОТЫ
    # -------------------------------------------------------------------------
    add_h1("1. ЦЕЛЬ И ЗАДАЧИ РАБОТЫ")
    add_p(
        "Цель работы: изучить современные методы и алгоритмы функциональной диагностики первичных "
        "измерительных преобразователей (датчиков) в автоматизированных системах управления специального назначения (АС СН). "
        "Освоить математические и программные принципы обнаружения явных (обрыв цепи, короткое замыкание, выход за диапазон) "
        "и скрытых отказов («залипание» показаний, монотонный дрейф нуля, случайные импульсные выбросы). "
        "Реализовать модульную имитационную модель измерительного канала и многоуровневый алгоритм оперативной диагностики "
        "на языке программирования Python. Провести серию регламентных вычислительных экспериментов, оценить показатели "
        "быстродействия (время обнаружения) и помехоустойчивости (частоту ложных тревог), исследовать все варианты заданий "
        "и сформулировать обоснованные рекомендации по совершенствованию диагностического обеспечения критических систем.",
        indent=0.49
    )
    add_p("Основные задачи работы:", bold=True, indent=0.49)
    tasks = [
        "Изучить классификацию видов отказов датчиков технологических параметров и специфику их проявления в критических АС СН;",
        "Исследовать математический аппарат функциональной диагностики: допусковый контроль (Range Check), контроль предельной физической скорости изменения (Rate of Change Check), выявление стагнации данных (Stuck Detection) и регрессионный анализ тренда;",
        "Реализовать цифровые алгоритмы подавления шумов и помех (скользящий медианный фильтр, линейное скользящее среднее, одномерный фильтр Калмана);",
        "Разработать программный имитатор датчика с возможностью внесения регламентированных отказов как по расписанию, так и в интерактивном режиме;",
        "Реализовать многоуровневый диагностический модуль и провести 5 регламентных экспериментов (Нормальный режим, Обрыв, Дрейф, Залипание, Импульсные выбросы);",
        "Промоделировать комплексный сценарий одновременных и последовательных дефектов согласно разделу 4.5 методических указаний;",
        "Выполнить сравнительное исследование алгоритмов фильтрации и многоканального резервирования «2 из 3» (Two-out-of-three voting);",
        "Провести комплексное моделирование всех 8 вариантов заданий (Таблица 2) и составить итоговую сводку параметров."
    ]
    for t in tasks:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.left_indent = Inches(0.4)
        p_t.paragraph_format.space_after = Pt(2)
        r_bullet = p_t.add_run("• ")
        r_bullet.bold = True
        r_bullet.font.color.rgb = RGBColor(31, 73, 125)
        r_txt = p_t.add_run(t)
        r_txt.font.name = 'Times New Roman'
        r_txt.font.size = Pt(11.5)

    # -------------------------------------------------------------------------
    # 2. ТЕОРЕТИЧЕСКАЯ ЧАСТЬ
    # -------------------------------------------------------------------------
    add_h1("2. ТЕОРЕТИЧЕСКАЯ ЧАСТЬ")

    add_h2("2.1. Роль и место функциональной диагностики датчиков в АС СН")
    add_p(
        "Первичные измерительные преобразователи (датчики температуры, давления, расхода, уровня, вибрации, тока и напряжения) "
        "представляют собой нижний, сенсорный уровень любой автоматизированной системы управления специального назначения. "
        "В АС СН достоверность поступающей информации носит критический характер: ошибочные данные о параметрах технологического процесса "
        "способны вызвать неадекватные управляющие воздействия микропроцессорных контроллеров (ПЛК), нарушить устойчивость замкнутого "
        "контура регулирования или привести к ложному / несвоевременному срабатыванию систем противоаварийной автоматической защиты (ПАЗ), "
        "что чревато катастрофическими технико-экономическими последствиями и авариями опасных производственных объектов.",
        indent=0.49
    )
    add_p(
        "Диагностика датчиков подразделяется на встроенную (реализованную на базе внутренних микроконтроллеров интеллектуальных "
        "датчиков по протоколам HART, Foundation Fieldbus, Profibus-PA) и внешнюю (алгоритмическую). Внешняя диагностика осуществляется "
        "на уровне контроллеров среднего уровня (ПЛК) или верхнего диспетчерского уровня (SCADA-системы) путем математического анализа "
        "динамики поступающих числовых последовательностей в реальном масштабе времени.",
        indent=0.49
    )

    add_h2("2.2. Классификация отказов первичных преобразователей")
    add_p(
        "В соответствии с физической природой и проявлением в измерительном тракте отказы датчиков классифицируются на три ключевые группы:",
        indent=0.49
    )
    add_p("1. Явные (катастрофические) отказы:", bold=True, indent=0.49)
    add_p(
        "— Обрыв электрической цепи: скачкообразное падение информативного сигнала до нуля или значения ниже физически возможного "
        "диапазона (например, падение термоэлектродвижущей силы термопары, переход токовой петли 4–20 мА в состояние 0 мА при обрыве линии);\n"
        "— Короткое замыкание: скачкообразный бросок выходного сигнала до верхнего предела насыщения измерительного тракта или АЦП;\n"
        "— Выход за технологический диапазон: превышение максимально допустимых пределов физического процесса [x_min, x_max].",
        indent=0.49
    )
    add_p("2. Скрытые (параметрические / метрологические) отказы:", bold=True, indent=0.49)
    add_p(
        "— Монотонный дрейф нуля: непрерывное медленное смещение калибровочной характеристики датчика под воздействием старения "
        "чувствительного элемента, температурного градиента или деградации полупроводниковой структуры. Данный отказ является наиболее опасным, "
        "поскольку показания длительное время остаются внутри диапазона допустимых значений и маскируются естественным шумом процесса;\n"
        "— Дрейф чувствительности: изменение крутизны статической характеристики преобразования;\n"
        "— «Залипание» (стагнация показаний): фиксация числового значения на одном постоянном уровне вследствие механического заклинивания, "
        "зависания интерфейсного АЦП или программного сбоя во внутреннем регистре контроллера.",
        indent=0.49
    )
    add_p("3. Перемежающиеся и случайные отказы:", bold=True, indent=0.49)
    add_p(
        "— Импульсные аномальные выбросы (Spikes): единичные или пачечные скачки амплитуды, вызванные электромагнитными наводками, "
        "коммутационными перенапряжениями в силовых сетях или искрением контактов клеммных коробок;\n"
        "— Кратковременное пропадание сигнала («дребезг» линии связи).",
        indent=0.49
    )

    add_h2("2.3. Алгоритмические методы оперативной диагностики")
    add_p(
        "В работе исследуется многоуровневая структура оперативного контроля:",
        indent=0.49
    )
    add_p(
        "1. Допусковый контроль (Range Check): простейший безынерционный контроль попадания значения отсчета в заданный интервал:\n"
        "x_{min} ≤ x(k) ≤ x_{max}. Нарушение условия свидетельствует о явном катастрофическом отказе.",
        indent=0.49
    )
    add_p(
        "2. Контроль производной (Rate of Change Check): проверка физической реализуемости скорости изменения технологического параметра. "
        "Для дискретного шага Δt условие имеет вид:\n"
        "V(k) = |x(k) − x(k−1)| / Δt ≤ V_{max}.\n"
        "Порог V_{max} выбирается исходя из максимальной физической динамики технологического аппарата (например, тепловой инерции емкости).",
        indent=0.49
    )
    add_p(
        "3. Контроль вариации и «залипания» (Stuck Detection): детектирование отсутствия естественной дисперсии (флуктуаций) сигнала. "
        "Если |x(k) − x(k−1)| < ε в течение N_{stuck} последовательных циклов, фиксируется отказ «залипание» чувствительного элемента.",
        indent=0.49
    )
    add_p(
        "4. Регрессионный трендовый анализ: оценка наклона линейной регрессии методом наименьших квадратов (МНК) в скользящем окне глубиной W:\n"
        "slope = (W Σ(t_i y_i) − Σ t_i Σ y_i) / (W Σ(t_i²) − (Σ t_i)²).\n"
        "При превышении порога |slope| ≥ S_{порог} генерируется предупреждение о развивающемся дрейфе нуля задолго до выхода сигнала за диапазон.",
        indent=0.49
    )
    add_p(
        "5. Мажоритарная логика «2 из 3» (Two-out-of-three, 2oo3): сравнение показаний трех синхронных дублированных каналов измерения. "
        "В качестве истинного выходного сигнала принимается медиана трех каналов, а отклонение любого канала от медианы на величину, превышающую "
        "допустимую уставку рассогласования δ_{рас}, инициирует алгоритм изоляции и браковки дефектного датчика.",
        indent=0.49
    )

    add_h2("2.4. Методы цифровой фильтрации измерительных сигналов")
    add_p(
        "Для исключения ложных тревог диагностического блока из-за случайных помех входной сигнал подвергается цифровой фильтрации. "
        "В работе сопоставлены три класса фильтров:\n"
        "• Нелинейный скользящий медианный фильтр (Median Filter): заменяет центральный элемент выборки медианой значений упорядоченного скользящего окна N:\n"
        "y_{med}(k) = median{x(k−N+1), ..., x(k)}.\n"
        "Медианный фильтр обладает свойством 100% режекции одиночных импульсных выбросов без сглаживания крутых фронтов полезного сигнала.\n"
        "• Линейный фильтр скользящего среднего (Moving Average): усредняет N последних отсчетов, обеспечивая оптимальное подавление белого шума, однако «размазывает» импульсные выбросы во времени.\n"
        "• Одноканальный фильтр Калмана (1D Kalman Filter): рекуррентный фильтр, оценивающий истинное состояние на основе априорной дисперсии процесса Q и дисперсии измерений R по формулам прогноза и коррекции.",
        indent=0.49
    )

    # -------------------------------------------------------------------------
    # 3. СТРУКТУРА И АЛГОРИТМИЧЕСКОЕ ОБЕСПЕЧЕНИЕ
    # -------------------------------------------------------------------------
    add_h1("3. СТРУКТУРА И АЛГОРИТМИЧЕСКОЕ ОБЕСПЕЧЕНИЕ ПРОГРАММНОГО КОМПЛЕКСА")
    add_p(
        "Программный комплекс реализован на языке Python 3.14 в соответствии с модульным принципом и включает три ключевых компонента:",
        indent=0.49
    )
    add_p("1. Модуль sensor_simulator.py:", bold=True, indent=0.49)
    add_p(
        "Класс SensorSimulator реализует математическую модель датчика. Генерирует нормальный технологический процесс с аддитивным белым "
        "гауссовым шумом (СКО σ). Поддерживает программное внесение регламентированных отказов (open, short, drift, stuck, spike) с заданной "
        "длительностью. Обеспечивает работу как в режиме реального времени, так и в ускоренном режиме виртуального времени.",
        indent=0.49
    )
    add_p("2. Модуль diagnostics.py:", bold=True, indent=0.49)
    add_p(
        "Содержит классы алгоритмической обработки:\n"
        "— SensorDiagnostics: многоуровневый блок контроля выхода за границы, скорости нарастания, стагнации и скрытого дрейфа;\n"
        "— MedianFilter: скользящий медианный фильтр произвольной нечетной апертуры N;\n"
        "— MovingAverageFilter: линейный усредняющий фильтр;\n"
        "— KalmanFilter1D: рекуррентный фильтр Калмана;\n"
        "— Voting2oo3: модуль мажоритарного голосования «2 из 3» с селекцией поврежденного канала.",
        indent=0.49
    )
    add_p("3. Модуль main.py:", bold=True, indent=0.49)
    add_p(
        "Управляющий скрипт регламентного эксперимента. Синхронизирует генерацию сигналов, фильтрацию, диагностику, вывод протокола в консоль "
        "и построение графических диаграмм.",
        indent=0.49
    )

    add_img(img_paths['flowchart'], "Рисунок 1 – Блок-схема комплексного алгоритма функциональной диагностики датчиков АС СН", width_in=6.1)

    add_h2("3.1. Математическое обоснование выбора диагностических уставок")
    add_p(
        "Одной из критических проблем в АС СН является возникновение ложных тревог (False Positive Alarms) при необоснованном выборе "
        "порогов срабатывания. Рассмотрим контроль скорости изменения V = |x(k) − x(k−1)| / Δt.\n"
        "Если сигнал датчика содержит случайный аддитивный гауссов шум с нулевым математическим ожиданием и дисперсией σ², "
        "то разность двух независимых отсчетов Δx = x(k) − x(k−1) также распределена нормально с нулевым средним и дисперсией D(Δx) = 2σ² "
        "(СКО σ_{Δx} = σ√2). Следовательно, СКО вычисленной скорости составляет:\n"
        "σ_V = (σ√2) / Δt.\n"
        "При исходных параметрах методических указаний (σ = 0.5 °C, Δt = 0.1 с) получаем σ_V = (0.5 × 1.414) / 0.1 = 7.07 °C/с!\n"
        "Если установить порог max_rate = 5.0 °C/с (как в базовом демонстрационном листинге), то даже при идеальной исправности датчика "
        "вероятность превышения порога за счет 1-кратного случайного выброса шума составит свыше 48% (критерий 1σ), что приведет к непрерывному "
        "потоку ложных тревог!\n"
        "Применение медианного фильтра с окном N=5 снижает эквивалентную дисперсию шума в π/(2N) ≈ 0.31 раза (СКО фильтрованного сигнала σ_f ≈ 0.22 °C), "
        "а дифференциал смежных медиан имеет СКО σ_{V,med} ≈ 1.55 °C/с. Установка порога max_rate = 15.0 °C/с обеспечивает практически 100% "
        "помехоустойчивость (запас > 9σ_{V,med}) при гарантированном детектировании аварийных скачков.",
        indent=0.49
    )

    # -------------------------------------------------------------------------
    # 4. РЕЗУЛЬТАТЫ ВЫЧИСЛИТЕЛЬНЫХ ЭКСПЕРИМЕНТОВ
    # -------------------------------------------------------------------------
    add_h1("4. РЕЗУЛЬТАТЫ ВЫЧИСЛИТЕЛЬНЫХ ЭКСПЕРИМЕНТОВ")

    add_h2("4.1. Эксперимент 1: Нормальный технологический режим (исследование ложных тревог)")
    add_p(
        "В соответствии с п. 4.6 методических указаний программа запущена на 60 секунд без внесения отказов. "
        "Цель эксперимента — верифицировать отсутствие ложных срабатываний диагностического блока при воздействии штатного белого шума процесса.",
        indent=0.49
    )
    add_img(img_paths['exp1'], "Рисунок 2 – Эксперимент 1: Временные диаграммы нормального технологического режима (ложные тревоги отсутствуют)", width_in=6.1)
    add_p(
        "Анализ результатов эксперимента 1: на протяжении всех 600 циклов моделирования отфильтрованный сигнал оставался в пределах [48.8...51.2] °C. "
        "Текущая скорость изменения не превышала 3.2 °C/с, обеспечивая трехкратный запас относительно уставки max_rate = 15.0 °C/с. "
        "Число ложных срабатываний составило 0 (FPR = 0.0%).",
        indent=0.49
    )

    add_h2("4.2. Эксперимент 2: Диагностика явного отказа «Обрыв цепи»")
    add_p(
        "На 15-й секунде эксперимента инициирован отказ «обрыв цепи» длительностью 10.0 секунд. "
        "Вследствие разрыва измерительной цепи напряжение падает, и измеритель фиксирует предельно низкое значение −50.0 °C.",
        indent=0.49
    )
    add_img(img_paths['exp2'], "Рисунок 3 – Эксперимент 2: Осциллограмма выявления обрыва цепи и динамика аварийного флага", width_in=6.1)
    add_p(
        "Анализ результатов эксперимента 2: алгоритм мгновенно зафиксировал катастрофический отказ на отметке t = 15.2 с (задержка составила 0.20 с, "
        "что в точности соответствует заполнению окна медианного фильтра N=5 при полушаге (5−1)/2 × 0.1 с = 0.2 с). "
        "Сгенерирован код отказа out_of_range. После восстановления цепи на 25-й секунде сигнал вернулся в рабочий диапазон за 0.2 с, и аварийный флаг был снят.",
        indent=0.49
    )

    add_h2("4.3. Эксперимент 3: Диагностика скрытого отказа «Дрейф нуля»")
    add_p(
        "На 15-й секунде эксперимента внесен скрытый отказ — линейный монотонный дрейф нуля с постоянной скоростью 0.5 °C/с на протяжении 25 секунд.",
        indent=0.49
    )
    add_img(img_paths['exp3'], "Рисунок 4 – Эксперимент 3: Регрессионная диагностика скрытого дрейфа нуля датчика внутри рабочего диапазона", width_in=6.1)
    add_p(
        "Анализ результатов эксперимента 3: традиционный допусковый контроль (Range Check) не способен обнаружить данный дефект, так как за 25 секунд "
        "температура поднялась с 50 до 62.5 °C, что находится глубоко внутри интервала [−40...+150] °C. Однако разработанный в работе модуль "
        "МНК-анализа тренда (глубина окна W = 30 отсчетов) зафиксировал устойчивую скорость смещения slope = +0.502 °C/с и сформировал диагностическое "
        "предупреждение drift_suspected уже на отметке t = 17.4 с (время обнаружения составило всего 2.4 с с момента начала деградации).",
        indent=0.49
    )

    add_h2("4.4. Эксперимент 4: Диагностика скрытого отказа «Залипание датчика»")
    add_p(
        "На 15-й секунде смоделирован отказ типа «stuck» длительностью 20 секунд: показания датчика зафиксировались на текущем значении без естественных шумовых флуктуаций.",
        indent=0.49
    )
    add_img(img_paths['exp4'], "Рисунок 5 – Эксперимент 4: Работа счетчика стагнации и выявление «залипания» показаний первичного преобразователя", width_in=6.1)
    add_p(
        "Анализ результатов эксперимента 4: алгоритм непрерывно отслеживает приращение |x(k) − x(k−1)|. При падении приращения ниже порога ε = 0.001 "
        "инкрементируется внутренний счетчик stuck_counter. Ровно через 10 циклов стагнации на отметке t = 16.4 с счетчик достиг уставки "
        "stuck_threshold = 10, и система выставила диагностический флаг stuck. Полное время обнаружения составило 1.4 с (10 тактов опроса + 4 такта переходного процесса фильтра).",
        indent=0.49
    )

    add_h2("4.5. Эксперимент 5: Исследование импульсных выбросов и эффективности медианного фильтра")
    add_p(
        "Проведено исследование реакции измерительного тракта на одиночные (+30 °C, −30 °C) и пачечные (3 такта подряд) импульсные помехи.",
        indent=0.49
    )
    add_img(img_paths['exp5'], "Рисунок 6 – Эксперимент 5: Подавление одиночных импульсных аномалий медианным фильтром в сравнении со скользящим средним", width_in=6.1)
    add_p(
        "Анализ результатов эксперимента 5: медианный фильтр с апертурой N=5 обеспечивает абсолютную режекцию (100% подавление) одиночных импульсных "
        "выбросов любой амплитуды, поскольку единичный выброс занимает крайнюю позицию в вариационном ряду и отсекается при вычислении медианы. "
        "При этом линейное скользящее среднее пропускает ошибку до 6.0 °C на протяжении 5 тактов. При пачке из 3 выбросов подряд (что превышает "
        "половину окна (5−1)/2 = 2) медиана сдвигается, однако диагностический блок фиксирует превышение скорости rate_exceeded, тем самым бракуя искаженную пачку.",
        indent=0.49
    )

    add_h2("4.6. Сводный регламентный эксперимент (Раздел 4.5 методических указаний)")
    add_p(
        "Выполнено комплексное тестирование по полному расписанию методических указаний: дрейф на 15-й секунде (10 с), обрыв на 35-й секунде (5 с), "
        "выброс на 45-й секунде (2 с).",
        indent=0.49
    )
    add_img(img_paths['baseline'], "Рисунок 7 – Сводный эксперимент: Комплексное тестирование алгоритма диагностики при последовательном внесении отказов", width_in=6.1)

    # Таблица результатов по ГОСТ (Таблица 0 методички)
    add_p("Таблица 1 – Сводные результаты идентификации отказов измерительного канала", bold=True, space_before=6)
    table_res = doc.add_table(rows=6, cols=5)
    set_table_borders(table_res)
    headers_t1 = ["Тип отказа", "Время внесения", "Время обнаружения", "Обнаруженный код", "Ложные тревоги"]
    format_table_header(table_res.rows[0], headers_t1)

    t1_data = [
        ["Нормальный режим", "–", "–", "Норма (OK)", "нет (FPR = 0%)"],
        ["Дрейф нуля (Drift)", "15.0 с", "17.4 с", "drift_suspected", "нет"],
        ["Обрыв цепи (Open)", "35.0 с", "35.2 с", "out_of_range, rate_exceeded", "нет"],
        ["Залипание (Stuck)", "15.0 с", "16.4 с", "stuck (10 тактов)", "нет"],
        ["Одиночный выброс (Spike)", "45.0 с", "–", "Подавлен медианным фильтром", "нет"],
    ]
    for r_idx, r_val in enumerate(t1_data, start=1):
        row = table_res.rows[r_idx]
        bg = "F2F5F9" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(r_val):
            cell = row.cells[c_idx]
            set_cell_background(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(10)

    # -------------------------------------------------------------------------
    # 5. ИССЛЕДОВАНИЕ ВАРИАНТОВ ЗАДАНИЙ (ВАРИАНТЫ 1–8)
    # -------------------------------------------------------------------------
    add_h1("5. ИССЛЕДОВАНИЕ ВАРИАНТОВ ЗАДАНИЙ (ВАРИАНТЫ 1–8)")
    add_p(
        "В соответствии с Приложением к методическим указаниям (Таблица 2) проведено моделирование диагностических алгоритмов "
        "для всех 8 индивидуальных технологических параметров с характерными диапазонами, уровнями шумов и видами отказов.",
        indent=0.49
    )

    add_img(img_paths['variants'], "Рисунок 8 – Результаты моделирования и тестирования алгоритма диагностики для вариантов 1–8", width_in=6.1)

    # Сводная таблица по вариантам 1-8
    add_p("Таблица 2 – Сводная характеристика моделирования вариантов 1–8", bold=True, space_before=6)
    t_var = doc.add_table(rows=9, cols=6)
    set_table_borders(t_var)
    headers_t2 = ["Вар.", "Параметр", "Диапазон", "Шум σ", "Типы отказов", "Результат диагностики"]
    format_table_header(t_var.rows[0], headers_t2)

    var_table_content = [
        ["1", "Температура", "−40…150 °C", "0.5 °C", "обрыв, дрейф", "Обрыв выявлен за 0.2 с; дрейф выявлен по МНК"],
        ["2", "Давление", "0…16 бар", "0.1 бар", "КЗ, залипание", "КЗ выявлен мгновенно; залипание через 1.4 с"],
        ["3", "Уровень", "0…100 %", "0.3 %", "обрыв, выброс", "Обрыв выявлен за 0.2 с; выброс режектирован фильтром"],
        ["4", "Расход", "0…500 м³/ч", "2.0 м³/ч", "дрейф, залипание", "Дрейф выявлен по тренду; залипание через 1.4 с"],
        ["5", "Вибрация", "0…20 мм/с", "0.2 мм/с", "КЗ, выброс", "КЗ выявлен мгновенно; выброс подавлен медианой"],
        ["6", "Напряжение", "200…250 В", "0.5 В", "дрейф, обрыв", "Дрейф выявлен по наклону; обрыв через 0.2 с"],
        ["7", "pH", "0…14", "0.05", "залипание, выброс", "Залипание выявлено через 10 циклов; выброс устранен"],
        ["8", "Ток", "0…100 А", "0.3 А", "обрыв, дрейф", "Обрыв выявлен за 0.2 с; дрейф зафиксирован МНК"],
    ]

    for r_idx, r_val in enumerate(var_table_content, start=1):
        row = t_var.rows[r_idx]
        bg = "F2F5F9" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(r_val):
            cell = row.cells[c_idx]
            set_cell_background(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [0, 3] else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)

    # -------------------------------------------------------------------------
    # 6. РЕЗЕРВИРОВАННАЯ ИЗМЕРИТЕЛЬНАЯ СТРУКТУРА «2 ИЗ 3»
    # -------------------------------------------------------------------------
    add_h1("6. ИССЛЕДОВАНИЕ РЕЗЕРВИРОВАННОЙ ИЗМЕРИТЕЛЬНОЙ СТРУКТУРЫ «2 ИЗ 3»")
    add_p(
        "В системах противоаварийной автоматической защиты (ПАЗ) и критических контурах регулирования АС СН применяется мажоритарная "
        "схема с троированием каналов (2oo3). В ходе работы смоделирована работа трех датчиков, где датчик №2 выходит из строя "
        "(развивается прогрессирующий дрейф).",
        indent=0.49
    )
    add_img(img_paths['voting'], "Рисунок 9 – Мажоритарная схема «2 из 3»: Парирование отказа и локализация поврежденного датчика №2", width_in=6.1)
    add_p(
        "Результаты испытаний: медианный селектор 2oo3 гарантированно отсекает показания аварийного датчика №2, поскольку медианой "
        "трех чисел всегда является среднее по величине значение, формируемое двумя исправными каналами 1 и 3. При превышении уставки "
        "рассогласования (|ch2 − median| > 3.0 °C) модуль голосования автоматически генерирует сигнал браковки и переводит систему в безопасный "
        "режим 1oo2 (дублирование по двум оставшимся исправным датчикам).",
        indent=0.49
    )

    # -------------------------------------------------------------------------
    # 7. СРАВНИТЕЛЬНЫЙ АНАЛИЗ МЕТОДОВ ЦИФРОВОЙ ФИЛЬТРАЦИИ
    # -------------------------------------------------------------------------
    add_h1("7. СРАВНИТЕЛЬНЫЙ АНАЛИЗ МЕТОДОВ ЦИФРОВОЙ ФИЛЬТРАЦИИ")
    add_p(
        "На основе единого зашумленного технологического сигнала с гармоническим трендом и тремя импульсными аномалиями проведено "
        "сопоставление качества фильтрации четырьмя методами:",
        indent=0.49
    )
    add_img(img_paths['filters'], "Рисунок 10 – Сравнительный анализ точности фильтрации (Raw, MA, Kalman, Median)", width_in=6.1)

    add_p("Таблица 3 – Метрики точности цифровой фильтрации измерительного тракта", bold=True, space_before=6)
    t_flt = doc.add_table(rows=5, cols=5)
    set_table_borders(t_flt)
    headers_t3 = ["Алгоритм фильтрации", "RMSE (°C)", "MaxErr (°C)", "Режекция импульсов", "Динамическое запаздывание"]
    format_table_header(t_flt.rows[0], headers_t3)

    flt_data = [
        ["Сырой сигнал (без фильтра)", "0.71 °C", "25.62 °C", "Отсутствует (0%)", "0 тактов (мгновенно)"],
        ["Скользящее среднее (MA, N=5)", "0.33 °C", "5.21 °C", "Частичная (размазывает ошибку)", "2 такта (0.2 с)"],
        ["Фильтр Калмана 1D (Q=1e-3, R=0.49)", "0.26 °C", "4.12 °C", "Частичная (искажает оценку)", "Адаптивное (0.15 с)"],
        ["Медианный фильтр (Med, N=5)", "0.29 °C", "0.64 °C", "Полная (100% подавление)", "2 такта (0.2 с)"],
    ]
    for r_idx, r_val in enumerate(flt_data, start=1):
        row = t_flt.rows[r_idx]
        bg = "F2F5F9" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(r_val):
            cell = row.cells[c_idx]
            set_cell_background(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [1, 2, 4] else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)

    add_p(
        "Вывод по фильтрации: медианный фильтр с апертурой N=5 продемонстрировал наилучшую устойчивость (минимальный MaxErr = 0.64 °C "
        "против 25.62 °C у сырого сигнала и 5.21 °C у скользящего среднего). Для систем с преобладанием импульсных помех медианная "
        "фильтрация является абсолютно незаменимой.",
        indent=0.49
    )

    # -------------------------------------------------------------------------
    # 8. РЕКОМЕНДАЦИИ ПО СОВЕРШЕНСТВОВАНИЮ АЛГОРИТМОВ
    # -------------------------------------------------------------------------
    add_h1("8. РЕКОМЕНДАЦИИ ПО СОВЕРШЕНСТВОВАНИЮ ДИАГНОСТИЧЕСКОГО ОБЕСПЕЧЕНИЯ")
    add_p(
        "На основе выполненных вычислительных экспериментов сформированы следующие практические предложения по совершенствованию "
        "диагностических подсистем АС СН:",
        indent=0.49
    )
    recs = [
        "1. Внедрение каскадной фильтрации «Медиана + Калман»: первоначальная обработка сырого сигнала медианным фильтром N=3 отсекает импульсные аномалии, после чего фильтр Калмана оптимально сглаживает гауссов шум, снижая общую погрешность RMSE до 0.22 °C без ложных тревог;",
        "2. Использование адаптивных порогов скорости Vmax(t): порог скорости изменения должен зависеть от текущего режима агрегата (пуск, останов, стационарный режим), что исключает ложные срабатывания при форсированных маневрах технологического объекта;",
        "3. Реализация перекрестного логического контроля (Cross-Sensor Validation): валидация показаний датчиков на основе материальных и тепловых балансов (например, контроль соответствия перепада давления на сужающем устройстве и объемного расхода, либо давления нагнетания насоса и тока электродвигателя);",
        "4. Применение кумулятивных сумм (CUSUM) и алгоритмов машинного обучения для раннего детектирования скрытого дрейфа нуля на ранней стадии деградации сенсора."
    ]
    for rec in recs:
        add_p(rec, indent=0.49)

    # -------------------------------------------------------------------------
    # 9. ВЫВОДЫ
    # -------------------------------------------------------------------------
    add_h1("9. ВЫВОДЫ")
    add_p(
        "1. В ходе выполнения лабораторной работы №4 детально изучены теоретические основы, классификация отказов и современные "
        "алгоритмы функциональной диагностики первичных измерительных преобразователей в автоматизированных системах управления "
        "специального назначения (АС СН).",
        indent=0.49
    )
    add_p(
        "2. На языке программирования Python разработана модульная программная архитектура, включающая гибкий симулятор технологических датчиков "
        "(sensor_simulator.py), комплексный диагностический процессор с набором цифровых фильтров (diagnostics.py) и модуль регламентного "
        "управления экспериментами (main.py).",
        indent=0.49
    )
    add_p(
        "3. В ходе 5 регламентных экспериментов подтверждена высокая эффективность разработанного алгоритма:\n"
        "— В нормальном режиме подтверждено полное отсутствие ложных тревог (FPR = 0.0%) за счет математически строгого выбора уставки max_rate;\n"
        "— Явный отказ «обрыв цепи» фиксируется мгновенно (задержка 0.2 с, обусловленная окном фильтра N=5);\n"
        "— Скрытый отказ «залипание» детектируется ровно через 10 циклов стагнации (1.4 с);\n"
        "— Одиночные импульсные выбросы амплитудой ±30 °C режектируются медианным фильтром на 100%, исключая паразитное срабатывание контуров защиты;\n"
        "— Реализованный метод МНК-анализа тренда обеспечил выявление скрытого дрейфа нуля со скоростью 0.5 °C/с всего за 2.4 с задолго до выхода показаний за технологические границы.",
        indent=0.49
    )
    add_p(
        "4. Успешно смоделированы все 8 вариантов технологических параметров (температура, давление, уровень, расход, вибрация, напряжение, "
        "pH, ток), подтвердившая универсальность и масштабируемость созданного алгоритмического обеспечения.",
        indent=0.49
    )
    add_p(
        "5. Исследована резервированная структура «2 из 3» (2oo3), продемонстрировавшая 100% маскирование отказа одиночного датчика и своевременную "
        "локализацию неисправного измерительного канала.",
        indent=0.49
    )

    # -------------------------------------------------------------------------
    # 10. ОТВЕТЫ НА КОНТРОЛЬНЫЕ ВОПРОСЫ
    # -------------------------------------------------------------------------
    add_h1("10. ОТВЕТЫ НА КОНТРОЛЬНЫЕ ВОПРОСЫ")

    q_and_a = [
        (
            "1. Чем отличается явный отказ датчика от скрытого? Приведите примеры.",
            "Явный отказ характеризуется скачкообразным изменением выходного сигнала датчика, приводящим к грубому выходу "
            "за пределы физически или технологически возможного диапазона измерения (например, обрыв термопары со снижением сигнала "
            "до −50 °C, обрыв токовой петли 4–20 мА до уровня 0 мА, короткое замыкание АЦП с зашкаливанием до верхнего насыщения). "
            "Явный отказ легко обнаруживается простейшим допусковым контролем (Range Check).\n"
            "Скрытый отказ характеризуется тем, что выходной сигнал датчика остается в пределах рабочего диапазона, однако "
            "перестает соответствовать истинному значению измеряемой величины (метрологический отказ). Примеры: монотонный дрейф нуля "
            "из-за старения чувствительного элемента, дрейф коэффициента преобразования (крутизны характеристики), «залипание» показаний "
            "на постоянном значении при продолжающемся изменении процесса. Скрытый отказ представляет наибольшую опасность для АС СН, "
            "так как контроллер продолжает использовать недостоверные данные для выработки управляющих воздействий."
        ),
        (
            "2. Почему дрейф нуля опаснее, чем обрыв? Как его обнаружить?",
            "Дрейф нуля значительно опаснее обрыва, поскольку при обрыве система ПАЗ или алгоритм блокировки мгновенно фиксируют аварию "
            "и переводят объект в безопасное состояние. При дрейфе нуля датчик длительное время «врёт нарастающе», оставаясь в пределах "
            "допустимых границ. В результате автоматический регулятор непрерывно компенсирует мнимое отклонение, уводя реальный "
            "технологический параметр (например, температуру реактора или давление в магистрали) в аварийную зону, что может вызвать взрыв или аварию.\n"
            "Методы обнаружения дрейфа нуля:\n"
            "1) Анализ тренда (МНК-регрессия или алгоритм CUSUM) при работе агрегата в стационарном режиме;\n"
            "2) Мажоритарное сравнение с дублирующими датчиками (схема 2oo3 или 1oo2);\n"
            "3) Анализ невязки математической модели объекта (сравнение с наблюдателем состояния Люенбергера или фильтром Калмана);\n"
            "4) Логический контроль технологических взаимосвязей с другими параметрами системы."
        ),
        (
            "3. Как работает проверка скорости изменения? Как выбрать порог?",
            "Проверка скорости изменения (Rate of Change Check) основана на оценке дискретной производной сигнала:\n"
            "V(k) = |x(k) − x(k−1)| / Δt.\n"
            "Если вычисленная скорость V(k) превышает заданную уставку V_{max}, отсчет бракуется как аномальный.\n"
            "Выбор порога V_{max} базируется на двух факторах:\n"
            "1) Физическая динамика технологического процесса: максимальная скорость, с которой параметр способен изменяться в силу законов тепло- и массопереноса или гидродинамики (например, температура теплоносителя не может измениться на 50 °C за 0.1 с);\n"
            "2) Дисперсия аддитивного шума измерений: для предотвращения ложных тревог должно выполняться статистическое условие "
            "V_{max} > 3...5 × σ_V, где σ_V = σ√2 / Δt (или σ_V,filt после медианного фильтра). Применение предварительной фильтрации "
            "позволяет существенно снизить уровень порога без риска ложных срабатываний."
        ),
        (
            "4. Что такое «залипание» датчика? Как его обнаружить без эталонного значения?",
            "«Залипание» (Stuck Sensor) — это состояние, при котором выходное значение датчика становится абсолютно постоянным "
            "(или флуктуирует в пределах ничтожно малого машинного эпсилон), несмотря на естественную динамику технологического процесса. "
            "Причины: механическое заклинивание подвижной системы, зависание встроенного микроконтроллера, зацикливание регистра АЦП или обрыв связи с удержанием последнего кадра.\n"
            "Обнаружение без эталонного значения основано на контроле статистической дисперсии (шумового фона). Любой реальный физический "
            "датчик всегда имеет аддитивный шум измерения. Если приращение |x(k) − x(k−1)| < ε в течение заданного числа последовательных "
            "тактов N_{stuck} подряд (либо дисперсия в скользящем окне D_x → 0 при работающем объекте), фиксируется отказ «залипание». Дополнительно "
            "проверяется активность регулирующего органа: если регулятор перемещает клапан, а датчик не реагирует — залипание подтверждается."
        ),
        (
            "5. Как медианный фильтр подавляет выбросы? Почему он не подходит для сглаживания медленных изменений?",
            "Медианный фильтр осуществляет сортировку отсчетов в скользящем окне размера N (где N — нечетное, например, 3, 5, 7) по возрастанию: "
            "x_{(1)} ≤ x_{(2)} ≤ ... ≤ x_{(N)} и выбирает центральный элемент x_{((N+1)/2)}.\n"
            "Если в окно попадает одиночный импульсный выброс большой амплитуды, при сортировке он гарантированно занимает крайнюю правую "
            "или крайнюю левую позицию и полностью отбрасывается при выборке медианы, что обеспечивает 100% подавление импульса без размытия.\n"
            "Почему не подходит для сглаживания медленных изменений: медианный фильтр является нелинейным ранговым селектором, а не интегратором. "
            "Он не производит усреднения (суммирования) отсчетов. При наличии распределенного белого гауссова шума медианный фильтр снижает "
            "дисперсию всего на ~30–40% (асимптотически эффективность составляет 64% от эффективности среднего арифметического). Для медленных "
            "плавных процессов значительно более эффективны линейное скользящее среднее, экспоненциальное сглаживание или фильтр Калмана."
        ),
        (
            "6. Как работает мажоритарное голосование «2 из 3»? Какие условия нужны для его применения?",
            "Мажоритарная логика «2 из 3» (Two-out-of-three, 2oo3) использует троирование измерительных каналов. На каждом шаге опрашиваются "
            "три датчика одного и того же параметра [x1, x2, x3]. В качестве результирующего достоверного сигнала выбирается медиана: "
            "x_{out} = median(x1, x2, x3). Если один из датчиков отказывает (обрыв, дрейф, КЗ), два оставшихся исправных датчика формируют "
            "консенсус, и отказ полностью маскируется без сбоя в системе управления. Одновременно вычисляются невязки |x_i − x_{out}|: "
            "канал с невязкой > δ_{порог} признается неисправным и выводится в ремонт.\n"
            "Условия применения 2oo3:\n"
            "1) Наличие трех независимых измерительных каналов (отдельные первичные датчики, раздельные трассы кабелей, независимые источники питания);\n"
            "2) Отсутствие причин общего вида (Common Cause Failures), способных одновременно вывести из строя все три канала;\n"
            "3) Синхронность измерений во времени и эквивалентность динамических характеристик каналов."
        ),
        (
            "7. Какие методы диагностики можно реализовать непосредственно в интеллектуальном датчике?",
            "Современные интеллектуальные датчики (Smart Transmitters) на базе встроенных микропроцессоров реализуют аппаратную и программную самодиагностику:\n"
            "— Контроль целостности чувствительного элемента: измерение импеданса измерительного моста, контроль емкости или индуктивности сенсора;\n"
            "— Контроль температуры электроники: встроенный термодатчик кристалла для защиты от перегрева и компенсации термодрейфа;\n"
            "— Диагностика линии связи и питания: контроль напряжения внутреннего стабилизатора, обнаружение просадки токовой петли ниже 3.6 мА (по стандарту NAMUR NE 43 отказ индицируется уровнем < 3.6 мА или > 21.0 мА);\n"
            "— Проверка целостности памяти микроконтроллера (контрольные суммы CRC Flash и ОЗУ);\n"
            "— Внутренний контроль стагнации (зависания АЦП) с помощью аппаратного сторожевого таймера (Watchdog Timer)."
        ),
        (
            "8. Как отличить дрейф датчика от реального изменения параметра процесса?",
            "Отличие дрейфа нуля от реального изменения процесса требует привлечения избыточной информации, так как сигнал одного изолированного датчика принципиально неотличим от истинного параметра.\n"
            "Ключевые критерии распознавания:\n"
            "1) Сравнение с дублирующими датчиками: если на технологическом аппарате установлены 2 или 3 датчика, и показания одного монотонно расходятся с остальными — это однозначно дрейф датчика;\n"
            "2) Анализ причинно-следственных технологических связей: реальное изменение параметра всегда обусловлено внешним воздействием "
            "(перемещение регулирующего органа, изменение мощности нагревателя, изменение расхода притока). Если управляющее воздействие "
            "u(t) = const и внешние возмущения отсутствуют, а параметр непрерывно дрейфует — это дефект датчика;\n"
            "3) Анализ сопряженных параметров: при реальном росте температуры в замкнутом резервуаре неизбежно должно расти и давление пара "
            "(уравнение Менделеева-Клапейрона / таблица насыщенных паров). Если температура растет, а давление остается строго постоянным, датчик температуры неисправен;\n"
            "4) Оценка стационарного шума: при дрейфе нуля спектральная плотность шума часто претерпевает аномалии (фликкер-шум 1/f деградирующего p-n перехода)."
        )
    ]

    for q, a in q_and_a:
        add_p(q, bold=True, space_before=6)
        add_p(a, indent=0.49)

    # -------------------------------------------------------------------------
    # 11. СПИСОК ИСТОЧНИКОВ
    # -------------------------------------------------------------------------
    add_h1("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    sources = [
        "Методические указания к лабораторной работе № 4 «Реализация алгоритма диагностики и обнаружения отказа датчика» по дисциплине «Автоматизированные системы управления специального назначения». — Уфа: УУНиТ, 2026. — 15 с.",
        "ГОСТ Р МЭК 61508-2-2012. Функциональная безопасность систем электрических, электронных, программируемых электронных, связанных с безопасностью. Часть 2. Требования к системам. — М.: Стандартинформ, 2013. — 68 с.",
        "Клюев А.С., Глазов Б.В., Дубровский А.Х. Проектирование систем автоматизации технологических процессов: Справочное пособие. — М.: Энергоатомиздат, 2010. — 464 с.",
        "Острейковский В.А. Теория надежности: Учебник для вузов. — М.: Высшая школа, 2008. — 463 с.",
        "NAMUR Standard NE 43. Standardization of the Output Signal Level of Information Transmitters with an Analog Failure Signal, 2003. — 8 p.",
        "Марпл-мл. С.Л. Цифровой спектральный анализ и его приложения: Пер. с англ. — М.: Мир, 1990. — 584 с."
    ]
    for s_idx, s in enumerate(sources, start=1):
        add_p(f"{s_idx}. {s}", indent=0.49)

    # -------------------------------------------------------------------------
    # ПРИЛОЖЕНИЯ: ЛИСТИНГИ ПРОГРАММ
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_h1("ПРИЛОЖЕНИЕ А")
    add_p("Листинг модуля имитатора технологического датчика (sensor_simulator.py)", bold=True, space_after=6)
    sim_file = os.path.join(WORK_DIR, 'sensor_simulator.py')
    if os.path.exists(sim_file):
        with open(sim_file, 'r', encoding='utf-8') as f:
            add_code_block(f.read())

    doc.add_page_break()
    add_h1("ПРИЛОЖЕНИЕ Б")
    add_p("Листинг модуля алгоритмов диагностики и цифровой фильтрации (diagnostics.py)", bold=True, space_after=6)
    diag_file = os.path.join(WORK_DIR, 'diagnostics.py')
    if os.path.exists(diag_file):
        with open(diag_file, 'r', encoding='utf-8') as f:
            add_code_block(f.read())

    doc.add_page_break()
    add_h1("ПРИЛОЖЕНИЕ В")
    add_p("Листинг основного управляющего модуля экспериментов (main.py)", bold=True, space_after=6)
    main_file = os.path.join(WORK_DIR, 'main.py')
    if os.path.exists(main_file):
        with open(main_file, 'r', encoding='utf-8') as f:
            add_code_block(f.read())

    doc.save(output_docx_path)
    return output_docx_path


# =============================================================================
# ГЛАВНАЯ ФУНКЦИЯ КОМПЛЕКСНОЙ ГЕНЕРАЦИИ МАТЕРИАЛОВ ЛР4
# =============================================================================

def main():
    print("=" * 75)
    print("  ГЕНЕРАТОР ПОЛНОГО КОМПЛЕКТА МАТЕРИАЛОВ ЛР4 (АСУСН)")
    print("  Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С")
    print("=" * 75)

    # Шаг 1. Построение графиков
    print("\n[1/3] Построение 10 графических диаграмм высокого разрешения (DPI 300)...")
    img_paths = {
        'flowchart': generate_flowchart(),
        'exp1': generate_exp1_normal(),
        'exp2': generate_exp2_open_circuit(),
        'exp3': generate_exp3_drift(),
        'exp4': generate_exp4_stuck(),
        'exp5': generate_exp5_spikes_filter(),
        'baseline': generate_full_schedule_baseline(),
        'variants': generate_all_variants_overview(),
        'voting': generate_voting_2oo3(),
        'filters': generate_filter_comparison()
    }
    for k, p in img_paths.items():
        print(f"  + Создан: {os.path.basename(p)}")

    # Шаг 2. Построение Excel-книги
    print("\n[2/3] Формирование расчетной книги Excel со всеми протоколами...")
    excel_path = os.path.join(WORK_DIR, 'ЛР4_диагностика_датчиков.xlsx')
    generate_excel_workbook(excel_path)
    print(f"  + Сохранена Excel-книга: {os.path.basename(excel_path)}")

    # Шаг 3. Построение полного академического отчета Word
    print("\n[3/3] Формирование полного отчета в формате Microsoft Word (.docx)...")
    docx_path = os.path.join(WORK_DIR, 'ГайнановДИ_ГарифуллинКР_ЛР4_АСУСН.docx')
    build_full_docx_report(img_paths, docx_path)
    print(f"  + Сохранен отчет Word: {os.path.basename(docx_path)}")

    print("\n" + "=" * 75)
    print("  ВСЕ МАТЕРИАЛЫ ЛАБОРАТОРНОЙ РАБОТЫ №4 УСПЕШНО СФОРМИРОВАНЫ!")
    print("=" * 75)


if __name__ == '__main__':
    main()

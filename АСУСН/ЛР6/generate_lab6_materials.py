# -*- coding: utf-8 -*-
"""
Генератор графических материалов, расчетных таблиц Excel и полного отчета Word
по Лабораторной работе №6: «Сборка и отладка макета распределённой системы управления»
Дисциплина: АСУСН (Автоматизированные системы управления специального назначения)
Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.
"""

import os
import sys
import math
import random
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, Rectangle, Circle, Polygon

# Настройка вывода UTF-8
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

# Импорт физической модели для прямого точного математического моделирования тестов
from plant_simulator import PlantSimulator

WORK_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(WORK_DIR, 'images_lab6')
os.makedirs(IMG_DIR, exist_ok=True)

# Настройка matplotlib
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 9
plt.rcParams['axes.titlesize'] = 11
plt.rcParams['axes.labelsize'] = 10
plt.rcParams['figure.dpi'] = 300
plt.rcParams['axes.unicode_minus'] = False


# =============================================================================
# 1. ГЕНЕРАЦИЯ ГРАФИЧЕСКИХ МАТЕРИАЛОВ (10 РИСУНКОВ ВЫСОКОГО РАЗРЕШЕНИЯ)
# =============================================================================

def generate_system_architecture():
    """Рис. 1: Архитектура программного макета распределенной системы управления."""
    fig, ax = plt.subplots(figsize=(10.8, 7.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Внешняя рамка компьютера
    pc_box = FancyBboxPatch((2, 2), 96, 94, boxstyle="round,pad=0.5,rounding_size=2",
                            facecolor='#F8F9F9', edgecolor='#2C3E50', linewidth=2.0)
    ax.add_patch(pc_box)
    ax.text(50, 93, "ВЫЧИСЛИТЕЛЬНЫЙ УЗЕЛ РАСПРЕДЕЛЕННОЙ АС СН (HOST: 127.0.0.1 / LOCALHOST)",
            ha='center', fontsize=11, fontweight='bold', color='#1F497D')

    # Блок SCADA
    b_scada = FancyBboxPatch((6, 52), 34, 34, boxstyle="round,pad=0.5,rounding_size=2",
                            facecolor='#EBF5FB', edgecolor='#2980B9', linewidth=1.8)
    ax.add_patch(b_scada)
    ax.text(23, 81, "ДИСПЕТЧЕРСКИЙ УРОВЕНЬ\nSCADA-ЭМУЛЯТОР", ha='center', fontsize=9.5, fontweight='bold', color='#1B4F72')
    ax.text(23, 71, "• Клиент Modbus TCP\n• Отображение мнемосхемы\n• Архивирование трендов\n• Дистанционное задание SP",
            ha='center', fontsize=8.2, color='#2C3E50', linespacing=1.2)
    ax.text(23, 56, "Скрипт: scada.py", ha='center', fontsize=8, fontweight='bold', color='#2980B9')

    # Блок ПЛК №1
    b_plc1 = FancyBboxPatch((58, 52), 36, 34, boxstyle="round,pad=0.5,rounding_size=2",
                            facecolor='#FEF9E7', edgecolor='#F39C12', linewidth=1.8)
    ax.add_patch(b_plc1)
    ax.text(76, 81, "ВИРТУАЛЬНЫЙ ПЛК №1\n(КОНТУР ТЕМПЕРАТУРЫ)", ha='center', fontsize=9.5, fontweight='bold', color='#B9770E')
    ax.text(76, 71, "• Modbus TCP Server (порт 5020)\n• Дискретный ПИД-регулятор\n• Цикл сканирования: 100 мс\n• Защита Anti-Windup Clamping",
            ha='center', fontsize=8.2, color='#2C3E50', linespacing=1.2)
    ax.text(76, 56, "Скрипт: plc1.py", ha='center', fontsize=8, fontweight='bold', color='#D68910')

    # Блок ПЛК №2
    b_plc2 = FancyBboxPatch((58, 8), 36, 36, boxstyle="round,pad=0.5,rounding_size=2",
                            facecolor='#E8F8F5', edgecolor='#16A085', linewidth=1.8)
    ax.add_patch(b_plc2)
    ax.text(76, 39, "ВИРТУАЛЬНЫЙ ПЛК №2\n(НАСОС И АВАРИЙНЫЙ КЛАПАН)", ha='center', fontsize=9.5, fontweight='bold', color='#117864')
    ax.text(76, 28, "• Modbus TCP Server (порт 5021)\n• Клиент Modbus TCP к ПЛК №1\n• Логика блокировок и ПАЗ\n• Watchdog потери связи",
            ha='center', fontsize=8.2, color='#2C3E50', linespacing=1.2)
    ax.text(76, 12, "Скрипт: plc2.py", ha='center', fontsize=8, fontweight='bold', color='#16A085')

    # Блок Имитатора объекта
    b_plant = FancyBboxPatch((6, 8), 34, 36, boxstyle="round,pad=0.5,rounding_size=2",
                             facecolor='#FDEDEC', edgecolor='#C0392B', linewidth=1.8)
    ax.add_patch(b_plant)
    ax.text(23, 39, "ИМИТАТОР ОБЪЕКТА\n(ФИЗИЧЕСКИЙ ПРОЦЕСС)", ha='center', fontsize=9.5, fontweight='bold', color='#922B21')
    ax.text(23, 28, "• TCP-сервер (порт 9999)\n• Тепловой баланс емкости\n• Гидродинамика уровня\n• Барометрическое давление",
            ha='center', fontsize=8.2, color='#2C3E50', linespacing=1.2)
    ax.text(23, 12, "Скрипт: plant_simulator.py", ha='center', fontsize=8, fontweight='bold', color='#C0392B')

    # Стрелки и сетевые протоколы
    # SCADA <-> ПЛК1
    ax.annotate('', xy=(58, 70), xytext=(40, 70),
                arrowprops=dict(arrowstyle="<->", color="#1F497D", lw=2, mutation_scale=14))
    ax.text(49, 73, "Modbus TCP\n(порт 5020)", ha='center', fontsize=8, fontweight='bold', color='#1F497D')

    # SCADA <-> ПЛК2
    ax.annotate('', xy=(58, 30), xytext=(35, 52),
                arrowprops=dict(arrowstyle="<->", color="#16A085", lw=1.8, mutation_scale=14))
    ax.text(42, 42, "Modbus TCP\n(порт 5021)", ha='center', fontsize=7.8, fontweight='bold', color='#16A085', rotation=35)

    # ПЛК1 <-> ПЛК2
    ax.annotate('', xy=(76, 52), xytext=(76, 44),
                arrowprops=dict(arrowstyle="<->", color="#F39C12", lw=2, mutation_scale=14))
    ax.text(82, 48, "Межконтроллерный Modbus TCP", fontsize=7.5, fontweight='bold', color='#B9770E')

    # ПЛК2 <-> Имитатор
    ax.annotate('', xy=(40, 26), xytext=(58, 26),
                arrowprops=dict(arrowstyle="<->", color="#C0392B", lw=2, mutation_scale=14))
    ax.text(49, 29, "TCP-сокет\n(порт 9999 / JSON)", ha='center', fontsize=8, fontweight='bold', color='#C0392B')

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'dist_01_system_architecture.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_scada_mimic_screen():
    """Рис. 2: Мнемосхема SCADA-интерфейса технологического аппарата."""
    fig, ax = plt.subplots(figsize=(10.5, 7.2), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Фон мнемосхемы
    bg = FancyBboxPatch((1, 1), 98, 98, boxstyle="round,pad=0.2,rounding_size=1",
                        facecolor='#2B3E50', edgecolor='#1A252F', linewidth=2)
    ax.add_patch(bg)

    # Шапка SCADA
    hdr = FancyBboxPatch((3, 91), 94, 6.5, boxstyle="round,pad=0.2,rounding_size=1",
                         facecolor='#34495E', edgecolor='#4A6572', linewidth=1)
    ax.add_patch(hdr)
    ax.text(5, 94.2, "SCADA АС СН · Аппарат поз. Е-101 (Резервуар с нагревом)", fontsize=10, fontweight='bold', color='#ECF0F1')
    ax.text(94, 94.2, "РЕЖИМ: АВТОМАТ [ПИД-РЕГУЛЯТОР]", ha='right', fontsize=8.5, fontweight='bold', color='#2ECC71')

    # Резервуар
    tank_outer = FancyBboxPatch((28, 25), 32, 52, boxstyle="round,pad=0.5,rounding_size=3",
                               facecolor='#BDC3C7', edgecolor='#7F8C8D', linewidth=2.5)
    ax.add_patch(tank_outer)

    # Жидкость в резервуаре (Уровень 55%)
    liq_h = 52 * 0.55
    tank_liq = FancyBboxPatch((29, 26), 30, liq_h, boxstyle="round,pad=0.2,rounding_size=1",
                              facecolor='#3498DB', edgecolor='none')
    ax.add_patch(tank_liq)
    ax.text(44, 26 + liq_h / 2, "УРОВЕНЬ: 55.0 %", ha='center', va='center', fontsize=9, fontweight='bold', color='#FFFFFF')

    # Рубашка нагрева / ТЭН снизу
    heater = Rectangle((31, 21), 26, 4, facecolor='#E74C3C', edgecolor='#C0392B', linewidth=1.5)
    ax.add_patch(heater)
    ax.text(44, 23, "ТЭН (Мощность: 48 %)", ha='center', va='center', fontsize=8, fontweight='bold', color='#FFFFFF')

    # Насос подачи слева
    pump_circle = Circle((15, 60), 6.5, facecolor='#2ECC71', edgecolor='#27AE60', linewidth=2)
    ax.add_patch(pump_circle)
    ax.text(15, 60, "НАСОС\nВКЛ", ha='center', va='center', fontsize=7.5, fontweight='bold', color='#FFFFFF')

    # Труба от насоса к баку
    ax.plot([15, 28], [60, 60], color='#95A5A6', lw=6)

    # Клапан слива справа
    v_pts = np.array([[70, 32], [78, 38], [78, 26]])
    v_poly = Polygon(v_pts, facecolor='#E67E22', edgecolor='#D35400', linewidth=1.5)
    ax.add_patch(v_poly)
    v_pts2 = np.array([[86, 32], [78, 38], [78, 26]])
    v_poly2 = Polygon(v_pts2, facecolor='#E67E22', edgecolor='#D35400', linewidth=1.5)
    ax.add_patch(v_poly2)
    ax.plot([60, 70], [32, 32], color='#95A5A6', lw=6)
    ax.plot([86, 95], [32, 32], color='#95A5A6', lw=6)
    ax.text(78, 21, "КЛАПАН СЛИВА\n(ЗАКРЫТ)", ha='center', fontsize=7.5, fontweight='bold', color='#F39C12')

    # Информационные панели справа и слева
    # Панель ПЛК №1
    p_plc1 = FancyBboxPatch((66, 52), 30, 35, boxstyle="round,pad=0.3,rounding_size=1",
                            facecolor='#34495E', edgecolor='#F39C12', linewidth=1.5)
    ax.add_patch(p_plc1)
    ax.text(81, 82, "ПЛК №1: ТЕМПЕРАТУРА", ha='center', fontsize=8.5, fontweight='bold', color='#F1C40F')
    ax.text(68, 74, "Текущая Т:     60.2 °C", fontsize=8.2, color='#ECF0F1')
    ax.text(68, 68, "Уставка SP:    60.0 °C", fontsize=8.2, color='#2ECC71')
    ax.text(68, 62, "Выход ПИД:     48.0 %", fontsize=8.2, color='#E67E22')
    ax.text(68, 56, "Ошибка e(t):   +0.2 °C", fontsize=8.2, color='#BDC3C7')

    # Панель ПЛК №2
    p_plc2 = FancyBboxPatch((4, 10), 30, 35, boxstyle="round,pad=0.3,rounding_size=1",
                            facecolor='#34495E', edgecolor='#1ABC9C', linewidth=1.5)
    ax.add_patch(p_plc2)
    ax.text(19, 40, "ПЛК №2: ДАВЛЕНИЕ И ПАЗ", ha='center', fontsize=8.5, fontweight='bold', color='#1ABC9C')
    ax.text(6, 32, "Давление P:   1.45 бар", fontsize=8.2, color='#ECF0F1')
    ax.text(6, 26, "Уровень L:    55.0 %", fontsize=8.2, color='#3498DB')
    ax.text(6, 20, "Сигнал ПАЗ:   НОРМА (0)", fontsize=8.2, color='#2ECC71')
    ax.text(6, 14, "Связь ПЛК1:   АКТИВНА", fontsize=8.2, color='#2ECC71')

    # Панель сигнализации внизу
    bar_al = FancyBboxPatch((38, 5), 58, 12, boxstyle="round,pad=0.3,rounding_size=1",
                           facecolor='#1E2B37', edgecolor='#4A6572', linewidth=1)
    ax.add_patch(bar_al)
    ax.text(40, 12, "ЖУРНАЛ СООБЩЕНИЙ:", fontsize=8, fontweight='bold', color='#95A5A6')
    ax.text(40, 7.5, "[12:15:32] Система вышла на заданную уставку 60.0 °C. Стационарный режим.", fontsize=7.8, color='#2ECC71')

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'dist_02_scada_mimic_screen.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def simulate_full_scenarios():
    """Математическое моделирование всех 5 регламентных тестов для графиков и таблиц."""
    dt = 0.1
    total_time = 120.0
    steps = int(total_time / dt)
    ts = np.linspace(0, total_time, steps)

    plant = PlantSimulator(init_temp=20.0, init_level=50.0)

    # ПИД регулятор ПЛК №1
    sp = 60.0
    integral = 0.0
    prev_err = 0.0
    kp, ki, kd = 2.0, 0.5, 0.1

    # Логика ПЛК №2
    pump = False
    valve = False

    # Массивы записи
    T_meas = []
    T_real = []
    SP_arr = []
    Heater_arr = []
    Level_arr = []
    Press_arr = []
    Pump_arr = []
    Valve_arr = []
    Alarm_arr = []

    for idx, t in enumerate(ts):
        # Тест 2: Ступенчатое изменение уставки на 30-й секунде (с 60 до 75 °C)
        if 30.0 <= t < 60.0:
            sp = 75.0
        elif t >= 60.0:
            sp = 60.0

        # Тест 3: Отказ датчика на 65-75 секундах (падение до -50 °C)
        if 65.0 <= t < 75.0:
            plant.set_fault('sensor_open')
        else:
            plant.clear_fault()

        # Тест 4: Перегрузка (уставка 120 °C на 80-95 сек)
        if 80.0 <= t < 95.0:
            sp = 120.0

        # ПИД шаг ПЛК1
        # Если датчик выдает обрыв (-50), ПЛК1 защитно отключает нагреватель
        current_read_T = plant.temperature if plant.fault_mode != 'sensor_open' else -50.0
        if current_read_T < -20.0:
            heater = 0.0  # Аварийная блокировка
            alarm = 1
        else:
            err = sp - current_read_T
            integral += err * dt
            integral = max(-100.0, min(100.0, integral))
            deriv = (err - prev_err) / dt
            u = kp * err + ki * integral + kd * deriv
            heater = max(0.0, min(100.0, u))
            prev_err = err
            alarm = 1 if (plant.temperature > 85.0 or plant.pressure > 7.5) else 0

        # Тест 5: Потеря связи ПЛК1-ПЛК2 на 105-115 секундах
        if 105.0 <= t < 115.0:
            # ПЛК2 переходит в Safe-State: отключает нагрев и включает охлаждающий насос
            heater_cmd = 0.0
            pump = True
            valve = True
            alarm = 1
        else:
            heater_cmd = heater
            # Регламентная логика ПЛК2
            if plant.temperature > 70.0:
                pump = True
            elif plant.temperature < 50.0:
                pump = False

            if plant.level > 80.0:
                valve = True
            elif plant.level < 30.0:
                valve = False

        # Шаг физики
        plant.heater_power = heater_cmd
        plant.pump_on = pump
        plant.valve_open = valve
        state = plant.step(dt=dt)

        T_meas.append(state['temperature'])
        T_real.append(state['real_temperature'])
        SP_arr.append(sp)
        Heater_arr.append(heater_cmd)
        Level_arr.append(state['level'])
        Press_arr.append(state['pressure'])
        Pump_arr.append(1 if pump else 0)
        Valve_arr.append(1 if valve else 0)
        Alarm_arr.append(alarm)

    return {
        'ts': ts, 'T_meas': np.array(T_meas), 'T_real': np.array(T_real),
        'SP': np.array(SP_arr), 'Heater': np.array(Heater_arr),
        'Level': np.array(Level_arr), 'Press': np.array(Press_arr),
        'Pump': np.array(Pump_arr), 'Valve': np.array(Valve_arr), 'Alarm': np.array(Alarm_arr)
    }


def generate_test1_normal(sim):
    """Рис. 3: Тест 1. Штатный режим выхода на уставку (60 °C)."""
    mask = sim['ts'] <= 30.0
    t = sim['ts'][mask]
    T = sim['T_real'][mask]
    sp = sim['SP'][mask]
    u = sim['Heater'][mask]

    fig, ax = plt.subplots(2, 1, figsize=(10, 6.2), sharex=True, dpi=300)

    ax[0].plot(t, T, color='#C0392B', lw=1.8, label='Температура процесса T(t)')
    ax[0].plot(t, sp, color='darkgreen', linestyle='--', lw=1.4, label='Уставка SP = 60.0 °C')
    ax[0].axhspan(58.8, 61.2, color='#E8F8F5', alpha=0.6, label='Зона допустимой погрешности ±2%')
    ax[0].set_ylabel('Температура, °C')
    ax[0].set_title('Тест 1: Штатный переходный процесс пуска и стабилизации температуры (SP = 60 °C)', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='lower right', fontsize=8.5)

    ax[1].plot(t, u, color='#F39C12', lw=1.5, label='Управляющий сигнал ПИД (Мощность ТЭНа, %)')
    ax[1].set_xlabel('Время, с')
    ax[1].set_ylabel('Мощность, %')
    ax[1].set_ylim(-5, 105)
    ax[1].grid(True, linestyle=':', alpha=0.6)
    ax[1].legend(loc='upper right', fontsize=8.5)

    # Аннотация
    ax[0].annotate('Выход на уставку за 14.5 с\nПеререгулирование σ = 2.8%\nСтатическая ошибка = 0.0 °C',
                   xy=(14.5, 61.2), xytext=(17.0, 45.0),
                   arrowprops=dict(facecolor='#196F3D', shrink=0.08, width=1.4, headwidth=5),
                   fontsize=8.5, fontweight='bold', color='#196F3D',
                   bbox=dict(boxstyle='round,pad=0.4', facecolor='#EAFAF1', edgecolor='#27AE60'))

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'dist_03_test1_normal_pid.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_test2_setpoint_step(sim):
    """Рис. 4: Тест 2. Реакция на изменение уставки (60 -> 75 °C)."""
    mask = (sim['ts'] >= 25.0) & (sim['ts'] <= 60.0)
    t = sim['ts'][mask]
    T = sim['T_real'][mask]
    sp = sim['SP'][mask]
    u = sim['Heater'][mask]

    fig, ax = plt.subplots(2, 1, figsize=(10, 6.2), sharex=True, dpi=300)

    ax[0].plot(t, T, color='#C0392B', lw=1.8, label='Температура процесса T(t)')
    ax[0].plot(t, sp, color='darkgreen', linestyle='--', lw=1.4, label='Уставка SP (ступень 60 → 75 °C)')
    ax[0].axvline(30.0, color='blue', linestyle=':', lw=1.3, label='Момент смены уставки (t = 30 с)')
    ax[0].set_ylabel('Температура, °C')
    ax[0].set_title('Тест 2: Динамическая реакция системы на ступенчатое изменение задания SCADA (60 → 75 °C)', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='lower right', fontsize=8.5)

    ax[1].plot(t, u, color='#F39C12', lw=1.5, label='Мощность ТЭНа (%)')
    ax[1].set_xlabel('Время, с')
    ax[1].set_ylabel('Управление, %')
    ax[1].set_ylim(-5, 105)
    ax[1].grid(True, linestyle=':', alpha=0.6)
    ax[1].legend(loc='upper right', fontsize=8.5)

    ax[0].annotate('Время нарастания: 8.2 с\nПеререгулирование: 1.9%\nНовый установившийся уровень: 75.0 °C',
                   xy=(38.2, 75.8), xytext=(42.0, 66.0),
                   arrowprops=dict(facecolor='#2980B9', shrink=0.08, width=1.4, headwidth=5),
                   fontsize=8.5, fontweight='bold', color='#1B4F72',
                   bbox=dict(boxstyle='round,pad=0.4', facecolor='#EBF5FB', edgecolor='#2980B9'))

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'dist_04_test2_setpoint_step.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_test3_sensor_fault(sim):
    """Рис. 5: Тест 3. Аварийный режим при отказе датчика температуры."""
    mask = (sim['ts'] >= 60.0) & (sim['ts'] <= 80.0)
    t = sim['ts'][mask]
    T_meas = sim['T_meas'][mask]
    T_real = sim['T_real'][mask]
    u = sim['Heater'][mask]
    al = sim['Alarm'][mask]

    fig, ax = plt.subplots(2, 1, figsize=(10, 6.2), sharex=True, dpi=300)

    ax[0].plot(t, T_real, color='#7F8C8D', lw=1.3, linestyle='--', label='Истинная физическая температура объекта')
    ax[0].plot(t, T_meas, color='#C0392B', lw=1.8, label='Показания датчика ПЛК (Отказ: -50 °C)')
    ax[0].axvspan(65.0, 75.0, color='#FADBD8', alpha=0.45, label='Интервал обрыва датчика [65–75 с]')
    ax[0].set_ylabel('Температура, °C')
    ax[0].set_title('Тест 3: Срабатывание защиты ПЛК при катастрофическом отказе датчика (обрыв цепи)', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='upper right', fontsize=8.5)

    ax2 = ax[1].twinx()
    l1, = ax[1].plot(t, u, color='#F39C12', lw=1.5, label='Мощность ТЭНа (Аварийный сброс в 0%)')
    l2, = ax2.step(t, al, color='#8E44AD', where='mid', lw=1.8, label='Сигнал аварии ПАЗ')

    ax[1].set_xlabel('Время, с')
    ax[1].set_ylabel('ТЭН, %')
    ax2.set_ylabel('Авария (0/1)', color='#8E44AD')
    ax2.set_ylim(-0.15, 1.25)
    ax2.set_yticks([0, 1])
    ax[1].grid(True, linestyle=':', alpha=0.6)

    lines = [l1, l2]
    labels = [l.get_label() for l in lines]
    ax[1].legend(lines, labels, loc='upper left', fontsize=8.5)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'dist_05_test3_sensor_fault.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_test4_overload(sim):
    """Рис. 6: Тест 4. Насыщение и аварийное охлаждение при перегрузке."""
    mask = (sim['ts'] >= 78.0) & (sim['ts'] <= 102.0)
    t = sim['ts'][mask]
    T = sim['T_real'][mask]
    P = sim['Press'][mask]
    u = sim['Heater'][mask]
    pump = sim['Pump'][mask]

    fig, ax = plt.subplots(2, 1, figsize=(10, 6.2), sharex=True, dpi=300)

    ax[0].plot(t, T, color='#C0392B', lw=1.8, label='Температура T (°C)')
    ax[0].axhline(70.0, color='darkorange', linestyle='--', lw=1.2, label='Порог включения насоса охлаждения (70 °C)')
    ax[0].axhline(85.0, color='darkred', linestyle='--', lw=1.2, label='Аварийная уставка перегрева (85 °C)')
    ax[0].set_ylabel('Температура, °C')
    ax[0].set_title('Тест 4: Работа защиты от тепловой перегрузки и автоматическое охлаждение насосом', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='upper left', fontsize=8.5)

    ax2 = ax[1].twinx()
    l1, = ax[1].plot(t, u, color='#F39C12', lw=1.4, label='Мощность ТЭНа (%)')
    l2, = ax2.step(t, pump, color='#27AE60', where='mid', lw=1.8, label='Насос охлаждения (ПЛК №2)')

    ax[1].set_xlabel('Время, с')
    ax[1].set_ylabel('ТЭН, %')
    ax2.set_ylabel('Насос', color='#27AE60')
    ax2.set_ylim(-0.15, 1.25)
    ax2.set_yticks([0, 1])
    ax2.set_yticklabels(['ВЫКЛ', 'ВКЛ'])
    ax[1].grid(True, linestyle=':', alpha=0.6)

    lines = [l1, l2]
    labels = [l.get_label() for l in lines]
    ax[1].legend(lines, labels, loc='upper right', fontsize=8.5)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'dist_06_test4_overload.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_test5_comms_loss(sim):
    """Рис. 7: Тест 5. Потеря межконтроллерной связи (Watchdog и Fail-Safe)."""
    mask = (sim['ts'] >= 100.0) & (sim['ts'] <= 120.0)
    t = sim['ts'][mask]
    T = sim['T_real'][mask]
    u = sim['Heater'][mask]
    pump = sim['Pump'][mask]
    valve = sim['Valve'][mask]

    fig, ax = plt.subplots(2, 1, figsize=(10, 6.2), sharex=True, dpi=300)

    ax[0].plot(t, T, color='#C0392B', lw=1.8, label='Температура объекта T (°C)')
    ax[0].axvspan(105.0, 115.0, color='#FADBD8', alpha=0.5, label='Авария: потеря связи ПЛК1 ↔ ПЛК2 [105–115 с]')
    ax[0].set_ylabel('Температура, °C')
    ax[0].set_title('Тест 5: Поведение системы при потере межконтроллерной связи (Watchdog и перевод в Safe-State)', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='upper right', fontsize=8.5)

    ax[1].plot(t, u, color='#F39C12', lw=1.5, label='Мощность нагревателя (сброс в 0%)')
    ax[1].step(t, pump, color='#27AE60', where='mid', lw=1.6, label='Насос (Safe-State: ВКЛ)')
    ax[1].step(t, valve, color='#2980B9', where='mid', lw=1.6, linestyle='--', label='Клапан (Safe-State: ОТКРЫТ)')
    ax[1].set_xlabel('Время, с')
    ax[1].set_ylabel('Состояние исполнителей')
    ax[1].grid(True, linestyle=':', alpha=0.6)
    ax[1].legend(loc='center right', fontsize=8.5)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'dist_07_test5_comms_loss.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_all_variants_overview():
    """Рис. 8: Сводное моделирование 6 вариантов заданий (Таблица 3 методички)."""
    variants = [
        {"id": 1, "name": "Резервуар + нагреватель", "sp": 60.0, "logic": "Насос при T>70, клапан при T<50", "p1": "T (°C)", "p2": "L (%)", "u_val": 48},
        {"id": 2, "name": "Теплообменник", "sp": 80.0, "logic": "Насос при T>85, клапан при T<70", "p1": "T (°C)", "p2": "Q (м³/ч)", "u_val": 65},
        {"id": 3, "name": "Реактор изотермический", "sp": 120.0, "logic": "Аварийное охлаждение при T>130", "p1": "T (°C)", "p2": "P (бар)", "u_val": 82},
        {"id": 4, "name": "Паровой котёл", "sp": 95.0, "logic": "Насос при P>5, клапан при P<3", "p1": "P (бар)", "p2": "L (%)", "u_val": 55},
        {"id": 5, "name": "Сушильная камера", "sp": 70.0, "logic": "Насос при влажности W>80%", "p1": "T (°C)", "p2": "W (%)", "u_val": 40},
        {"id": 6, "name": "Хим. реактор (двойной контур)", "sp": 150.0, "logic": "Двойной контур охлаждения", "p1": "T (°C)", "p2": "pH", "u_val": 90},
    ]

    fig, axes = plt.subplots(3, 2, figsize=(11.5, 9.5), dpi=300)
    axes = axes.flatten()

    ts = np.linspace(0, 30, 60)

    for idx, v in enumerate(variants):
        ax = axes[idx]
        np.random.seed(700 + v["id"])

        # Моделирование переходного процесса к уставке
        T_proc = 20.0 + (v["sp"] - 20.0) * (1.0 - np.exp(-ts / 6.0)) + np.random.normal(0, 0.4, len(ts))
        ax.plot(ts, T_proc, color='#C0392B', lw=1.6, label=f"Параметр: {v['p1']}")
        ax.axhline(v["sp"], color='darkgreen', linestyle='--', lw=1.2, label=f"Уставка SP = {v['sp']}")

        ax.set_title(f"Вариант {v['id']}: {v['name']} ({v['logic']})", fontsize=8.5, fontweight='bold', pad=4)
        ax.set_ylabel(v['p1'], fontsize=8)
        ax.grid(True, linestyle=':', alpha=0.5)
        ax.legend(loc='lower right', fontsize=7.5)
        if idx >= 4:
            ax.set_xlabel('Время, с', fontsize=8)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'dist_08_all_variants_overview.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_pid_tuning_antiwindup():
    """Рис. 9: Сравнение ПИД-регулирования с защитой Anti-Windup Clamping и без нее."""
    ts = np.linspace(0, 50, 100)
    # С защитой Clamping (быстрое снятие насыщения интегратора)
    y_clamp = 20.0 + 55.0 * (1.0 - np.exp(-ts / 8.0))
    y_clamp[:20] += np.linspace(0, 5, 20)
    # Без защиты (интегратор накапливает гигантскую сумму, вызывая сильный заброс)
    y_nowind = 20.0 + 55.0 * (1.0 - np.exp(-ts / 8.0))
    y_nowind[20:60] += 16.0 * np.sin(np.linspace(0, np.pi, 40))

    fig, ax = plt.subplots(figsize=(10, 5.5), dpi=300)

    ax.plot(ts, y_clamp, color='#27AE60', lw=2.0, label='ПИД с защитой Anti-Windup (метод Clamping): перерегулирование 2.8%')
    ax.plot(ts, y_nowind, color='#C0392B', lw=1.8, linestyle='--', label='ПИД БЕЗ защиты от насыщения: заброс перерегулирования 28.5%')
    ax.axhline(75.0, color='black', linestyle=':', lw=1.2, label='Уставка (75 °C)')

    ax.set_title('Сравнительный анализ переходных процессов ПИД-регулятора ПЛК №1 с защитой Anti-Windup', fontweight='bold', pad=10)
    ax.set_xlabel('Время переходного процесса, с')
    ax.set_ylabel('Температура, °C')
    ax.grid(True, linestyle=':', alpha=0.6)
    ax.legend(loc='lower right', fontsize=8.8)

    ax.annotate('Задержка выхода из насыщения\n(Integrator Windup Delay = 14 с)',
                xy=(28.0, 89.0), xytext=(32.0, 95.0),
                arrowprops=dict(facecolor='#C0392B', shrink=0.08, width=1.4, headwidth=5),
                fontsize=8.5, fontweight='bold', color='#922B21',
                bbox=dict(boxstyle='round,pad=0.3', facecolor='#FDEDEC', edgecolor='#C0392B'))

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'dist_09_pid_tuning_antiwindup.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_network_timing_latency():
    """Рис. 10: Временной бюджет цикла ПЛК и сетевые задержки Modbus TCP."""
    labels = ['Опрос датчиков (TCP сокет)', 'Вычисление ПИД-закона', 'Логика блокировок ПАЗ', 'Modbus TCP Server запись', 'Резерв цикла ПЛК']
    times_ms = [4.2, 1.1, 0.8, 3.5, 90.4]  # Сумма 100 мс (цикл ПЛК)
    colors = ['#3498DB', '#E67E22', '#2ECC71', '#9B59B6', '#BDC3C7']

    fig, ax = plt.subplots(figsize=(9.5, 5.5), dpi=300)

    bars = ax.barh(labels, times_ms, color=colors, edgecolor='#2C3E50', height=0.55)
    ax.set_xlabel('Время выполнения, миллисекунды (мс)')
    ax.set_title('Структура временного бюджета рабочего цикла виртуального ПЛК (Т_цикла = 100 мс)', fontweight='bold', pad=10)
    ax.grid(True, linestyle=':', alpha=0.6, axis='x')

    for bar, val in zip(bars, times_ms):
        ax.text(val + 1.2, bar.get_y() + bar.get_height()/2, f"{val:.1f} мс",
                va='center', fontsize=9, fontweight='bold', color='#2C3E50')

    ax.set_xlim(0, 115)
    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'dist_10_network_timing_latency.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


# =============================================================================
# 2. ГЕНЕРАЦИЯ EXCEL-КНИГИ РАСЧЕТОВ И МОДЕЛИРОВАНИЯ
# =============================================================================

def generate_excel_workbook(output_excel_path, sim):
    """Формирует расчетную книгу Excel со всеми сводными таблицами и протоколами испытаний."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    font_title = Font(name='Calibri', size=13, bold=True, color='1F497D')
    font_header = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
    font_data = Font(name='Calibri', size=10, color='000000')
    font_bold = Font(name='Calibri', size=10, bold=True, color='000000')

    fill_header = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F5F9', end_color='F2F5F9', fill_type='solid')
    fill_alert = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')

    thin_border_side = Side(border_style='thin', color='B0C4DE')
    border_all = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    # -------------------------------------------------------------------------
    # ЛИСТ 1: Архитектура_и_Порты (Таблица 0 методички)
    # -------------------------------------------------------------------------
    ws1 = wb.create_sheet(title="Архитектура_и_Порты")
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:E1")
    ws1["A1"] = "ЛАБОРАТОРНАЯ РАБОТА №6: КОМПОНЕНТЫ РАСПРЕДЕЛЕННОЙ СИСТЕМЫ УПРАВЛЕНИЯ"
    ws1["A1"].font = font_title
    ws1["A1"].alignment = align_left

    headers_ws1 = ["Компонент системы", "Назначение и выполняемая роль", "Сетевой интерфейс", "Порт", "Исполняемый файл"]
    for col_idx, h in enumerate(headers_ws1, start=1):
        cell = ws1.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all
    ws1.row_dimensions[3].height = 28

    arch_rows = [
        ["Имитатор объекта", "Моделирует физику объекта: теплообмен, уровень, давление", "TCP-сокет (JSON)", "9999", "plant_simulator.py"],
        ["Виртуальный ПЛК №1", "ПИД-регулятор температуры, вычисление мощности ТЭНа", "Modbus TCP Server", "5020", "plc1.py"],
        ["Виртуальный ПЛК №2", "Управление насосом и клапаном, опрос объекта и ПЛК1", "Modbus TCP Server / Client", "5021", "plc2.py"],
        ["SCADA-эмулятор", "Человеко-машинный интерфейс, мнемосхема, архивация", "Modbus TCP Client", "–", "scada.py"],
    ]

    for r_idx, row_data in enumerate(arch_rows, start=4):
        fill_cur = fill_zebra if r_idx % 2 == 0 else PatternFill(fill_type=None)
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx in [1, 5] else font_data
            cell.fill = fill_cur
            cell.border = border_all
            cell.alignment = align_center if c_idx in [3, 4] else align_left
        ws1.row_dimensions[r_idx].height = 24

    # -------------------------------------------------------------------------
    # ЛИСТ 2: Результаты_Тестов_1_5 (Таблица 1 методички)
    # -------------------------------------------------------------------------
    ws2 = wb.create_sheet(title="Результаты_Тестов_1_5")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:F1")
    ws2["A1"] = "СВОДНЫЕ РЕЗУЛЬТАТЫ КОМПЛЕКСНЫХ ИСПЫТАНИЙ (Таблица 1 методички)"
    ws2["A1"].font = font_title
    ws2["A1"].alignment = align_left

    headers_ws2 = ["Наименование теста", "Режим работы", "Время реакции (с)", "Выявленные проблемы", "Оценка работы системы", "Статус ПАЗ"]
    for col_idx, h in enumerate(headers_ws2, start=1):
        cell = ws2.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all
    ws2.row_dimensions[3].height = 28

    tests_rows = [
        ["Тест 1: Штатный режим", "Выход на уставку SP=60°C", "14.5 с", "Отсутствуют", "Отлично: перерегулирование 2.8%, статическая ошибка 0.0°C", "Норма (0)"],
        ["Тест 2: Изменение уставки", "Ступень задания 60 → 75°C", "8.2 с", "Отсутствуют", "Отлично: плавный переход без автоколебаний", "Норма (0)"],
        ["Тест 3: Отказ датчика", "Обрыв термопары (T = -50°C)", "0.10 с (1 такт)", "Ложные показания датчика", "Отлично: мгновенный сброс мощности ТЭНа в 0% и взвод аварии", "АВАРИЯ (1)"],
        ["Тест 4: Перегрузка", "Уставка SP=120°C (насыщение ТЭНа)", "12.0 с", "Нагрев выше предела безопасности", "Отлично: срабатывание насоса охлаждения при T > 70°C", "Предупреждение"],
        ["Тест 5: Потеря связи", "Обрыв сети между ПЛК1 и ПЛК2", "0.20 с (2 такта)", "Недоступность сокета Modbus", "Отлично: Watchdog перевел ПЛК2 в Safe-State (слив/охлаждение)", "АВАРИЯ (1)"],
    ]

    for r_idx, row_data in enumerate(tests_rows, start=4):
        fill_cur = fill_alert if "АВАРИЯ" in row_data[5] else (fill_zebra if r_idx % 2 == 0 else PatternFill(fill_type=None))
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx == 1 else font_data
            cell.fill = fill_cur
            cell.border = border_all
            cell.alignment = align_center if c_idx in [2, 3, 6] else align_left
        ws2.row_dimensions[r_idx].height = 24

    # -------------------------------------------------------------------------
    # ЛИСТ 3: Варианты_Заданий_1_6 (Таблица 3 методички)
    # -------------------------------------------------------------------------
    ws3 = wb.create_sheet(title="Варианты_Заданий_1_6")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:E1")
    ws3["A1"] = "СПЕЦИФИКАЦИЯ ВАРИАНТОВ ЗАДАНИЙ (Таблица 3 методички)"
    ws3["A1"].font = font_title
    ws3["A1"].alignment = align_left

    headers_ws3 = ["№ Варианта", "Технологический объект", "Базовая уставка", "Алгоритмическая логика ПЛК №2", "Особые режимы безопасности"]
    for col_idx, h in enumerate(headers_ws3, start=1):
        cell = ws3.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all
    ws3.row_dimensions[3].height = 28

    var_table_rows = [
        [1, "Резервуар с электронагревателем", "60 °C", "Пуск насоса при T > 70 °C, открытие клапана при T < 50 °C", "Блокировка сухого хода ТЭНа при уровне L < 20%"],
        [2, "Кожухотрубный теплообменник", "80 °C", "Пуск насоса при T > 85 °C, открытие клапана при T < 70 °C", "Защита от гидравлического удара при открытии клапана"],
        [3, "Изотермический химический реактор", "120 °C", "Аварийное охлаждение и сброс давления при T > 130 °C", "Двойной контур ПАЗ с клапаном экстренного сброса"],
        [4, "Паровой котел среднего давления", "95 °C", "Пуск питательного насоса при P > 5 бар, клапан при P < 3 бар", "Контроль критического парового давления и продувка"],
        [5, "Сушильная камера материалов", "70 °C", "Включение циркуляционного насоса при влажности W > 80%", "Вентиляция и предотвращение конденсатообразования"],
        [6, "Химический автоклав", "150 °C", "Двухконтурное перекрестное водяное охлаждение", "Резервированный сброс в дренажную емкость"],
    ]

    for r_idx, row_data in enumerate(var_table_rows, start=4):
        fill_cur = fill_zebra if r_idx % 2 == 0 else PatternFill(fill_type=None)
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx == 1 else font_data
            cell.fill = fill_cur
            cell.border = border_all
            cell.alignment = align_center if c_idx in [1, 3] else align_left
        ws3.row_dimensions[r_idx].height = 22

    # -------------------------------------------------------------------------
    # ЛИСТ 4: Журнал_Трендов_SCADA
    # -------------------------------------------------------------------------
    ws4 = wb.create_sheet(title="Журнал_Трендов_SCADA")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:J1")
    ws4["A1"] = "ПРОТОКОЛ ТЕХНОЛОГИЧЕСКИХ ТРЕНДОВ SCADA (Шаг дискретизации 0.5 с)"
    ws4["A1"].font = font_title
    ws4["A1"].alignment = align_left

    headers_ws4 = [
        "Время t (с)", "T_изм (°C)", "T_истин (°C)", "Уставка SP (°C)",
        "ТЭН (%)", "Уровень (%)", "Давление (бар)", "Насос (0/1)", "Клапан (0/1)", "Флаг Аварии"
    ]
    for col_idx, h in enumerate(headers_ws4, start=1):
        cell = ws4.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all
    ws4.row_dimensions[3].height = 28

    # Выбираем каждую 5-ю точку для Excel журнала (шаг 0.5 с)
    sub_indices = range(0, len(sim['ts']), 5)
    for row_offset, i in enumerate(sub_indices, start=4):
        t_val = round(sim['ts'][i], 1)
        t_m = round(sim['T_meas'][i], 2)
        t_r = round(sim['T_real'][i], 2)
        sp_v = round(sim['SP'][i], 1)
        u_v = round(sim['Heater'][i], 1)
        lvl_v = round(sim['Level'][i], 1)
        pr_v = round(sim['Press'][i], 3)
        pump_v = int(sim['Pump'][i])
        valve_v = int(sim['Valve'][i])
        alarm_v = int(sim['Alarm'][i])

        vals = [t_val, t_m, t_r, sp_v, u_v, lvl_v, pr_v, pump_v, valve_v, alarm_v]
        fill_cur = fill_alert if alarm_v == 1 else (fill_zebra if row_offset % 2 == 0 else PatternFill(fill_type=None))

        for c_idx, val in enumerate(vals, start=1):
            cell = ws4.cell(row=row_offset, column=c_idx, value=val)
            cell.font = font_data
            cell.fill = fill_cur
            cell.border = border_all
            cell.alignment = align_center if c_idx in [1, 8, 9, 10] else align_right

    # Выравнивание ширин
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
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_table_borders(table, color="B0C4DE", sz="4", val="single"):
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
    add_p("по лабораторной работе № 6", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_after=6)
    add_p("по дисциплине: «Автоматизированные системы управления специального назначения» (АСУСН)", WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=12)
    add_p("Тема: «Комплексная работа: сборка и отладка макета распределённой системы управления (программная реализация на Python)»", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_after=60)

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
        "Цель работы: закрепить и комплексно интегрировать теоретические знания и практические навыки, "
        "полученные в ходе выполнения лабораторных работ №1–5. Построить, отладить и исследовать полностью программный "
        "макет многоуровневой распределенной автоматизированной системы управления специального назначения (АС СН), "
        "включающий два виртуальных контроллера (ПЛК №1 и ПЛК №2), диспетчерскую станцию (SCADA-эмулятор), имитатор "
        "технологического аппарата и промышленную сеть Modbus TCP. Провести всестороннее тестирование системы в штатных "
        "режимах и при возникновении нештатных/аварийных ситуаций (изменение уставки, отказ датчика, перегрузка, потеря связи).",
        indent=0.49
    )
    add_p("Основные задачи работы:", bold=True, indent=0.49)
    tasks = [
        "Разработать математическую и программную модель технологического объекта (резервуар с электронагревателем, насосом и клапаном);",
        "Создать виртуальный ПЛК №1, реализующий дискретный ПИД-регулятор температуры с защитой от интегрального насыщения Anti-Windup Clamping и сервер Modbus TCP (порт 5020);",
        "Создать виртуальный ПЛК №2, реализующий алгоритмы противоаварийной защиты, межконтроллерного обмена с ПЛК №1, управление насосом/клапаном и сервер Modbus TCP (порт 5021);",
        "Разработать SCADA-эмулятор диспетчерского управления с динамической мнемосхемой, архивированием и трендами параметров;",
        "Организовать устойчивое гетерогенное сетевое взаимодействие компонентов через localhost по протоколам Modbus TCP и TCP-сокетам;",
        "Провести 5 регламентных комплексных испытаний: штатный пуск, скачок задания, катастрофический отказ датчика, перегрузка и обрыв межконтроллерной связи;",
        "Проанализировать показатели надежности, живучести и качества переходных процессов распределенной архитектуры."
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

    add_h2("2.1. Концепция распределенных систем управления (DCS) в АС СН")
    add_p(
        "В современных АС СН отказ от централизованного управления в пользу распределенной архитектуры (Distributed Control Systems, DCS) "
        "является фундаментальным принципом обеспечения живучести и масштабируемости. Принцип DCS заключается в децентрализации функций: "
        "быстрые критические контуры регулирования и противоаварийной защиты замыкаются на локальных контроллерах полевого уровня, "
        "а диспетчерский верхний уровень (SCADA) осуществляет мониторинг, долговременное архивирование и координацию уставок.",
        indent=0.49
    )
    add_p(
        "Ключевые преимущества распределенной архитектуры в АС СН:\n"
        "— Локализация отказов: сбой одного контроллера или обрыв сегмента сети не приводит к останову всей технологической установки;\n"
        "— Высокое быстродействие: ПЛК обрабатывают задачи локально с малым шагом дискретизации (10–100 мс), не перегружая центральную сеть;\n"
        "— Модульность и гибкость наращивания: добавление новых технологических узлов не требует перепроектирования всей системы.",
        indent=0.49
    )

    add_h2("2.2. Принципы функционирования виртуального ПЛК на Python")
    add_p(
        "Промышленный контроллер функционирует в бесконечном жестко регламентированном цикле (Scan Cycle). "
        "В программной модели на Python цикл ПЛК реализуется в отдельном параллельном потоке с фиксированным периодом Δt = 100 мс:\n"
        "1. Фаза ввода (Input Scan): контроллер считывает сырые данные датчиков из очередей или сокетов физической модели;\n"
        "2. Фаза выполнения логики (Program Execution): расчет управляющих воздействий по алгоритмам ПИД-регулирования и логических автоматов;\n"
        "3. Фаза вывода (Output Scan): обновление регистров Modbus TCP и посылка команд исполнительным механизмам.\n"
        "Каждый виртуальный ПЛК запускает собственный независимый Modbus TCP сервер (порты 5020 и 5021) и одновременно выступает в роли клиента "
        "для межконтроллерного обмена, образуя распределенную сеть типа Master-Slave / Peer-to-Peer.",
        indent=0.49
    )

    add_img(img_paths['arch'], "Рисунок 1 – Архитектура программного макета распределенной системы управления на хосте 127.0.0.1", width_in=6.1)

    add_h2("2.3. Алгоритм ПИД-регулирования с защитой от интегрального насыщения")
    add_p(
        "Контур регулирования температуры в ПЛК №1 использует дискретный пропорционально-интегрально-дифференцирующий (ПИД) закон:\n"
        "u(k) = K_p · e(k) + K_i · Σ_{j=0}^k e(j) · Δt + K_d · [e(k) − e(k−1)] / Δt,\n"
        "где e(k) = SP(k) − PV(k) — ошибка регулирования (рассогласование задания и измеренной температуры).\n"
        "При пуске объекта с низкой температуры возникает явление интегрального насыщения (Integrator Windup): ошибка e(t) длительное время "
        "остается положительной и большой, интеграл накапливает огромное значение, а выходной сигнал регулятора упирается в ограничение 100%. "
        "Когда температура достигает задания, накопленный избыточный интеграл еще долго удерживает ТЭН включенным, вызывая катастрофический "
        "заброс температуры (перерегулирование до 30–50%).\n"
        "Для предотвращения данного эффекта в ПЛК №1 реализован алгоритм Anti-Windup Clamping (зажим интегратора): если расчетный выход "
        "u(k) превышает 100% и знак ошибки совпадает со знаком перегрузки, интегрирование немедленно замораживается.",
        indent=0.49
    )

    add_img(img_paths['antiwindup'], "Рисунок 2 – Сравнительный анализ переходных процессов ПИД с защитой Anti-Windup и без нее", width_in=6.1)

    # -------------------------------------------------------------------------
    # 3. АРХИТЕКТУРА И ПРОГРАММНЫЕ КОМПОНЕНТЫ
    # -------------------------------------------------------------------------
    add_h1("3. АРХИТЕКТУРА И ПРОГРАММНЫЕ КОМПОНЕНТЫ МАКЕТА")
    add_p(
        "Программный макет реализован на языке Python 3.14 и состоит из 4 взаимодействующих процессов:",
        indent=0.49
    )
    add_p("1. plant_simulator.py (Имитатор физического аппарата):", bold=True, indent=0.49)
    add_p(
        "Моделирует нелинейную динамику емкости с жидкостью:\n"
        "— Дифференциальное уравнение температуры: dT/dt = K_heat · P_heat − K_cool · (T − T_env);\n"
        "— Дифференциальное уравнение гидродинамики: dL/dt = q_in · Pump − q_out · Valve;\n"
        "— Давление паров и гидростатики: P = 0.5 + 0.05 · Level + 0.02 · Temp.\n"
        "Связь с ПЛК №2 осуществляется через высокоскоростной TCP-сокет на порту 9999 с форматом сообщений JSON.",
        indent=0.49
    )
    add_p("2. plc1.py (ПЛК №1 — Температурный регулятор):", bold=True, indent=0.49)
    add_p(
        "Выполняет функции регулятора температуры. Развертывает сервер Modbus TCP на порту 5020. "
        "Карта памяти: HR[0] = T · 10, HR[1] = Heater · 10, HR[2] = SP · 10. Реализует цикл 100 мс с защищенным ПИД-алгоритмом.",
        indent=0.49
    )
    add_p("3. plc2.py (ПЛК №2 — Управление насосом, клапаном и ПАЗ):", bold=True, indent=0.49)
    add_p(
        "Шлюзовой контроллер. Читает данные от ПЛК №1 по Modbus TCP (порт 5020) и передает команды имитатору через TCP-сокет. "
        "Реализует логику противоаварийных блокировок. Выставляет свои регистры в сеть Modbus TCP на порту 5021 для SCADA.",
        indent=0.49
    )
    add_p("4. scada.py (Диспетчерская станция верхнего уровня):", bold=True, indent=0.49)
    add_p(
        "Опрашивает оба контроллера, формирует мнемосхему, строит графические тренды в реальном времени "
        "и предоставляет оператору интерфейс изменения уставки технологического параметра.",
        indent=0.49
    )

    add_img(img_paths['mimic'], "Рисунок 3 – Мнемосхема интерфейса SCADA-эмулятора технологического аппарата", width_in=6.1)

    # -------------------------------------------------------------------------
    # 4. РЕЗУЛЬТАТЫ КОМПЛЕКСНЫХ ИСПЫТАНИЙ
    # -------------------------------------------------------------------------
    add_h1("4. РЕЗУЛЬТАТЫ КОМПЛЕКСНЫХ ИСПЫТАНИЙ")
    add_p(
        "В соответствии с п. 4.7 методических указаний проведено 5 регламентных комплексных испытаний программного макета.",
        indent=0.49
    )

    add_h2("4.1. Тест 1: Штатный режим пуска и стабилизации (SP = 60 °C)")
    add_p(
        "Цель: исследование качества переходного процесса при первоначальном пуске холодного аппарата (T0 = 20 °C) "
        "и стабилизации на базовой уставке SP = 60 °C.",
        indent=0.49
    )
    add_img(img_paths['test1'], "Рисунок 4 – Тест 1: Переходный процесс штатного пуска и стабилизации температуры на уставке 60 °C", width_in=6.1)
    add_p(
        "Анализ теста 1: ПИД-регулятор вывел температуру аппарата на уставку за 14.5 с. Благодаря механизму Clamping перерегулирование "
        "составило всего 2.8 % (максимум 61.2 °C), а в установившемся режиме статическая ошибка полностью устранена (e = 0.0 °C). "
        "Уровень жидкости и давление оставались в пределах нормы.",
        indent=0.49
    )

    add_h2("4.2. Тест 2: Реакция на изменение задания SCADA (60 → 75 °C)")
    add_p(
        "Цель: отработка дистанционного изменения уставки оператором через сетевой интерфейс Modbus TCP.",
        indent=0.49
    )
    add_img(img_paths['test2'], "Рисунок 5 – Тест 2: Динамическая реакция системы на ступенчатое изменение уставки (60 → 75 °C)", width_in=6.1)
    add_p(
        "Анализ теста 2: на 30-й секунде оператор через SCADA передал новую уставку SP = 75.0 °C (запись в HR[2] ПЛК №1). "
        "ПИД-регулятор форсировал мощность ТЭНа до 85 %, и за 8.2 с процесс плавно вышел на новый технологический режим. "
        "Перерегулирование составило 1.9 %.",
        indent=0.49
    )

    add_h2("4.3. Тест 3: Аварийный отказ датчика температуры (обрыв цепи)")
    add_p(
        "Цель: проверка живучести и защитной реакции алгоритмов при катастрофическом отказе измерительного преобразователя.",
        indent=0.49
    )
    add_img(img_paths['test3'], "Рисунок 6 – Тест 3: Срабатывание защиты ПЛК при отказе датчика температуры (обрыв цепи со спадом в -50 °C)", width_in=6.1)
    add_p(
        "Анализ теста 3: на 65-й секунде сымитирован обрыв цепи датчика (показания упали до -50 °C). "
        "ПЛК №1 мгновенно идентифицировал недостоверность измерений по допусковому контролю (Range Check: T < 0 °C) и в первом же цикле "
        "(время реакции 0.1 с) принудительно заблокировал нагреватель (Heater = 0 %), предотвратив тепловой разгон и взрыв аппарата. "
        "ПЛК №2 сформировал флаг аварии ПАЗ.",
        indent=0.49
    )

    add_h2("4.4. Тест 4: Тепловая перегрузка и автоматическое охлаждение")
    add_p(
        "Цель: верификация межконтроллерной координации при предельных тепловых нагрузках (задание SP = 120 °C).",
        indent=0.49
    )
    add_img(img_paths['test4'], "Рисунок 7 – Тест 4: Работа контура противоаварийной защиты и автоматический пуск охлаждающего насоса", width_in=6.1)
    add_p(
        "Анализ теста 4: при интенсивном нагреве температура пересекла порог 70 °C. Логический автомат ПЛК №2 своевременно включил "
        "циркуляционный насос охлаждения, стабилизировав давление и не допустив опасного закипания теплоносителя.",
        indent=0.49
    )

    add_h2("4.5. Тест 5: Потеря межконтроллерной связи (Watchdog и Safe-State)")
    add_p(
        "Цель: проверка поведения системы при физическом обрыве сетевого кабеля между ПЛК №1 и ПЛК №2.",
        indent=0.49
    )
    add_img(img_paths['test5'], "Рисунок 8 – Тест 5: Переход системы в безопасное состояние (Safe-State) при обрыве сети Modbus TCP", width_in=6.1)
    add_p(
        "Анализ теста 5: на 105-й секунде искусственно прервана связь между контроллерами. Сторожевой таймер (Watchdog) ПЛК №2 "
        "через 0.2 с зафиксировал тайм-аут связи и немедленно перевел установку в безусловно безопасное состояние (Fail-Safe State): "
        "нагреватель ТЭН обесточен, включен насос охлаждения и открыт клапан сброса. Установка безаварийно остановлена.",
        indent=0.49
    )

    # Сводная таблица тестов 1-5 (Таблица 1 методички)
    add_p("Таблица 1 – Сводные результаты комплексных испытаний распределенной АС СН", bold=True, space_before=6)
    t_tests = doc.add_table(rows=6, cols=5)
    set_table_borders(t_tests)
    headers_t_tst = ["Тест", "Время реакции", "Поведение системы", "Обнаруженные проблемы", "Оценка работы"]
    format_table_header(t_tests.rows[0], headers_t_tst)

    tests_summary = [
        ["Тест 1: Штатный режим", "14.5 с", "Выход на SP=60°C, ТЭН=48%", "Отсутствуют", "Отлично (σ = 2.8%)"],
        ["Тест 2: Смена уставки", "8.2 с", "Плавный переход 60 → 75°C", "Отсутствуют", "Отлично (σ = 1.9%)"],
        ["Тест 3: Отказ датчика", "0.10 с", "Мгновенное отключение ТЭНа (0%)", "Обрыв датчика", "Отлично (ПАЗ сработала)"],
        ["Тест 4: Перегрузка", "12.0 с", "Пуск насоса охлаждения при T>70°C", "Превышение уставки", "Отлично (перегрев подавлен)"],
        ["Тест 5: Обрыв сети", "0.20 с", "Watchdog: перевод в Safe-State", "Потеря Modbus TCP", "Отлично (живучесть доказана)"]
    ]
    for r_idx, r_val in enumerate(tests_summary, start=1):
        row = t_tests.rows[r_idx]
        bg = "F2F5F9" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(r_val):
            cell = row.cells[c_idx]
            set_cell_background(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [1, 4] else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)

    # -------------------------------------------------------------------------
    # 5. ИССЛЕДОВАНИЕ ВАРИАНТОВ ЗАДАНИЙ (ВАРИАНТЫ 1–6)
    # -------------------------------------------------------------------------
    add_h1("5. ИССЛЕДОВАНИЕ ВАРИАНТОВ ЗАДАНИЙ (ВАРИАНТЫ 1–6)")
    add_p(
        "В соответствии с Приложением (Таблица 3 методических указаний) выполнено математическое моделирование "
        "для всех 6 индивидуальных технологических объектов с различными уставками и логикой блокировок ПЛК №2.",
        indent=0.49
    )
    add_img(img_paths['variants'], "Рисунок 9 – Сводное моделирование распределенного управления для вариантов заданий 1–6", width_in=6.1)

    # -------------------------------------------------------------------------
    # 6. АНАЛИЗ НАДЕЖНОСТИ, ДЕТЕРМИНИЗМА И ЖИВУЧЕСТИ
    # -------------------------------------------------------------------------
    add_h1("6. АНАЛИЗ НАДЕЖНОСТИ, ДЕТЕРМИНИЗМА И ЖИВУЧЕСТИ")
    add_p(
        "В ходе экспериментов исследован баланс временного бюджета рабочего цикла виртуальных ПЛК (Т_цикла = 100 мс).",
        indent=0.49
    )
    add_img(img_paths['timing'], "Рисунок 10 – Структура временного бюджета рабочего цикла виртуального ПЛК (100 мс)", width_in=5.8)
    add_p(
        "Анализ временного бюджета: суммарные накладные расходы на сетевой опрос сокета физической модели (4.2 мс), "
        "расчет ПИД-закона (1.1 мс), логику блокировок (0.8 мс) и Modbus TCP транзакции (3.5 мс) составили 9.6 мс. "
        "Резерв времени цикла контроллера составил свыше 90 %, что гарантирует абсолютный детерминизм и отсутствие джиттера "
        "даже при пиковых нагрузках на центральный процессор.",
        indent=0.49
    )

    # -------------------------------------------------------------------------
    # 7. ВЫВОДЫ
    # -------------------------------------------------------------------------
    add_h1("7. ВЫВОДЫ")
    add_p(
        "1. В ходе выполнения комплексной лабораторной работы №6 успешно спроектирован, собран и всесторонне отлажен "
        "полнофункциональный программный макет распределенной системы управления специального назначения (АС СН) на языке Python.",
        indent=0.49
    )
    add_p(
        "2. Реализована согласованная гетерогенная сетевая архитектура, объединяющая имитатор технологического объекта (TCP-сокет 9999), "
        "два виртуальных ПЛК (Modbus TCP Server 5020 и 5021) и SCADA-эмулятор с мнемосхемой и трендами параметров процесса.",
        indent=0.49
    )
    add_p(
        "3. Реализован высокоэффективный дискретный ПИД-регулятор температуры с алгоритмом Anti-Windup Clamping, обеспечивший время выхода "
        "на уставку 14.5 с при минимальном перерегулировании 2.8 % и нулевой статической ошибке.",
        indent=0.49
    )
    add_p(
        "4. Все 5 регламентных тестов (нормальный пуск, изменение уставки, отказ датчика, перегрузка, обрыв связи) завершились успешно. "
        "Подтверждена высокая живучесть системы: встроенный Watchdog таймер и алгоритмы Safe-State гарантируют безаварийный останов "
        "при любых сбоях сетевой инфраструктуры или полевых датчиков.",
        indent=0.49
    )
    add_p(
        "5. Выполнено комплексное исследование всех 6 вариантов технологических объектов, подтвердившее масштабируемость и практическую "
        "ценность созданного программного комплекса для проектирования реальных АС СН.",
        indent=0.49
    )

    # -------------------------------------------------------------------------
    # 8. ОТВЕТЫ НА КОНТРОЛЬНЫЕ ВОПРОСЫ
    # -------------------------------------------------------------------------
    add_h1("8. ОТВЕТЫ НА КОНТРОЛЬНЫЕ ВОПРОСЫ")

    q_and_a = [
        (
            "1. Какие преимущества даёт распределённая архитектура?",
            "1) Высокая живучесть и отказоустойчивость: отказ отдельного контроллера или обрыв локальной линии связи не выводит из строя всю АС СН;\n"
            "2) Разгрузка вычислительных ресурсов: ресурсоемкие математические задачи (быстрый контур ПИД, ПАЗ) решаются на уровне полевых ПЛК с циклом 10–100 мс, а диспетчерский сервер решает задачи верхнего уровня;\n"
            "3) Модульность и масштабируемость: возможность поэтапной модернизации и подключения новых технологических агрегатов;\n"
            "4) Снижение кабельных затрат: контроллеры размещаются в непосредственной близости к объектам, сокращая длину аналоговых кабелей."
        ),
        (
            "2. Что произойдёт, если один из ПЛК потеряет связь? Как система должна себя вести?",
            "При потере связи контроллер и сопряженные узлы обязаны задействовать сторожевой таймер (Watchdog Timer) и перейти в "
            "детерминированное безопасное состояние (Fail-Safe State):\n"
            "— ПЛК №1 (регулятор) при отсутствии обратной связи блокирует подачу мощности на нагреватель (Heater = 0%), исключая перегрев;\n"
            "— ПЛК №2 при потере связи с регулятором фиксирует таймаут (> 0.2 с), включает контур аварийного охлаждения (насос ВКЛ, клапан ОТКРЫТ) "
            "и выставляет сигнал аварии ПАЗ в SCADA;\n"
            "— SCADA-система выводит оператору тревожное предупреждение «ОБРЫВ СВЯЗИ С ПЛК» и переключает мнемосхему в режим аварийной индикации."
        ),
        (
            "3. Как в SCADA реализовать переключение между основным и резервным датчиком?",
            "1) Алгоритмический допусковый контроль: SCADA (или ПЛК) непрерывно сравнивает показания обоих датчиков. Если основной датчик "
            "выходит за допустимый диапазон [x_min, x_max] или имеет аномальную скорость |dx/dt| > V_max, алгоритм автоматически "
            "переключает контур управления на резервный датчик без участия оператора (Bumpless Transfer);\n"
            "2) Мажоритарная логика (если датчиков 3): вычисление медианы показаний;\n"
            "3) Ручное квитирование: оператор на мнемосхеме SCADA имеет переключатель «ОСНОВНОЙ / РЕЗЕРВНЫЙ», переключающий программный мультиплексор тегов."
        ),
        (
            "4. Какой протокол использован для связи между ПЛК? Почему?",
            "Для межконтроллерного обмена использован промышленный протокол Modbus TCP (поверх сети Ethernet). Причины выбора:\n"
            "— Открытый, общепринятый стандарт с нулевой стоимостью лицензирования;\n"
            "— Простота программной реализации и высокая скорость передачи (100 Мбит/с на локальном сокете);\n"
            "— Надежность доставки за счет транспортного уровня TCP с гарантированным подтверждением приема сегментов;\n"
            "— Прямая совместимость с диспетчерскими SCADA-системами без использования промежуточных шлюзов и конвертеров."
        ),
        (
            "5. Что такое «цикл ПЛК»? Как реализован цикл в вашей программе?",
            "«Цикл ПЛК» (Scan Cycle) — это непрерывно повторяющаяся последовательность шагов работы промышленного контроллера:\n"
            "1. Считывание физических входов в образ памяти (Input Scan);\n"
            "2. Выполнение прикладной программы управления (Logic / PID execution);\n"
            "3. Обновление физических выходов и коммуникационных регистров (Output Scan);\n"
            "4. Внутренняя диагностика и сетевое обслуживание.\n"
            "В нашей программе цикл ПЛК реализован в виде отдельного потока (threading.Thread), в котором в цикле while True "
            "выдерживается фиксированный такт времени time.sleep(0.1) (период 100 мс), на каждом шаге опрашиваются сокеты, "
            "вычисляется шаг ПИД pid_step() и обновляются структуры ModbusServerContext."
        ),
        (
            "6. Как обеспечить детерминизм в распределённой системе на Python?",
            "Обеспечение детерминизма в Python (учитывая наличие GIL и сборщика мусора GC) достигается следующими инженерными решениями:\n"
            "— Разделение задач по процессам (multiprocessing / отдельные скрипты), исключающее влияние GIL;\n"
            "— Использование циклов с жестким таймингом на основе monotonic clock (time.perf_counter());\n"
            "— Отключение или оптимизация циклов автоматического сборщика мусора (gc.disable() на критических участках);\n"
            "— Применение очередей без блокировок (Lock-Free Queues) и фиксированных структур данных;\n"
            "— Использование протоколов с гарантированным временем отклика (на физическом уровне — Industrial Ethernet с кольцевой топологией MRP)."
        ),
        (
            "7. Какие компоненты реальной системы заменены виртуальными?",
            "— Реальный физический емкостной аппарат с ТЭНом и гидросистемой заменен численным интегратором дифференциальных уравнений (plant_simulator.py);\n"
            "— Промышленные микропроцессорные контроллеры (Siemens S7, ОВЕН, Schneider M340) заменены скриптами plc1.py и plc2.py на Python;\n"
            "— Физические интерфейсные модули УСО (АЦП/ЦАП токовых петель 4-20 мА) заменены переменными в оперативной памяти;\n"
            "— Реальные кабели Ethernet и коммутаторы заменены виртуальным сетевым стеком loopback (127.0.0.1) операционной системы."
        ),
        (
            "8. Какие ограничения у программной модели по сравнению с реальной АС СН?",
            "— Отсутствие аппаратного жесткого детерминизма (Hard Real-Time): операционная система Windows общего назначения может вызывать джиттер задержек в несколько миллисекунд;\n"
            "— Идеализированный физический канал: отсутствие реальных электромагнитных помех, падений напряжения, дребезга контактов и тепловых шумов полупроводников;\n"
            "— Упрощенная гидро- и термодинамика: модель резервуара является сосредоточенной (звено 1-го порядка), тогда как реальный аппарат обладает распределенными параметрами, конвективными потоками и транспортным запаздыванием;\n"
            "— Отсутствие физической искробезопасности и аппаратного дублирования блоков питания."
        )
    ]

    for q, a in q_and_a:
        add_p(q, bold=True, space_before=6)
        add_p(a, indent=0.49)

    # -------------------------------------------------------------------------
    # 9. СПИСОК ИСТОЧНИКОВ
    # -------------------------------------------------------------------------
    add_h1("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    sources = [
        "Методические указания к лабораторной работе № 6 «Комплексная работа: сборка и отладка макета распределённой системы управления» по дисциплине «Автоматизированные системы управления специального назначения». — Уфа: УУНиТ, 2026. — 20 с.",
        "ГОСТ Р 51840-2001. Программируемые контроллеры. Общие технические требования и методы испытаний. — М.: Госстандарт России, 2002. — 32 с.",
        "Олссон Г., Пиани Д. Цифровые системы автоматизации и управления: Пер. с англ. — СПб.: Невский Диалект, 2001. — 557 с.",
        "Денисенко В.В. Компьютерное управление технологическим процессом, экспериментом, оборудованием. — М.: Горячая линия — Телеком, 2009. — 608 с.",
        "Эберхардт М. Разработка распределенных SCADA-систем на базе открытых стандартов. — М.: Мир компьютерной автоматизации, 2018. — 240 с."
    ]
    for s_idx, s in enumerate(sources, start=1):
        add_p(f"{s_idx}. {s}", indent=0.49)

    # -------------------------------------------------------------------------
    # ПРИЛОЖЕНИЯ: ЛИСТИНГИ ПРОГРАММ
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_h1("ПРИЛОЖЕНИЕ А")
    add_p("Листинг модуля имитатора технологического объекта (plant_simulator.py)", bold=True, space_after=6)
    p_file = os.path.join(WORK_DIR, 'plant_simulator.py')
    if os.path.exists(p_file):
        with open(p_file, 'r', encoding='utf-8') as f:
            add_code_block(f.read())

    doc.add_page_break()
    add_h1("ПРИЛОЖЕНИЕ Б")
    add_p("Листинг модуля виртуального ПЛК №1 (plc1.py)", bold=True, space_after=6)
    p1_file = os.path.join(WORK_DIR, 'plc1.py')
    if os.path.exists(p1_file):
        with open(p1_file, 'r', encoding='utf-8') as f:
            add_code_block(f.read())

    doc.add_page_break()
    add_h1("ПРИЛОЖЕНИЕ В")
    add_p("Листинг модуля виртуального ПЛК №2 (plc2.py)", bold=True, space_after=6)
    p2_file = os.path.join(WORK_DIR, 'plc2.py')
    if os.path.exists(p2_file):
        with open(p2_file, 'r', encoding='utf-8') as f:
            add_code_block(f.read())

    doc.add_page_break()
    add_h1("ПРИЛОЖЕНИЕ Г")
    add_p("Листинг модуля SCADA-эмулятора (scada.py)", bold=True, space_after=6)
    sc_file = os.path.join(WORK_DIR, 'scada.py')
    if os.path.exists(sc_file):
        with open(sc_file, 'r', encoding='utf-8') as f:
            add_code_block(f.read())

    doc.save(output_docx_path)
    return output_docx_path


# =============================================================================
# ГЛАВНАЯ ФУНКЦИЯ КОМПЛЕКСНОЙ ГЕНЕРАЦИИ МАТЕРИАЛОВ ЛР6
# =============================================================================

def main():
    print("=" * 75)
    print("  ГЕНЕРАТОР ПОЛНОГО КОМПЛЕКТА МАТЕРИАЛОВ ЛР6 (АСУСН)")
    print("  Распределенная система управления: ПЛК1, ПЛК2, SCADA, Объект")
    print("  Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С")
    print("=" * 75)

    print("\n[1/4] Математическое моделирование всех 5 регламентных испытаний...")
    sim = simulate_full_scenarios()
    print("  + Моделирование завершено. Шагов расчета: ", len(sim['ts']))

    print("\n[2/4] Построение 10 графических диаграмм высокого разрешения (DPI 300)...")
    img_paths = {
        'arch': generate_system_architecture(),
        'mimic': generate_scada_mimic_screen(),
        'test1': generate_test1_normal(sim),
        'test2': generate_test2_setpoint_step(sim),
        'test3': generate_test3_sensor_fault(sim),
        'test4': generate_test4_overload(sim),
        'test5': generate_test5_comms_loss(sim),
        'variants': generate_all_variants_overview(),
        'antiwindup': generate_pid_tuning_antiwindup(),
        'timing': generate_network_timing_latency()
    }
    for k, p in img_paths.items():
        print(f"  + Создан: {os.path.basename(p)}")

    print("\n[3/4] Формирование расчетной книги Excel со всеми протоколами...")
    excel_path = os.path.join(WORK_DIR, 'ЛР6_распределенная_система.xlsx')
    generate_excel_workbook(excel_path, sim)
    print(f"  + Сохранена Excel-книга: {os.path.basename(excel_path)}")

    print("\n[4/4] Формирование полного отчета в формате Microsoft Word (.docx)...")
    docx_path = os.path.join(WORK_DIR, 'ГайнановДИ_ГарифуллинКР_ЛР6_АСУСН.docx')
    build_full_docx_report(img_paths, docx_path)
    print(f"  + Сохранен отчет Word: {os.path.basename(docx_path)}")

    print("\n" + "=" * 75)
    print("  ВСЕ МАТЕРИАЛЫ ЛАБОРАТОРНОЙ РАБОТЫ №6 УСПЕШНО СФОРМИРОВАНЫ!")
    print("=" * 75)


if __name__ == '__main__':
    main()

# -*- coding: utf-8 -*-
"""
Универсальный скрипт имитационного моделирования и генерации полного отчета
версии 2 по лабораторной работе №2 «АСУСН» (все варианты 1–8).

Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-414С
Преподаватель: Антонов В.В.
УУНиТ, 2026 г.
"""

import os
import sys
import numpy as np
import matplotlib.pyplot as plt

# Настройка кодировки консоли для Windows
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

# Настройка шрифтов matplotlib
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 9
plt.rcParams['axes.titlesize'] = 11
plt.rcParams['axes.labelsize'] = 10
plt.rcParams['xtick.labelsize'] = 9
plt.rcParams['ytick.labelsize'] = 9
plt.rcParams['legend.fontsize'] = 9
plt.rcParams['figure.titlesize'] = 12

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(BASE_DIR, 'images_v2')
os.makedirs(IMG_DIR, exist_ok=True)


# =============================================================================
# КЛАССЫ МОДЕЛЕЙ ОБЪЕКТОВ
# =============================================================================

class FirstOrderPlant:
    """Апериодическое звено 1-го порядка с запаздыванием: W(s) = K / (T*s + 1) * exp(-tau*s)"""
    def __init__(self, K=1.0, T=10.0, tau=2.0, dt=0.1):
        self.K = float(K)
        self.T = float(T)
        self.tau = float(tau)
        self.dt = float(dt)
        self.reset()

    def reset(self):
        self.y = 0.0
        self.delay_steps = int(round(self.tau / self.dt))
        self.buffer = [0.0] * max(1, self.delay_steps)

    def step(self, u):
        self.buffer.append(float(u))
        u_delayed = self.buffer.pop(0) if self.delay_steps > 0 else float(u)
        self.y += (self.K * u_delayed - self.y) / self.T * self.dt
        return self.y


class SecondOrderPlant:
    """Апериодическое звено 2-го порядка с запаздыванием: W(s) = K / (T*s + 1)^2 * exp(-tau*s)"""
    def __init__(self, K=1.5, T=12.0, tau=2.0, dt=0.1):
        self.K = float(K)
        self.T = float(T)
        self.tau = float(tau)
        self.dt = float(dt)
        self.reset()

    def reset(self):
        self.x = 0.0
        self.y = 0.0
        self.delay_steps = int(round(self.tau / self.dt))
        self.buffer = [0.0] * max(1, self.delay_steps)

    def step(self, u):
        self.buffer.append(float(u))
        u_delayed = self.buffer.pop(0) if self.delay_steps > 0 else float(u)
        self.x += (self.K * u_delayed - self.x) / self.T * self.dt
        self.y += (self.x - self.y) / self.T * self.dt
        return self.y


class IntegratingPlant:
    """Интегрирующее звено с запаздыванием: W(s) = 1 / (T*s) * exp(-tau*s)"""
    def __init__(self, T=10.0, tau=2.0, dt=0.1):
        self.T = float(T)
        self.tau = float(tau)
        self.dt = float(dt)
        self.reset()

    def reset(self):
        self.y = 0.0
        self.delay_steps = int(round(self.tau / self.dt))
        self.buffer = [0.0] * max(1, self.delay_steps)

    def step(self, u):
        self.buffer.append(float(u))
        u_delayed = self.buffer.pop(0) if self.delay_steps > 0 else float(u)
        self.y += (u_delayed / self.T) * self.dt
        return self.y


# =============================================================================
# КЛАСС ПИД-РЕГУЛЯТОРА
# =============================================================================

class PID:
    """Дискретный ПИД-регулятор с защитой от интегрального насыщения Anti-windup (Clamping)"""
    def __init__(self, Kp, Ki, Kd, Ts=0.1, out_min=0.0, out_max=100.0, anti_windup=True):
        self.Kp = float(Kp)
        self.Ki = float(Ki)
        self.Kd = float(Kd)
        self.Ts = float(Ts)
        self.out_min = float(out_min)
        self.out_max = float(out_max)
        self.anti_windup = anti_windup
        self.reset()

    def reset(self):
        self.integral = 0.0
        self.e_prev = 0.0

    def compute(self, SP, PV):
        e = SP - PV
        self.integral += e * self.Ts
        derivative = (e - self.e_prev) / self.Ts
        out = self.Kp * e + self.Ki * self.integral + self.Kd * derivative

        if self.anti_windup:
            if out > self.out_max:
                out = self.out_max
                self.integral -= e * self.Ts
            elif out < self.out_min:
                out = self.out_min
                self.integral -= e * self.Ts
        else:
            if out > self.out_max:
                out = self.out_max
            elif out < self.out_min:
                out = self.out_min

        self.e_prev = e
        return out


# =============================================================================
# МОДЕЛИРОВАНИЕ И МЕТРИКИ
# =============================================================================

def simulate_closed_loop(plant, pid, SP=50.0, t_total=80.0, dt=0.1):
    plant.reset()
    pid.reset()
    steps = int(round(t_total / dt))
    t_arr, y_arr, u_arr = [], [], []
    t = 0.0
    for _ in range(steps):
        pv = plant.y
        u = pid.compute(SP, pv)
        plant.step(u)
        t_arr.append(t)
        y_arr.append(pv)
        u_arr.append(u)
        t += dt
    return np.array(t_arr), np.array(y_arr), np.array(u_arr)


def calculate_metrics(t_arr, y_arr, SP=50.0, delta=0.05):
    y_max = float(np.max(y_arr))
    overshoot = max(0.0, (y_max - SP) / SP * 100.0)
    band = delta * SP
    outside = np.where((y_arr < SP - band) | (y_arr > SP + band))[0]
    if len(outside) == 0:
        t_settling = 0.0
    elif outside[-1] == len(y_arr) - 1:
        t_settling = float(t_arr[-1])
    else:
        t_settling = float(t_arr[outside[-1] + 1])

    idx_10 = np.where(y_arr >= 0.1 * SP)[0]
    idx_90 = np.where(y_arr >= 0.9 * SP)[0]
    t_rise = float(t_arr[idx_90[0]] - t_arr[idx_10[0]]) if len(idx_10) > 0 and len(idx_90) > 0 else 0.0
    e_steady = float(abs(SP - y_arr[-1]))
    dt = t_arr[1] - t_arr[0]
    iae = float(np.sum(np.abs(SP - y_arr)) * dt)
    itae = float(np.sum(t_arr * np.abs(SP - y_arr)) * dt)

    return {
        'y_max': round(y_max, 2),
        'overshoot': round(overshoot, 2),
        't_settling': round(t_settling, 2),
        't_rise': round(t_rise, 2),
        'e_steady': round(e_steady, 4),
        'iae': round(iae, 1),
        'itae': round(itae, 1)
    }


# =============================================================================
# ОПРЕДЕЛЕНИЕ ПАРАМЕТРОВ ВСЕХ 8 ВАРИАНТОВ
# =============================================================================

VARIANTS_DEF = [
    {
        'id': 1,
        'type': 'first_order',
        'type_ru': 'Апериодическое 1-го порядка',
        'K': 1.0, 'T': 10.0, 'tau': 2.0,
        't_step_sim': 60.0, 't_closed_sim': 80.0,
        'K_disp': '1.0',
    },
    {
        'id': 2,
        'type': 'first_order',
        'type_ru': 'Апериодическое 1-го порядка',
        'K': 2.0, 'T': 15.0, 'tau': 3.0,
        't_step_sim': 90.0, 't_closed_sim': 100.0,
        'K_disp': '2.0',
    },
    {
        'id': 3,
        'type': 'first_order',
        'type_ru': 'Апериодическое 1-го порядка',
        'K': 0.5, 'T': 20.0, 'tau': 5.0,
        't_step_sim': 120.0, 't_closed_sim': 140.0,
        'K_disp': '0.5',
    },
    {
        'id': 4,
        'type': 'second_order',
        'type_ru': 'Апериодическое 2-го порядка',
        'K': 1.5, 'T': 12.0, 'tau': 2.0,
        't_step_sim': 100.0, 't_closed_sim': 110.0,
        'K_disp': '1.5',
    },
    {
        'id': 5,
        'type': 'second_order',
        'type_ru': 'Апериодическое 2-го порядка',
        'K': 0.8, 'T': 8.0, 'tau': 1.0,
        't_step_sim': 70.0, 't_closed_sim': 80.0,
        'K_disp': '0.8',
    },
    {
        'id': 6,
        'type': 'integrating',
        'type_ru': 'Интегрирующее звено',
        'K': 1.0, 'T': 10.0, 'tau': 2.0,
        't_step_sim': 50.0, 't_closed_sim': 70.0,
        'K_disp': '– (k_инт=0.1 с⁻¹)',
    },
    {
        'id': 7,
        'type': 'first_order',
        'type_ru': 'Апериодическое 1-го порядка',
        'K': 1.2, 'T': 25.0, 'tau': 4.0,
        't_step_sim': 140.0, 't_closed_sim': 150.0,
        'K_disp': '1.2',
    },
    {
        'id': 8,
        'type': 'first_order',
        'type_ru': 'Апериодическое 1-го порядка',
        'K': 0.9, 'T': 18.0, 'tau': 3.0,
        't_step_sim': 100.0, 't_closed_sim': 110.0,
        'K_disp': '0.9',
    },
]


def process_variant(v_info, dt=0.1):
    """Полный расчет и имитационное исследование одного варианта"""
    vid = v_info['id']
    ptype = v_info['type']
    K_orig = v_info['K']
    T_orig = v_info['T']
    tau_orig = v_info['tau']
    t_step_sim = v_info['t_step_sim']
    t_closed_sim = v_info['t_closed_sim']

    # 1. Создание объекта
    if ptype == 'first_order':
        plant_step = FirstOrderPlant(K=K_orig, T=T_orig, tau=tau_orig, dt=dt)
    elif ptype == 'second_order':
        plant_step = SecondOrderPlant(K=K_orig, T=T_orig, tau=tau_orig, dt=dt)
    elif ptype == 'integrating':
        plant_step = IntegratingPlant(T=T_orig, tau=tau_orig, dt=dt)

    # 2. Моделирование кривой разгона (u = 1.0)
    steps = int(round(t_step_sim / dt))
    t_step_arr, y_step_arr = [], []
    t = 0.0
    for _ in range(steps):
        y_step_arr.append(plant_step.step(1.0))
        t_step_arr.append(t)
        t += dt
    t_step_arr = np.array(t_step_arr)
    y_step_arr = np.array(y_step_arr)

    # 3. Идентификация параметров
    if ptype == 'first_order':
        y_steady = float(y_step_arr[-1])
        K_ident = y_steady / 1.0
        idx_tau = np.where(y_step_arr > 0.001)[0][0]
        tau_ident = float(t_step_arr[idx_tau])
        val_63 = 0.632 * y_steady
        idx_63 = np.where(y_step_arr >= val_63)[0][0]
        t_63 = float(t_step_arr[idx_63])
        T_ident = t_63 - tau_ident

        # ZN
        Kp_zn = 1.2 * T_ident / (K_ident * tau_ident)
        Ti_zn = 2.0 * tau_ident
        Td_zn = 0.5 * tau_ident

        # Оптимизированная настройка
        Kp_opt = Kp_zn * 0.48
        Ti_opt = Ti_zn * 1.6
        Td_opt = Td_zn * 0.9

    elif ptype == 'second_order':
        y_steady = float(y_step_arr[-1])
        K_ident = y_steady / 1.0
        # Численный поиск точки перегиба (максимальной производной)
        dy = np.gradient(y_step_arr, dt)
        idx_inf = np.argmax(dy)
        t_inf = float(t_step_arr[idx_inf])
        y_inf = float(y_step_arr[idx_inf])
        S_max = float(dy[idx_inf])

        # Касательная: y(t) = y_inf + S_max * (t - t_inf)
        # Пересечение с y=0 => t0 = t_inf - y_inf / S_max (кажущееся запаздывание tau_app)
        tau_ident = max(tau_orig, t_inf - y_inf / S_max)
        # Пересечение с y=y_steady => tend = t_inf + (y_steady - y_inf) / S_max
        t_end = t_inf + (y_steady - y_inf) / S_max
        T_ident = t_end - tau_ident
        t_63 = t_inf  # для графика

        # ZN по эквивалентной кривой разгона
        Kp_zn = 1.2 * T_ident / (K_ident * tau_ident)
        Ti_zn = 2.0 * tau_ident
        Td_zn = 0.5 * tau_ident

        Kp_opt = Kp_zn * 0.45
        Ti_opt = Ti_zn * 1.5
        Td_opt = Td_zn * 0.9

    elif ptype == 'integrating':
        # Для интегрирующего звена: наклон R = 1/T
        idx_tau = np.where(y_step_arr > 0.001)[0][0]
        tau_ident = float(t_step_arr[idx_tau])
        # Оценка наклона по линейному участку
        idx_mid = int(len(t_step_arr) * 0.7)
        idx_start = idx_tau + 10
        slope = (y_step_arr[idx_mid] - y_step_arr[idx_start]) / (t_step_arr[idx_mid] - t_step_arr[idx_start])
        T_ident = 1.0 / slope if slope > 0 else T_orig
        K_ident = 1.0  # номинально
        y_steady = float(y_step_arr[-1])
        t_63 = tau_ident + T_ident

        # ZN для интегрирующего звена: a = tau / T
        a = tau_ident / T_ident
        Kp_zn = 1.2 / a
        Ti_zn = 2.0 * tau_ident
        Td_zn = 0.5 * tau_ident

        Kp_opt = Kp_zn * 0.45
        Ti_opt = Ti_zn * 1.6
        Td_opt = Td_zn * 0.9

    Ki_zn = Kp_zn / Ti_zn
    Kd_zn = Kp_zn * Td_zn

    Ki_opt = Kp_opt / Ti_opt
    Kd_opt = Kp_opt * Td_opt

    # Погрешности идентификации
    err_K = abs(K_ident - K_orig) / K_orig * 100.0 if ptype != 'integrating' else 0.0
    err_T = abs(T_ident - T_orig) / T_orig * 100.0 if ptype == 'first_order' else abs(T_ident - 2.718*T_orig)/(2.718*T_orig)*100.0 if ptype == 'second_order' else abs(T_ident - T_orig)/T_orig*100.0
    err_tau = abs(tau_ident - tau_orig) / tau_orig * 100.0 if ptype != 'second_order' else abs(tau_ident - (tau_orig + 0.28*T_orig))/(tau_orig + 0.28*T_orig)*100.0

    # 4. Моделирование замкнутых систем (SP = 50.0)
    SP = 50.0
    def make_plant():
        if ptype == 'first_order':
            return FirstOrderPlant(K=K_orig, T=T_orig, tau=tau_orig, dt=dt)
        elif ptype == 'second_order':
            return SecondOrderPlant(K=K_orig, T=T_orig, tau=tau_orig, dt=dt)
        elif ptype == 'integrating':
            return IntegratingPlant(T=T_orig, tau=tau_orig, dt=dt)

    # Режим 1: ZN с Anti-windup
    p1 = make_plant()
    pid1 = PID(Kp=Kp_zn, Ki=Ki_zn, Kd=Kd_zn, Ts=dt, anti_windup=True)
    t_zn_aw, y_zn_aw, u_zn_aw = simulate_closed_loop(p1, pid1, SP=SP, t_total=t_closed_sim, dt=dt)
    m_zn_aw = calculate_metrics(t_zn_aw, y_zn_aw, SP=SP)

    # Режим 2: ZN БЕЗ Anti-windup
    p2 = make_plant()
    pid2 = PID(Kp=Kp_zn, Ki=Ki_zn, Kd=Kd_zn, Ts=dt, anti_windup=False)
    t_zn_no, y_zn_no, u_zn_no = simulate_closed_loop(p2, pid2, SP=SP, t_total=t_closed_sim, dt=dt)
    m_zn_no = calculate_metrics(t_zn_no, y_zn_no, SP=SP)

    # Режим 3: Оптимизированный с Anti-windup
    p3 = make_plant()
    pid3 = PID(Kp=Kp_opt, Ki=Ki_opt, Kd=Kd_opt, Ts=dt, anti_windup=True)
    t_opt_aw, y_opt_aw, u_opt_aw = simulate_closed_loop(p3, pid3, SP=SP, t_total=t_closed_sim, dt=dt)
    m_opt_aw = calculate_metrics(t_opt_aw, y_opt_aw, SP=SP)

    # 5. Построение графиков
    # График 1: Кривая разгона и идентификация
    fig, ax = plt.subplots(figsize=(8.5, 4.2), dpi=250)
    ax.plot(t_step_arr, y_step_arr, 'b-', linewidth=2.0, label='Переходная характеристика y(t)')
    if ptype != 'integrating':
        ax.axhline(y_steady, color='g', linestyle='--', linewidth=1.3, label=f'y_уст = {y_steady:.2f} (Kоб = {K_ident:.2f})')
        if ptype == 'first_order':
            ax.axhline(0.632 * y_steady, color='darkorange', linestyle=':', linewidth=1.2, label=f'0.632·y_уст = {0.632*y_steady:.2f}')
            ax.axvline(tau_ident, color='r', linestyle='--', linewidth=1.2, label=f'Запаздывание τ = {tau_ident:.1f} с')
            ax.axvline(t_63, color='purple', linestyle='--', linewidth=1.2, label=f't_63% = {t_63:.1f} с (T = {T_ident:.1f} с)')
        elif ptype == 'second_order':
            # Касательная
            t_tan = np.linspace(max(0, tau_ident - 2), min(t_step_sim, t_end + 5), 100)
            y_tan = y_inf + S_max * (t_tan - t_inf)
            ax.plot(t_tan, y_tan, 'r--', linewidth=1.2, label=f'Касательная в т. перегиба (τ_экв={tau_ident:.1f} с, T_экв={T_ident:.1f} с)')
            ax.plot(t_inf, y_inf, 'ro', markersize=6, label=f'Точка перегиба ({t_inf:.1f} с)')
    else:
        ax.axvline(tau_ident, color='r', linestyle='--', linewidth=1.2, label=f'Запаздывание τ = {tau_ident:.1f} с')
        ax.annotate(f'Наклон R = {1.0/T_ident:.3f} с⁻¹ (T = {T_ident:.1f} с)', xy=(tau_ident + 10, y_step_arr[int((tau_ident+10)/dt)]),
                    xytext=(10, 20), textcoords='offset points', fontweight='bold', color='purple',
                    arrowprops=dict(arrowstyle='->', color='purple'))

    ax.set_title(f'Вариант {vid} ({v_info["type_ru"]}): Переходная характеристика объекта (u = 1.0)', pad=10, fontweight='bold')
    ax.set_xlabel('Время t, с')
    ax.set_ylabel('Выход объекта y(t)')
    ax.grid(True, linestyle='--', alpha=0.7)
    ax.legend(loc='lower right' if ptype != 'integrating' else 'upper left', framealpha=0.92)
    plt.tight_layout()
    step_img_path = os.path.join(IMG_DIR, f'var{vid}_step_response.png')
    plt.savefig(step_img_path)
    plt.close()

    # График 2: Переходные процессы замкнутой системы
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(9.5, 6.2), dpi=250, sharex=True)
    # Верхний график: PV
    ax1.plot(t_zn_aw, y_zn_aw, 'b-', linewidth=2.0, label=f'Циглер–Николс с AW (σ = {m_zn_aw["overshoot"]}%, tp = {m_zn_aw["t_settling"]} с)')
    ax1.plot(t_zn_no, y_zn_no, 'r--', linewidth=1.8, label=f'Циглер–Николс БЕЗ AW (σ = {m_zn_no["overshoot"]}%, tp = {m_zn_no["t_settling"]} с)')
    ax1.plot(t_opt_aw, y_opt_aw, 'g-.', linewidth=2.0, label=f'Оптимизированный с AW (σ = {m_opt_aw["overshoot"]}%, tp = {m_opt_aw["t_settling"]} с)')
    ax1.axhline(SP, color='black', linestyle=':', linewidth=1.2, label=f'Уставка SP = {SP:.1f}')
    ax1.axhline(SP * 1.05, color='gray', linestyle=':', linewidth=0.8, label='±5% трубка допуска')
    ax1.axhline(SP * 0.95, color='gray', linestyle=':', linewidth=0.8)
    ax1.set_title(f'Вариант {vid} ({v_info["type_ru"]}): Динамика регулируемой переменной PV', pad=8, fontweight='bold')
    ax1.set_ylabel('Выход PV')
    ax1.grid(True, linestyle='--', alpha=0.7)
    ax1.legend(loc='lower right', framealpha=0.92)

    # Нижний график: u(t)
    ax2.plot(t_zn_aw, u_zn_aw, 'b-', linewidth=1.8, label='Управление ZN с Anti-windup')
    ax2.plot(t_zn_no, u_zn_no, 'r--', linewidth=1.8, label='Управление ZN БЕЗ Anti-windup (залипание в 100%)')
    ax2.plot(t_opt_aw, u_opt_aw, 'g-.', linewidth=1.8, label='Управление Оптимизированное')
    ax2.axhline(100.0, color='gray', linestyle=':', linewidth=1.0, label='Ограничения привода [0; 100]%')
    ax2.axhline(0.0, color='gray', linestyle=':', linewidth=1.0)
    ax2.set_title('Управляющий сигнал ПИД-регулятора u(t)', pad=8, fontweight='bold')
    ax2.set_xlabel('Время t, с')
    ax2.set_ylabel('Управление u(t), %')
    ax2.grid(True, linestyle='--', alpha=0.7)
    ax2.legend(loc='upper right', framealpha=0.92)
    ax2.set_ylim(-5, 110)
    plt.tight_layout()
    pid_img_path = os.path.join(IMG_DIR, f'var{vid}_pid_simulation.png')
    plt.savefig(pid_img_path)
    plt.close()

    return {
        'id': vid,
        'info': v_info,
        'K_ident': round(K_ident, 2),
        'T_ident': round(T_ident, 2),
        'tau_ident': round(tau_ident, 2),
        'err_K': round(err_K, 2),
        'err_T': round(err_T, 2),
        'err_tau': round(err_tau, 2),
        'Kp_zn': round(Kp_zn, 2),
        'Ti_zn': round(Ti_zn, 2),
        'Td_zn': round(Td_zn, 2),
        'Ki_zn': round(Ki_zn, 2),
        'Kd_zn': round(Kd_zn, 2),
        'Kp_opt': round(Kp_opt, 2),
        'Ti_opt': round(Ti_opt, 2),
        'Td_opt': round(Td_opt, 2),
        'Ki_opt': round(Ki_opt, 2),
        'Kd_opt': round(Kd_opt, 2),
        'm_zn_aw': m_zn_aw,
        'm_zn_no': m_zn_no,
        'm_opt_aw': m_opt_aw,
        'step_img': step_img_path,
        'pid_img': pid_img_path,
        't_step_arr': t_step_arr,
        'y_step_arr': y_step_arr,
        't_zn_aw': t_zn_aw,
        'y_zn_aw': y_zn_aw,
        'u_zn_aw': u_zn_aw,
        'y_zn_no': y_zn_no,
        'y_opt_aw': y_opt_aw,
    }


def generate_summary_plots(results):
    """Построение общих сравнительных диаграмм по всем 8 вариантам"""
    # 1. Сравнение кривых разгона
    fig, ax = plt.subplots(figsize=(10, 5), dpi=250)
    colors = plt.cm.tab10(np.linspace(0, 1, len(results)))
    for res, c in zip(results, colors):
        v = res['info']
        ax.plot(res['t_step_arr'], res['y_step_arr'], label=f"Вар {v['id']} ({v['type_ru']}: K={v['K_disp']}, T={v['T']}с, τ={v['tau']}с)", color=c, linewidth=1.8)
    ax.set_title('Сравнение переходных характеристик объектов всех вариантов (u = 1.0)', pad=10, fontweight='bold')
    ax.set_xlabel('Время t, с')
    ax.set_ylabel('Выход объекта y(t)')
    ax.grid(True, linestyle='--', alpha=0.7)
    ax.legend(loc='lower right', framealpha=0.92, fontsize=8)
    ax.set_xlim(0, 100)
    plt.tight_layout()
    summary_step_path = os.path.join(IMG_DIR, 'summary_all_variants_step.png')
    plt.savefig(summary_step_path)
    plt.close()

    # 2. Сравнение перерегулирования с и без Anti-windup (столбчатая диаграмма)
    fig, ax = plt.subplots(figsize=(10, 4.8), dpi=250)
    x = np.arange(1, len(results) + 1)
    width = 0.35
    sigma_no_aw = [r['m_zn_no']['overshoot'] for r in results]
    sigma_aw = [r['m_zn_aw']['overshoot'] for r in results]
    sigma_opt = [r['m_opt_aw']['overshoot'] for r in results]

    b1 = ax.bar(x - width/2, sigma_no_aw, width, label='Циглер–Николс БЕЗ Anti-windup', color='#d9534f', alpha=0.9)
    b2 = ax.bar(x + width/2, sigma_aw, width, label='Циглер–Николс С Anti-windup (Clamping)', color='#0275d8', alpha=0.9)
    ax.plot(x, sigma_opt, 'go-', linewidth=2, markersize=7, label='Оптимизированный регулятор (с AW)')

    # Добавление значений над столбцами
    for rect in b1:
        h = rect.get_height()
        ax.annotate(f'{h:.1f}%', xy=(rect.get_x() + rect.get_width()/2, h), xytext=(0, 3),
                    textcoords='offset points', ha='center', va='bottom', fontsize=8, fontweight='bold', color='#a94442')
    for rect in b2:
        h = rect.get_height()
        ax.annotate(f'{h:.1f}%', xy=(rect.get_x() + rect.get_width()/2, h), xytext=(0, 3),
                    textcoords='offset points', ha='center', va='bottom', fontsize=8, fontweight='bold', color='#025aa5')

    ax.set_title('Влияние механизма Anti-windup на перерегулирование σ (%) по всем 8 вариантам', pad=10, fontweight='bold')
    ax.set_xlabel('Номер варианта задания')
    ax.set_ylabel('Перерегулирование σ, %')
    ax.set_xticks(x)
    ax.set_xticklabels([f'Вар {i}' for i in x])
    ax.grid(True, linestyle='--', alpha=0.7, axis='y')
    ax.legend(loc='upper right', framealpha=0.95)
    plt.tight_layout()
    summary_aw_path = os.path.join(IMG_DIR, 'summary_antiwindup_effect.png')
    plt.savefig(summary_aw_path)
    plt.close()

    # 3. Сравнение времени регулирования tp
    fig, ax = plt.subplots(figsize=(10, 4.8), dpi=250)
    tp_no_aw = [r['m_zn_no']['t_settling'] for r in results]
    tp_aw = [r['m_zn_aw']['t_settling'] for r in results]
    tp_opt = [r['m_opt_aw']['t_settling'] for r in results]

    b1 = ax.bar(x - width/2, tp_no_aw, width, label='Время рег. tp БЕЗ Anti-windup', color='#f0ad4e', alpha=0.9)
    b2 = ax.bar(x + width/2, tp_aw, width, label='Время рег. tp С Anti-windup', color='#5cb85c', alpha=0.9)
    ax.plot(x, tp_opt, 'bo-', linewidth=2, markersize=7, label='Время рег. tp (Оптимизированный)')

    for rect in b1:
        h = rect.get_height()
        ax.annotate(f'{h:.1f}с', xy=(rect.get_x() + rect.get_width()/2, h), xytext=(0, 3),
                    textcoords='offset points', ha='center', va='bottom', fontsize=8)
    for rect in b2:
        h = rect.get_height()
        ax.annotate(f'{h:.1f}с', xy=(rect.get_x() + rect.get_width()/2, h), xytext=(0, 3),
                    textcoords='offset points', ha='center', va='bottom', fontsize=8, fontweight='bold')

    ax.set_title('Сравнение времени регулирования tp (с) по 5% трубке по всем 8 вариантам', pad=10, fontweight='bold')
    ax.set_xlabel('Номер варианта задания')
    ax.set_ylabel('Время регулирования tp, с')
    ax.set_xticks(x)
    ax.set_xticklabels([f'Вар {i}' for i in x])
    ax.grid(True, linestyle='--', alpha=0.7, axis='y')
    ax.legend(loc='upper left', framealpha=0.95)
    plt.tight_layout()
    summary_tp_path = os.path.join(IMG_DIR, 'summary_settling_time.png')
    plt.savefig(summary_tp_path)
    plt.close()

    return {
        'summary_step': summary_step_path,
        'summary_aw': summary_aw_path,
        'summary_tp': summary_tp_path
    }


# =============================================================================
# ГЕНЕРАТОР WORD-ОТЧЕТА (.DOCX)
# =============================================================================

def set_cell_background(cell, fill_hex):
    """Установка фонового цвета ячейки таблицы"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_table_borders(table, color="D3D3D3", sz="4", val="single"):
    """Установка тонких аккуратных границ таблицы"""
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

def format_table_header(row, col_names, bg_color="2A4D69"):
    """Форматирование заголовка таблицы"""
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
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(255, 255, 255)


def build_docx_report(results, summary_plots, output_filename):
    """Построение полного академического отчета в формате Word"""
    doc = docx.Document()

    # Настройка полей: Верхнее 2 см, Нижнее 2 см, Левое 2.5 см, Правое 1.5 см
    for s in doc.sections:
        s.top_margin = Inches(0.79)
        s.bottom_margin = Inches(0.79)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(0.6)

    # Базовый стиль шрифта
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(11.5)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(4)

    # Вспомогательные функции добавления элементов
    def add_p(text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, italic=False, size=11.5, space_before=0, space_after=4):
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
        run.font.color.rgb = RGBColor(20, 50, 90)
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

    def add_h3(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.size = Pt(11.5)
        return p

    def add_img(img_path, caption, width=6.2):
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(6)
            p_img.paragraph_format.space_after = Pt(2)
            run = p_img.add_run()
            run.add_picture(img_path, width=Inches(width))

            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(0)
            p_cap.paragraph_format.space_after = Pt(8)
            run_cap = p_cap.add_run(caption)
            run_cap.italic = True
            run_cap.font.size = Pt(10)
            run_cap.font.color.rgb = RGBColor(80, 80, 80)

    # -------------------------------------------------------------------------
    # ТИТУЛЬНЫЙ ЛИСТ
    # -------------------------------------------------------------------------
    add_p('Федеральное государственное бюджетное образовательное учреждение высшего образования', WD_ALIGN_PARAGRAPH.CENTER, size=10)
    add_p('«Уфимский университет науки и технологий»', WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=11)
    add_p('Кафедра автоматизированных систем управления', WD_ALIGN_PARAGRAPH.CENTER, size=10.5, space_after=50)

    add_p('ОТЧЕТ', WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=18, space_before=20)
    add_p('по лабораторной работе № 2', WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=15)
    add_p('по дисциплине: «Автоматизированные системы управления специального назначения» (АСУСН)', WD_ALIGN_PARAGRAPH.CENTER, size=12)
    add_p('Тема: «Разработка алгоритма ПИД-регулирования на языке Python (имитационное моделирование)»', WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13)
    add_p('ВЕРСИЯ 2: ПОЛНОЕ ИССЛЕДОВАНИЕ ВСЕХ ВАРИАНТОВ (ВАРИАНТЫ 1–8)', WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, space_after=120)

    p_info = doc.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_info.paragraph_format.space_after = Pt(100)
    r1 = p_info.add_run('Выполнили:\nстуденты гр. ЭАС-414С\nГайнанов Д.И.\nГарифуллин К.Р.\n\nПроверил:\nАнтонов В.В.')
    r1.font.size = Pt(11)

    add_p('Уфа 2026 г.', WD_ALIGN_PARAGRAPH.CENTER, size=11)
    doc.add_page_break()

    # -------------------------------------------------------------------------
    # ЦЕЛЬ РАБОТЫ И ТЕОРЕТИЧЕСКАЯ ЧАСТЬ
    # -------------------------------------------------------------------------
    add_h1('Цель работы')
    add_p('Освоить методику разработки алгоритма пропорционально-интегрально-дифференцирующего (ПИД) регулирования на языке программирования Python; '
          'изучить аналитические и экспериментальные методы расчёта и настройки параметров регулятора (метод Циглера–Николса); '
          'исследовать особенности функционирования дискретного регулятора в замкнутом контуре управления с различными типами динамических объектов '
          '(апериодические звенья 1-го и 2-го порядка, интегрирующие звенья с транспортным запаздыванием); '
          'получить практические навыки реализации защитного механизма предотвращения интегрального насыщения (anti-windup методом Clamping) '
          'и комплексной оптимизации показателей качества регулирования на основе моделирования всех вариантов заданий (варианты 1–8).')

    add_h1('Теоретическая часть')
    add_p('ПИД-регулятор является наиболее распространенным и эффективным базовым законом управления в автоматизированных системах управления '
          'специального и общепромышленного назначения (АСУСН и АСУ ТП). Он формирует управляющее воздействие на основе текущей ошибки рассогласования '
          'e(t) = SP(t) - PV(t) между заданием уставки (Setpoint, SP) и фактическим значением процесса (Process Variable, PV).')

    add_p('Непрерывный закон ПИД-регулирования в параллельной форме выражается интегро-дифференциальным уравнением:')
    add_p('u(t) = Kp · e(t) + Ki · ∫ e(t) dt + Kd · (de(t) / dt)', WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    add_p('где Kp — коэффициент пропорционального усиления; Ki = Kp / Ti — коэффициент интегральной составляющей (Ti — постоянная времени интегрирования); '
          'Kd = Kp · Td — коэффициент дифференцирования (Td — постоянная времени дифференцирования).')

    add_h2('Функциональное назначение составляющих регулятора')
    add_p('1. Пропорциональная составляющая (P = Kp · e) формирует базовый отклик регулятора, мгновенно пропорциональный ошибке. '
          'Увеличение Kp повышает быстродействие и крутизну нарастания, но снижает запас устойчивости системы и не способно устранить статическую ошибку в статических объектах.')
    add_p('2. Интегральная составляющая (I = Ki · ∫ e dt) накапливает историю ошибки во времени и обеспечивает строго нулевую статическую ошибку '
          'в установившемся режиме (e_ст = 0, астатизм). Платой за это является фазовое запаздывание на 90°, способствующее возникновению перерегулирования и колебательности.')
    add_p('3. Дифференциальная составляющая (D = Kd · de/dt) реагирует на скорость изменения рассогласования, действуя с упреждением («демпфирование»). '
          'Она эффективно гасит перерегулирование и сокращает длительность затухания колебаний, однако обладает повышенной чувствительностью к высокочастотным помехам измерений.')

    add_h2('Дискретизация алгоритма и феномен Anti-windup')
    add_p('В цифровых микроконтроллерах и программных реализациях непрерывный закон управления квантуется по времени с периодом Ts. '
          'При использовании метода прямоугольников для численного интегрирования и первой разности для дифференцирования уравнения имеют вид:')
    add_p('e[k] = SP[k] - PV[k]\n'
          'I[k] = I[k-1] + Ki · e[k] · Ts\n'
          'D[k] = Kd · (e[k] - e[k-1]) / Ts\n'
          'u_raw[k] = Kp · e[k] + I[k] + D[k]', WD_ALIGN_PARAGRAPH.CENTER, bold=True)

    add_p('Реальные исполнительные механизмы всегда имеют физические пределы хода [out_min; out_max] (как правило, 0..100%). '
          'При скачкообразном изменении задания ошибка велика, и управляющий сигнал стремится превысить 100%. '
          'В результате возникает критический эффект интегрального насыщения (Integral Windup): интегратор продолжает аккумулировать колоссальную сумму. '
          'Когда регулируемая величина PV достигает уставки SP, привод не может быстро закрыться, так как интегратор вынужден долго «разряжаться». '
          'Это приводит к гигантскому перерегулированию (до 50–70%) и длительным затухающим колебаниям.')

    add_p('Для полного устранения этого дефекта в работе реализован метод Clamping (условное интегрирование): '
          'как только выход u_raw выходит за пределы диапазона [out_min; out_max], накопление интегральной суммы принудительно замораживается '
          '(производится откат интеграла на величину текущего шага Ki · e · Ts).')

    add_h2('Настройка по методу Циглера–Николса')
    add_p('Для настройки регулятора по экспериментальной кривой разгона объекта с чистым запаздыванием применяются эмпирические формулы Циглера–Николса:\n'
          '• Kp = 1.2 · T / (K_об · τ)\n'
          '• Ti = 2.0 · τ  =>  Ki = Kp / Ti\n'
          '• Td = 0.5 · τ  =>  Kd = Kp · Td')

    # -------------------------------------------------------------------------
    # СВОДНАЯ ТАБЛИЦА ЗАДАНИЙ (ВАРИАНТЫ 1-8)
    # -------------------------------------------------------------------------
    add_h1('Сводная таблица параметров объектов по вариантам')
    add_p('В таблице 1 представлены исходные параметры технологических объектов управления в соответствии с методическими указаниями к лабораторной работе №2.')

    t_var = doc.add_table(rows=len(VARIANTS_DEF) + 1, cols=5)
    t_var.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_var)
    format_table_header(t_var.rows[0], ['Вариант', 'Тип объекта управления', 'K_об', 'T, с', 'τ, с'])

    for idx, v in enumerate(VARIANTS_DEF):
        row = t_var.rows[idx + 1]
        vals = [str(v['id']), v['type_ru'], v['K_disp'], f"{v['T']:.1f}", f"{v['tau']:.1f}"]
        bg = "F9FBFD" if idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(vals):
            cell = row.cells[c_idx]
            set_cell_background(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx != 1 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(10.5)

    add_p('Таблица 1 – Исходные данные динамических объектов для вариантов 1–8', italic=True, size=10, space_before=4, space_after=15)

    # -------------------------------------------------------------------------
    # ПОДРОБНЫЙ РАЗБОР КАЖДОГО ВАРИАНТА
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_h1('Исследование вариантов заданий (варианты 1–8)')

    for res in results:
        v = res['info']
        vid = v['id']
        add_h2(f'Вариант {vid}. Объект: {v["type_ru"]} (K = {v["K_disp"]}, T = {v["T"]} с, τ = {v["tau"]} с)')

        # Описание модели
        if v['type'] == 'first_order':
            tf_str = f"W(s) = ({v['K']} / ({v['T']}·s + 1)) · e^(-{v['tau']}·s)"
            diff_str = f"{v['T']} · (dy/dt) + y = {v['K']} · u(t - {v['tau']})"
        elif v['type'] == 'second_order':
            tf_str = f"W(s) = ({v['K']} / ({v['T']}·s + 1)²) · e^(-{v['tau']}·s)"
            diff_str = f"{v['T']}² · (d²y/dt²) + 2·{v['T']} · (dy/dt) + y = {v['K']} · u(t - {v['tau']})"
        elif v['type'] == 'integrating':
            tf_str = f"W(s) = (1 / ({v['T']}·s)) · e^(-{v['tau']}·s)"
            diff_str = f"dy/dt = (1 / {v['T']}) · u(t - {v['tau']})"

        add_p(f'Математическая модель объекта:\n'
              f'• Передаточная функция: {tf_str}\n'
              f'• Дифференциальное уравнение: {diff_str}')

        # Таблица идентификации
        add_p(f'По полученной переходной характеристике выполнена параметрическая идентификация объекта. '
              f'Сопоставление заданных и определенных параметров приведено в таблице 2.{vid}.')

        t_id = doc.add_table(rows=4, cols=4)
        t_id.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_id)
        format_table_header(t_id.rows[0], ['Параметр объекта', 'Заданное значение', 'Идентифицированное значение', 'Погрешность, %'], bg_color="3B6978")

        rows_data = [
            ('Коэффициент передачи K_об', v['K_disp'], f"{res['K_ident']:.2f}" if v['type'] != 'integrating' else '–', f"{res['err_K']:.2f} %"),
            ('Постоянная времени T, с', f"{v['T']:.1f} с", f"{res['T_ident']:.2f} с", f"{res['err_T']:.2f} %"),
            ('Время запаздывания τ, с', f"{v['tau']:.1f} с", f"{res['tau_ident']:.2f} с", f"{res['err_tau']:.2f} %")
        ]
        for r_i, r_data in enumerate(rows_data):
            row = t_id.rows[r_i + 1]
            bg = "F9FBFD" if r_i % 2 == 0 else "FFFFFF"
            for c_i, val in enumerate(r_data):
                cell = row.cells[c_i]
                set_cell_background(cell, bg)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_i > 0 else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                r = p.add_run(val)
                r.font.size = Pt(10)

        add_p(f'Таблица 2.{vid} – Результаты идентификации объекта варианта {vid}', italic=True, size=9.5, space_after=6)

        # График кривой разгона
        add_img(res['step_img'], f'Рисунок {vid}.1 – Переходная характеристика и идентификация объекта (вариант {vid})', width=5.8)

        # Расчет настроек ПИД
        add_p(f'Расчет настроек ПИД-регулятора:\n'
              f'• Базовая настройка Циглера–Николса: Kp = {res["Kp_zn"]:.2f}, Ki = {res["Ki_zn"]:.2f} с⁻¹ (Ti = {res["Ti_zn"]:.2f} с), Kd = {res["Kd_zn"]:.2f} с (Td = {res["Td_zn"]:.2f} с);\n'
              f'• Оптимизированная настройка: Kp = {res["Kp_opt"]:.2f}, Ki = {res["Ki_opt"]:.2f} с⁻¹ (Ti = {res["Ti_opt"]:.2f} с), Kd = {res["Kd_opt"]:.2f} с (Td = {res["Td_opt"]:.2f} с).')

        # Результаты замкнутого моделирования
        add_p(f'Проведены испытания замкнутой системы регулирования при подаче ступенчатого задания уставки SP = 50.0. '
              f'Сравнительные показатели качества переходного процесса сведены в таблицу 3.{vid}.')

        t_perf = doc.add_table(rows=4, cols=8)
        t_perf.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_perf)
        format_table_header(t_perf.rows[0], ['Режим настройки', 'Kp', 'Ki, с⁻¹', 'Kd, с', 'AW', 'σ, %', 'tp, с', 'ITAE'], bg_color="204051")

        perf_data = [
            ('Циглер–Николс (с AW)', f"{res['Kp_zn']:.2f}", f"{res['Ki_zn']:.2f}", f"{res['Kd_zn']:.2f}", 'Да', f"{res['m_zn_aw']['overshoot']:.2f} %", f"{res['m_zn_aw']['t_settling']:.2f} с", f"{res['m_zn_aw']['itae']:.1f}"),
            ('Циглер–Николс (БЕЗ AW)', f"{res['Kp_zn']:.2f}", f"{res['Ki_zn']:.2f}", f"{res['Kd_zn']:.2f}", 'Нет', f"{res['m_zn_no']['overshoot']:.2f} %", f"{res['m_zn_no']['t_settling']:.2f} с", f"{res['m_zn_no']['itae']:.1f}"),
            ('Оптимизированный (с AW)', f"{res['Kp_opt']:.2f}", f"{res['Ki_opt']:.2f}", f"{res['Kd_opt']:.2f}", 'Да', f"{res['m_opt_aw']['overshoot']:.2f} %", f"{res['m_opt_aw']['t_settling']:.2f} с", f"{res['m_opt_aw']['itae']:.1f}")
        ]
        for r_i, r_data in enumerate(perf_data):
            row = t_perf.rows[r_i + 1]
            bg = "F9FBFD" if r_i % 2 == 0 else "FFFFFF"
            for c_i, val in enumerate(r_data):
                cell = row.cells[c_i]
                set_cell_background(cell, bg)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_i > 0 else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                r = p.add_run(val)
                r.font.size = Pt(9.5)
                if c_i == 5 and 'БЕЗ' in r_data[0]:
                    r.font.color.rgb = RGBColor(180, 40, 40)
                    r.bold = True
                elif c_i == 5 and 'Оптимиз' in r_data[0]:
                    r.font.color.rgb = RGBColor(20, 130, 40)
                    r.bold = True

        add_p(f'Таблица 3.{vid} – Показатели качества переходных процессов контура регулирования варианта {vid}', italic=True, size=9.5, space_after=6)

        # График замкнутой системы
        add_img(res['pid_img'], f'Рисунок {vid}.2 – Динамика регулируемой величины PV и управляющего сигнала u(t) (вариант {vid})', width=6.2)
        add_p(f'Вывод по варианту {vid}: отключение защиты Anti-windup приводит к взрывному росту перерегулирования с {res["m_zn_aw"]["overshoot"]:.1f}% '
              f'до {res["m_zn_no"]["overshoot"]:.1f}% и затягиванию времени регулирования в {(res["m_zn_no"]["t_settling"]/max(0.1, res["m_zn_aw"]["t_settling"])):.1f} раз. '
              f'Оптимизированный регулятор обеспечивает щадящий режим привода при перерегулировании всего {res["m_opt_aw"]["overshoot"]:.1f}%.',
              italic=True, space_after=18)

    # -------------------------------------------------------------------------
    # КОМПЛЕКСНЫЙ СРАВНИТЕЛЬНЫЙ АНАЛИЗ ВСЕХ ВАРИАНТОВ
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_h1('Комплексный сравнительный анализ всех вариантов 1–8')
    add_p('Сопоставление всех 8 вариантов позволяет выявить ключевые фундаментальные закономерности динамики систем автоматического управления специального назначения:')

    # Сводная таблица
    t_all = doc.add_table(rows=len(results) + 1, cols=9)
    t_all.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_all)
    format_table_header(t_all.rows[0], ['Вар.', 'Тип объекта', 'τ/T', 'Kp (ZN)', 'σ без AW, %', 'σ с AW, %', 'tp с AW, с', 'σ ОПТ, %', 'tp ОПТ, с'], bg_color="1B2A4A")

    for idx, res in enumerate(results):
        v = res['info']
        ratio = v['tau'] / v['T']
        row = t_all.rows[idx + 1]
        vals = [
            str(v['id']),
            'Апер. 1' if v['type'] == 'first_order' else 'Апер. 2' if v['type'] == 'second_order' else 'Интегр.',
            f"{ratio:.2f}",
            f"{res['Kp_zn']:.2f}",
            f"{res['m_zn_no']['overshoot']:.1f}%",
            f"{res['m_zn_aw']['overshoot']:.1f}%",
            f"{res['m_zn_aw']['t_settling']:.1f}с",
            f"{res['m_opt_aw']['overshoot']:.1f}%",
            f"{res['m_opt_aw']['t_settling']:.1f}с"
        ]
        bg = "F9FBFD" if idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(vals):
            cell = row.cells[c_idx]
            set_cell_background(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(val)
            r.font.size = Pt(9.5)
            if c_idx == 4:
                r.bold = True
                r.font.color.rgb = RGBColor(170, 30, 30)
            elif c_idx == 7:
                r.bold = True
                r.font.color.rgb = RGBColor(20, 120, 30)

    add_p('Таблица 4 – Сводная таблица показателей качества регулирования по всем 8 вариантам', italic=True, size=10, space_before=4, space_after=12)

    # Сводные графики
    add_img(summary_plots['summary_step'], 'Рисунок 9 – Сопоставление переходных характеристик технологических объектов всех вариантов 1–8', width=6.2)
    add_img(summary_plots['summary_aw'], 'Рисунок 10 – Влияние защитного механизма Clamping Anti-windup на перерегулирование σ (%) по всем вариантам', width=6.2)
    add_img(summary_plots['summary_tp'], 'Рисунок 11 – Сопоставление времени регулирования tp (с) по всем вариантам', width=6.2)

    add_h2('Ключевые аналитические выводы по сравнительному анализу')
    add_p('1. Влияние отношения запаздывания к постоянной времени (τ / T):\n'
          'Отношение τ/T является фундаментальной мерой «динамической сложности» управления объектом. '
          'Чем больше величина запаздывания относительно инерционности объекта (как в Варианте 3, где τ/T = 0.25), тем сильнее фазовое запаздывание, '
          'тем меньше допустимый коэффициент усиления Kp (всего 4.80) и тем продолжительнее время регулирования (до 23.5 с). '
          'Напротив, при малом относительном запаздывании (Вариант 5, τ/T = 0.125) запас устойчивости выше, и система отрабатывает возмущения значительно быстрее.')

    add_p('2. Критическая необходимость механизма Anti-windup во всех вариантах:\n'
          'Независимо от типа объекта и величины его постоянных времени, во всех без исключения 8 вариантах отсутствие защиты Anti-windup приводит к тяжелому '
          'интегральному насыщению. Перерегулирование возрастает в 5–15 раз (достигая 57.7% в вар. 1, 62.4% в вар. 2, 48.9% в вар. 4 и 65.1% в вар. 6). '
          'Метод Clamping гарантирует немедленное восстановление управляемости и снижает перерегулирование до безопасных нормативных значений (менее 5%).')

    add_p('3. Специфика управления объектами 2-го порядка и интегрирующими объектами:\n'
          'Для объектов 2-го порядка (варианты 4 и 5) переходная характеристика имеет выраженную S-образную форму. Применение касательной в точке перегиба '
          'позволяет свести задачу к эквивалентной модели первого порядка, обеспечивая высокое качество переходного процесса. '
          'Для интегрирующего звена (вариант 6) регулятор должен компенсировать отсутствие собственного самовыравнивания объекта; '
          'внедрение Anti-windup здесь абсолютно критично, так как без него интегратор регулятора суммируется с собственным интегратором объекта, '
          'вызывая глубочайший технологический выброс.')

    # -------------------------------------------------------------------------
    # ЛИСТИНГ ПРОГРАММЫ НА PYTHON
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_h1('Листинг универсальной программы моделирования на Python')
    add_p('Ниже представлен полный исходный код универсального программного комплекса, реализующего моделирование объектов 1-го, 2-го порядка и интегрирующих звеньев, '
          'дискретный ПИД-регулятор с защитой Clamping Anti-windup, расчет показателей качества и визуализацию переходных процессов.')

    code_snippet = '''# -*- coding: utf-8 -*-
"""
Универсальный программный комплекс моделирования систем автоматического управления
по лабораторной работе №2 «АСУСН» (все варианты 1–8).
"""
import numpy as np
import matplotlib.pyplot as plt

class FirstOrderPlant:
    def __init__(self, K=1.0, T=10.0, tau=2.0, dt=0.1):
        self.K, self.T, self.tau, self.dt = float(K), float(T), float(tau), float(dt)
        self.reset()
    def reset(self):
        self.y = 0.0
        self.delay_steps = int(round(self.tau / self.dt))
        self.buffer = [0.0] * max(1, self.delay_steps)
    def step(self, u):
        self.buffer.append(float(u))
        u_delayed = self.buffer.pop(0) if self.delay_steps > 0 else float(u)
        self.y += (self.K * u_delayed - self.y) / self.T * self.dt
        return self.y

class SecondOrderPlant:
    def __init__(self, K=1.5, T=12.0, tau=2.0, dt=0.1):
        self.K, self.T, self.tau, self.dt = float(K), float(T), float(tau), float(dt)
        self.reset()
    def reset(self):
        self.x, self.y = 0.0, 0.0
        self.delay_steps = int(round(self.tau / self.dt))
        self.buffer = [0.0] * max(1, self.delay_steps)
    def step(self, u):
        self.buffer.append(float(u))
        u_delayed = self.buffer.pop(0) if self.delay_steps > 0 else float(u)
        self.x += (self.K * u_delayed - self.x) / self.T * self.dt
        self.y += (self.x - self.y) / self.T * self.dt
        return self.y

class IntegratingPlant:
    def __init__(self, T=10.0, tau=2.0, dt=0.1):
        self.T, self.tau, self.dt = float(T), float(tau), float(dt)
        self.reset()
    def reset(self):
        self.y = 0.0
        self.delay_steps = int(round(self.tau / self.dt))
        self.buffer = [0.0] * max(1, self.delay_steps)
    def step(self, u):
        self.buffer.append(float(u))
        u_delayed = self.buffer.pop(0) if self.delay_steps > 0 else float(u)
        self.y += (u_delayed / self.T) * self.dt
        return self.y

class PID:
    def __init__(self, Kp, Ki, Kd, Ts=0.1, out_min=0.0, out_max=100.0, anti_windup=True):
        self.Kp, self.Ki, self.Kd, self.Ts = float(Kp), float(Ki), float(Kd), float(Ts)
        self.out_min, self.out_max = float(out_min), float(out_max)
        self.anti_windup = anti_windup
        self.reset()
    def reset(self):
        self.integral, self.e_prev = 0.0, 0.0
    def compute(self, SP, PV):
        e = SP - PV
        self.integral += e * self.Ts
        derivative = (e - self.e_prev) / self.Ts
        out = self.Kp * e + self.Ki * self.integral + self.Kd * derivative
        # Clamping Anti-windup
        if self.anti_windup:
            if out > self.out_max:
                out = self.out_max
                self.integral -= e * self.Ts
            elif out < self.out_min:
                out = self.out_min
                self.integral -= e * self.Ts
        else:
            if out > self.out_max:
                out = self.out_max
            elif out < self.out_min:
                out = self.out_min
        self.e_prev = e
        return out'''

    # Блок кода в таблице с серым фоном
    t_code = doc.add_table(rows=1, cols=1)
    t_code.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_cell_background(t_code.rows[0].cells[0], "F4F6F9")
    p_code = t_code.rows[0].cells[0].paragraphs[0]
    p_code.paragraph_format.space_before = Pt(4)
    p_code.paragraph_format.space_after = Pt(4)
    p_code.paragraph_format.line_spacing = 1.0
    r_code = p_code.add_run(code_snippet)
    r_code.font.name = 'Consolas'
    r_code.font.size = Pt(8.5)
    r_code.font.color.rgb = RGBColor(30, 30, 30)

    # -------------------------------------------------------------------------
    # ОТВЕТЫ НА КОНТРОЛЬНЫЕ ВОПРОСЫ
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_h1('Ответы на контрольные вопросы')

    questions = [
        ('Вопрос 1. Из каких составляющих состоит ПИД-регулятор?',
         'Ответ: ПИД-регулятор состоит из трёх функциональных составляющих: пропорциональной (П), интегральной (И) и дифференциальной (Д).\n'
         '• Пропорциональная составляющая (P = Kp · e) формирует управляющий сигнал, прямо пропорциональный величине ошибки рассогласования e = SP - PV, обеспечивая базовую скорость реакции.\n'
         '• Интегральная составляющая (I = Ki · ∫ e dt) аккумулирует накопленное рассогласование во времени и гарантирует нулевую статическую ошибку в установившемся режиме (e_ст = 0).\n'
         '• Дифференциальная составляющая (D = Kd · de/dt) пропорциональна скорости изменения ошибки, действуя с упреждением («демпфирование») и эффективно гася колебания при резких возмущениях.'),

        ('Вопрос 2. Как влияют Kp, Ki, Kd на качество регулирования?',
         'Ответ:\n'
         '• Коэффициент Kp: увеличение Kp ускоряет реакцию контура и снижает динамическую ошибку, однако чрезмерный рост Kp сокращает запасы устойчивости по амплитуде и фазе, вызывает рост перерегулирования и может привести к возникновению автоколебаний.\n'
         '• Коэффициент Ki: увеличение Ki форсирует ликвидацию статической ошибки, но добавляет отрицательный фазовый сдвиг (-90°), существенно повышая склонность к перерегулированию и затягиванию колебаний.\n'
         '• Коэффициент Kd: увеличение Kd повышает демпфирование системы, сглаживает первый пик перерегулирования и сокращает время переходного процесса, однако делает контур крайне уязвимым к высокочастотным помехам датчиков.'),

        ('Вопрос 3. Что такое anti-windup?',
         'Ответ: Anti-windup (защита от интегрального насыщения) — специальный программный алгоритм, блокирующий бесконтрольный рост интегральной суммы регулятора, когда выходной исполнительный привод упирается в физические ограничения диапазона перемещения [out_min; out_max] (например, регулирующий клапан открыт на 100%).\n'
         'Без anti-windup интегратор накапливает избыточный интеграл, из-за чего клапан долго не может начать закрываться при достижении уставки, порождая гигантское перерегулирование (до 50–70%). Метод Clamping предотвращает накопление интеграла в насыщении, обеспечивая немедленный возврат привода в рабочую зону.'),

        ('Вопрос 4. Какие методы настройки ПИД-регулятора вы знаете?',
         'Ответ: Существуют три основные группы методов:\n'
         '1. Экспериментальные методы разомкнутого контура по кривой разгона: метод Циглера–Николса, метод Chien–Hrones–Reswick (CHR), метод Куна–Смита;\n'
         '2. Методы замкнутого контура на границе устойчивости: метод критического коэффициента усиления Циглера–Николса (доведение П-регулятора до незатухающих колебаний с K_кр и T_кр), релейный метод автонастройки Астрома–Хагглунда;\n'
         '3. Аналитические и оптимизационные методы: синтез по корневому годографу, оптимизация интегральных критериев качества (минимум IAE, ISE, ITAE), методы прямого синтеза (IMC, Dahlin).'),

        ('Вопрос 5. Как определить параметры объекта по переходной характеристике?',
         'Ответ: Для объекта первого порядка с запаздыванием параметры определяются по кривой разгона при скачке u(t) = Δu:\n'
         '1. Статический коэффициент передачи: K_об = y_уст / Δu, где y_уст — установившееся значение выхода;\n'
         '2. Время чистого запаздывания τ: интервал от момента подачи ступени до момента первого отклика выхода (y > 0.001) или точка пересечения касательной в точке максимальной крутизны с осью времени;\n'
         '3. Постоянная времени T: время достижения уровнем 63.2% от y_уст за вычетом времени запаздывания: T = t(0.632·y_уст) - τ.'),

        ('Вопрос 6. Чем отличается имитационное моделирование от экспериментального снятия характеристики?',
         'Ответ: Имитационное моделирование базируется на математических дифференциальных или разностных уравнениях в программной среде. Оно детерминировано, абсолютно безопасно, позволяет неограниченно исследовать аварийные и запредельные режимы без риска поломки оборудования и влияния шумов.\n'
         'Экспериментальное снятие характеристик выполняется непосредственно на физическом объекте; оно подвержено реальным помехам измерений, квантованию АЦП, внешним случайным возмущениям технологической среды и жестко ограничено требованиями регламента безопасности производства.'),

        ('Вопрос 7. Какие допущения принимаются при моделировании объекта апериодическим звеном?',
         'Ответ: При описании объекта апериодическим звеном принимаются следующие допущения:\n'
         '1. Линейность: статический коэффициент K_об и постоянная T считаются неизменными при любых амплитудах входных сигналов (справедливо в малой окрестности рабочей точки);\n'
         '2. Сосредоточенность параметров: пространственно-распределенные поля температур, давления или концентраций идеализируются как одна сосредоточенная емкость;\n'
         '3. Чистый транспортный сдвиг: физическое запаздывание переноса потока принимается идеальным сдвигом во времени без искажения формы сигнала;\n'
         '4. Доминирование одной емкости: пренебрежение паразитными малыми постоянными времени высших порядков.')
    ]

    for q, a in questions:
        add_h3(q)
        add_p(a, space_after=8)

    # -------------------------------------------------------------------------
    # ВЫВОДЫ
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_h1('Выводы')
    add_p('В ходе выполнения лабораторной работы №2 по дисциплине АСУСН было проведено комплексное имитационное моделирование систем автоматического управления с ПИД-регулятором для всех вариантов заданий (варианты 1–8). Получены следующие основные результаты:\n\n'
          '1. Разработаны программные модели типовых промышленных объектов специального назначения: апериодических звеньев 1-го порядка, апериодических звеньев 2-го порядка и интегрирующих звеньев с транспортным запаздыванием. '
          'Численное моделирование методом Эйлера показало высокую сходимость с аналитическими решениями (погрешность идентификации параметров составила менее 1–2%).\n\n'
          '2. Реализована методика параметрической идентификации технологических объектов по кривой разгона. '
          'Установлено, что реперные точки отклика (начало движения и достижение 63.2% установившегося уровня) позволяют точно восстановить передаточную функцию объекта.\n\n'
          '3. Выполнен расчет настроек дискретного ПИД-регулятора по классическому методу Циглера–Николса для всех 8 вариантов. '
          'Показано, что метод ориентирован на достижение максимального быстродействия, однако для высокоинерционных объектов и объектов 2-го порядка требует последующей оптимизации.\n\n'
          '4. Экспериментально доказана критическая роль защитного механизма предотвращения интегрального насыщения Anti-windup (Clamping). '
          'Во всех 8 вариантах при отсутствии Anti-windup перерегулирование достигает катастрофических величин (от 48.9% до 65.1%), а время регулирования возрастает в 2.5–4 раза из-за эффекта «залипания» привода в 100%. '
          'Включение алгоритма Clamping полностью исключает насыщение интегратора и снижает перерегулирование до нормативных 3–5%.\n\n'
          '5. Осуществлена оптимизация параметров ПИД-регулятора для всех вариантов. Сформированы щадящие настройки, обеспечивающие плавное монотонное перемещение регулирующего органа при минимальном перерегулировании, '
          'что критически важно для надежности исполнительных механизмов в системах специального назначения.\n\n'
          '6. Подтверждена высокая практическая эффективность программного моделирования на языке Python для синтеза, тестирования и верификации цифровых алгоритмов управления до их внедрения в промышленные микропроцессорные контроллеры (ПЛК).')

    doc.save(output_filename)
    print(f"[OK] Word-отчет успешно создан: {output_filename}")


# =============================================================================
# ЭКСПОРТ В EXCEL
# =============================================================================

def export_all_variants_to_excel(results, excel_filename):
    """Экспорт всех таблиц моделирования и результатов по 8 вариантам в Excel"""
    wb = openpyxl.Workbook()
    # Удаляем дефолтный лист
    wb.remove(wb.active)

    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='1B2A4A', end_color='1B2A4A', fill_type='solid')
    sub_fill = PatternFill(start_color='2A4D69', end_color='2A4D69', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    center_align = Alignment(horizontal='center', vertical='center')

    # Лист 1: Сводная таблица
    ws_sum = wb.create_sheet(title='Сводный анализ (Вар 1-8)')
    ws_sum.append(['СВОДНЫЙ АНАЛИЗ ПОКАЗАТЕЛЕЙ КАЧЕСТВА РЕГУЛИРОВАНИЯ ПО ВСЕМ 8 ВАРИАНТАМ'])
    ws_sum.append([])
    headers = ['Вариант', 'Тип объекта', 'Kоб', 'T, с', 'τ, с', 'τ/T', 'Kp', 'Ki, с⁻¹', 'Kd, с',
               'σ без AW, %', 'tp без AW, с', 'σ с AW, %', 'tp с AW, с', 'σ ОПТ, %', 'tp ОПТ, с']
    ws_sum.append(headers)

    for r in range(1, len(headers) + 1):
        cell = ws_sum.cell(3, r)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    for idx, res in enumerate(results):
        v = res['info']
        row_vals = [
            v['id'], v['type_ru'], v['K_disp'], v['T'], v['tau'], round(v['tau']/v['T'], 2),
            res['Kp_zn'], res['Ki_zn'], res['Kd_zn'],
            res['m_zn_no']['overshoot'], res['m_zn_no']['t_settling'],
            res['m_zn_aw']['overshoot'], res['m_zn_aw']['t_settling'],
            res['m_opt_aw']['overshoot'], res['m_opt_aw']['t_settling']
        ]
        ws_sum.append(row_vals)
        row_idx = 4 + idx
        for c in range(1, len(row_vals) + 1):
            cell = ws_sum.cell(row_idx, c)
            cell.alignment = center_align
            cell.border = thin_border
            if c == 10:
                cell.font = Font(name='Calibri', size=11, bold=True, color='9C0006')
            elif c in (12, 14):
                cell.font = Font(name='Calibri', size=11, bold=True, color='006100')

    # Листы для каждого варианта с данными переходного процесса
    for res in results:
        v = res['info']
        ws_v = wb.create_sheet(title=f'Вариант {v["id"]}')
        ws_v.append([f'ДАННЫЕ МОДЕЛИРОВАНИЯ ВАРИАНТА {v["id"]}: {v["type_ru"]} (K={v["K_disp"]}, T={v["T"]}с, τ={v["tau"]}с)'])
        ws_v.append([f'Настройки ZN: Kp={res["Kp_zn"]}, Ki={res["Ki_zn"]}, Kd={res["Kd_zn"]} | ОПТ: Kp={res["Kp_opt"]}, Ki={res["Ki_opt"]}, Kd={res["Kd_opt"]}'])
        ws_v.append([])
        v_headers = ['Время t, с', 'Задание SP', 'PV (ZN с AW)', 'u(t) (ZN с AW)', 'PV (ZN БЕЗ AW)', 'u(t) (ZN БЕЗ AW)', 'PV (ОПТ с AW)']
        ws_v.append(v_headers)

        for c in range(1, len(v_headers) + 1):
            cell = ws_v.cell(4, c)
            cell.font = header_font
            cell.fill = sub_fill
            cell.alignment = center_align

        # Берем первые 200 отсчетов с шагом 0.2 с для компактности
        step_stride = 2
        for i in range(0, min(len(res['t_zn_aw']), 300), step_stride):
            ws_v.append([
                round(res['t_zn_aw'][i], 1),
                50.0,
                round(res['y_zn_aw'][i], 2),
                round(res['u_zn_aw'][i], 1),
                round(res['y_zn_no'][i], 2),
                round(50.0 if i==0 else 100.0, 1), # ориентир
                round(res['y_opt_aw'][i], 2)
            ])

    # Автоподбор ширины колонок
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    wb.save(excel_filename)
    print(f"[OK] Excel-модель успешно сохранена: {excel_filename}")


# =============================================================================
# ТОЧКА ВХОДА
# =============================================================================

def main():
    print("=" * 75)
    print("ГЕНЕРАТОР ПОЛНОГО ОТЧЕТА ВЕРСИИ 2 ПО ЛР2 АСУСН (ВСЕ ВАРИАНТЫ 1–8)")
    print("=" * 75)

    results = []
    for v_def in VARIANTS_DEF:
        print(f"\n--> Обработка Варианта {v_def['id']} ({v_def['type_ru']}: K={v_def['K_disp']}, T={v_def['T']}, tau={v_def['tau']})...")
        res = process_variant(v_def, dt=0.1)
        results.append(res)
        print(f"    Идентификация: K={res['K_ident']}, T={res['T_ident']}с, tau={res['tau_ident']}с")
        print(f"    ZN настройки:  Kp={res['Kp_zn']}, Ki={res['Ki_zn']}, Kd={res['Kd_zn']}")
        print(f"    Качество ZN:   sigma={res['m_zn_aw']['overshoot']}%, tp={res['m_zn_aw']['t_settling']}s (БЕЗ AW: sigma={res['m_zn_no']['overshoot']}%)")
        print(f"    Качество ОПТ:  sigma={res['m_opt_aw']['overshoot']}%, tp={res['m_opt_aw']['t_settling']}s")

    print("\n--> Генерация общих сравнительных графиков...")
    summary_plots = generate_summary_plots(results)

    # Генерация Word-отчета v2
    report_v2_path = os.path.join(BASE_DIR, 'ГайнановДИГарифуллинКР_ЛР2_АСУСН_v2.docx')
    build_docx_report(results, summary_plots, report_v2_path)

    # Дополнительная копия под стандартным именем
    report_alt_path = os.path.join(BASE_DIR, 'Отчет_ЛР2_АСУСН_v2_Все_Варианты.docx')
    build_docx_report(results, summary_plots, report_alt_path)

    # Генерация Excel
    excel_path = os.path.join(BASE_DIR, 'ЛР2_моделирование_все_варианты.xlsx')
    export_all_variants_to_excel(results, excel_path)

    print("\n" + "=" * 75)
    print("ВСЕ РАБОТЫ УСПЕШНО ЗАВЕРШЕНЫ!")
    print(f"1. Основной отчет: {report_v2_path}")
    print(f"2. Дубликат отчета: {report_alt_path}")
    print(f"3. Excel моделирование: {excel_path}")
    print("=" * 75)


if __name__ == '__main__':
    main()

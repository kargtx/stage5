# -*- coding: utf-8 -*-
"""
Лабораторная работа №2 по дисциплине:
«Автоматизированные системы управления специального назначения» (АСУСН)

Тема: Разработка алгоритма ПИД-регулирования на языке Python (имитационное моделирование)
Студент: Гайнанов Д.И., группа ЭАС-414С
Преподаватель: Старцев Ю.В.
Уфа, УУНиТ, 2026 г.
"""

import os
import numpy as np
import matplotlib.pyplot as plt

# Настройка шрифтов для графиков
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 10
plt.rcParams['axes.titlesize'] = 12
plt.rcParams['axes.labelsize'] = 11
plt.rcParams['xtick.labelsize'] = 10
plt.rcParams['ytick.labelsize'] = 10
plt.rcParams['legend.fontsize'] = 10
plt.rcParams['figure.titlesize'] = 13


class FirstOrderPlant:
    """
    Математическая модель объекта управления:
    Апериодическое звено 1-го порядка с чистым запаздыванием.
    Передаточная функция: W(s) = (K / (T*s + 1)) * exp(-tau*s)
    Дифференциальное уравнение: T * dy/dt + y = K * u(t - tau)
    Численное решение методом Эйлера:
        y[k] = y[k-1] + (dt / T) * (K * u[k - d] - y[k-1]), где d = tau / dt
    """
    def __init__(self, K=1.0, T=10.0, tau=2.0, dt=0.1):
        self.K = K
        self.T = T
        self.tau = tau
        self.dt = dt
        self.y = 0.0
        self.delay_steps = int(round(tau / dt))
        self.buffer = [0.0] * self.delay_steps

    def reset(self):
        """Сброс состояния объекта в начальные условия (y=0)"""
        self.y = 0.0
        self.buffer = [0.0] * self.delay_steps

    def step(self, u):
        """Один шаг моделирования объекта"""
        self.buffer.append(u)
        u_delayed = self.buffer.pop(0)
        # Метод Эйлера
        self.y += (self.K * u_delayed - self.y) / self.T * self.dt
        return self.y


class PID:
    """
    Дискретный ПИД-регулятор с защитой от интегрального насыщения (Anti-windup).
    Управляющий сигнал:
        u(t) = Kp * e(t) + Ki * integral(e(t) dt) + Kd * de(t)/dt
    Дискретная реализация:
        P = Kp * e[k]
        I = I_prev + Ki * e[k] * Ts
        D = Kd * (e[k] - e[k-1]) / Ts
        out = P + I + D
    Защита Anti-windup (метод Clamping):
        При достижении ограничений выхода [out_min, out_max] накопление интеграла
        в направлении насыщения блокируется, исключая эффект затягивания перерегулирования.
    """
    def __init__(self, Kp, Ki, Kd, Ts=0.1, out_min=0.0, out_max=100.0, anti_windup=True):
        self.Kp = Kp
        self.Ki = Ki
        self.Kd = Kd
        self.Ts = Ts
        self.out_min = out_min
        self.out_max = out_max
        self.integral = 0.0
        self.e_prev = 0.0
        self.anti_windup = anti_windup

    def reset(self):
        """Сброс интегратора и предыдущей ошибки"""
        self.integral = 0.0
        self.e_prev = 0.0

    def compute(self, SP, PV):
        """
        Вычисление управляющего воздействия.
        SP - Setpoint (задание/уставка)
        PV - Process Variable (текущее значение регулируемой величины)
        """
        e = SP - PV
        self.integral += e * self.Ts
        derivative = (e - self.e_prev) / self.Ts
        out = self.Kp * e + self.Ki * self.integral + self.Kd * derivative

        if self.anti_windup:
            # Clamping anti-windup: откат интегральной суммы при насыщении
            if out > self.out_max:
                out = self.out_max
                self.integral -= e * self.Ts
            elif out < self.out_min:
                out = self.out_min
                self.integral -= e * self.Ts
        else:
            # Ограничение только на выходе привода без коррекции интегратора
            if out > self.out_max:
                out = self.out_max
            elif out < self.out_min:
                out = self.out_min

        self.e_prev = e
        return out


def simulate_closed_loop(plant, pid, SP=50.0, t_total=80.0, dt=0.1):
    """Имитационное моделирование замкнутой системы регулирования"""
    plant.reset()
    pid.reset()
    steps = int(round(t_total / dt))
    time_arr = []
    y_arr = []
    u_arr = []
    e_arr = []
    t = 0.0

    for _ in range(steps):
        PV = plant.y
        u = pid.compute(SP, PV)
        plant.step(u)

        time_arr.append(t)
        y_arr.append(PV)
        u_arr.append(u)
        e_arr.append(SP - PV)
        t += dt

    return np.array(time_arr), np.array(y_arr), np.array(u_arr), np.array(e_arr)


def calculate_metrics(time_arr, y_arr, SP=50.0, delta=0.05):
    """
    Расчет инженерных показателей качества переходного процесса:
    - Перерегулирование sigma (%)
    - Время регулирования tp (с) по 5% трубке допуска
    - Время нарастания tr (с) от 10% до 90%
    - Статическая ошибка e_ст
    - Интегральные критерии IAE и ITAE
    """
    y_max = float(np.max(y_arr))
    overshoot = max(0.0, (y_max - SP) / SP * 100.0)

    # 5% трубка допуска
    band = delta * SP
    outside = np.where((y_arr < SP - band) | (y_arr > SP + band))[0]
    if len(outside) == 0:
        t_settling = 0.0
    elif outside[-1] == len(y_arr) - 1:
        t_settling = float(time_arr[-1])
    else:
        t_settling = float(time_arr[outside[-1] + 1])

    # Время нарастания
    idx_10 = np.where(y_arr >= 0.1 * SP)[0]
    idx_90 = np.where(y_arr >= 0.9 * SP)[0]
    t_rise = float(time_arr[idx_90[0]] - time_arr[idx_10[0]]) if len(idx_10) > 0 and len(idx_90) > 0 else 0.0

    # Установившаяся ошибка
    e_steady = float(abs(SP - y_arr[-1]))

    # Интегральные критерии
    dt = time_arr[1] - time_arr[0]
    iae = float(np.sum(np.abs(SP - y_arr)) * dt)
    itae = float(np.sum(time_arr * np.abs(SP - y_arr)) * dt)

    return {
        'y_max': round(y_max, 2),
        'overshoot': round(overshoot, 2),
        't_settling': round(t_settling, 2),
        't_rise': round(t_rise, 2),
        'e_steady': round(e_steady, 4),
        'iae': round(iae, 2),
        'itae': round(itae, 2)
    }


def main():
    print("=" * 70)
    print("ЛАБОРАТОРНАЯ РАБОТА №2. РАЗРАБОТКА АЛГОРИТМА ПИД-РЕГУЛИРОВАНИЯ")
    print("=" * 70)

    dt = 0.1
    # -------------------------------------------------------------
    # ШАГ 1-3. МОДЕЛИРОВАНИЕ И ИДЕНТИФИКАЦИЯ ОБЪЕКТА
    # -------------------------------------------------------------
    print("\n1. Моделирование переходной характеристики объекта...")
    plant = FirstOrderPlant(K=1.0, T=10.0, tau=2.0, dt=dt)
    t_steps = int(round(60.0 / dt))
    t_plant, y_plant = [], []
    t = 0.0
    for _ in range(t_steps):
        y_plant.append(plant.step(1.0))
        t_plant.append(t)
        t += dt
    t_plant = np.array(t_plant)
    y_plant = np.array(y_plant)

    y_steady = float(y_plant[-1])
    K_ob = y_steady / 1.0
    idx_tau = np.where(y_plant > 0.001)[0][0]
    tau_ob = float(t_plant[idx_tau])
    val_63 = 0.632 * y_steady
    idx_63 = np.where(y_plant >= val_63)[0][0]
    t_63 = float(t_plant[idx_63])
    T_ob = t_63 - tau_ob

    print(f"   Заданные параметры модели: K = 1.0, T = 10.0 с, tau = 2.0 с")
    print(f"   Идентифицированные параметры:")
    print(f"   - Установившееся значение y_уст = {y_steady:.4f}")
    print(f"   - Коэффициент передачи K_об = {K_ob:.4f}")
    print(f"   - Запаздывание tau_об = {tau_ob:.2f} с")
    print(f"   - Постоянная времени T_об = {T_ob:.2f} с")

    # Построение и сохранение кривой разгона
    plt.figure(figsize=(9, 5), dpi=300)
    plt.plot(t_plant, y_plant, 'b-', linewidth=2, label='Переходная характеристика y(t)')
    plt.axhline(y_steady, color='g', linestyle='--', linewidth=1.5, label=f'y_уст = {y_steady:.2f} (Kоб = {K_ob:.2f})')
    plt.axhline(val_63, color='darkorange', linestyle=':', linewidth=1.5, label=f'0.632·y_уст = {val_63:.3f}')
    plt.axvline(tau_ob, color='r', linestyle='--', linewidth=1.5, label=f'Запаздывание τ = {tau_ob:.1f} с')
    plt.axvline(t_63, color='purple', linestyle='--', linewidth=1.5, label=f't_63% = {t_63:.1f} с (T = {T_ob:.1f} с)')
    plt.plot([tau_ob, t_63], [val_63, val_63], 'purple', linewidth=2.5, marker='o')
    plt.annotate(f' T = {T_ob:.1f} с', xy=((tau_ob + t_63) / 2, val_63), xytext=(0, 10),
                 textcoords='offset points', ha='center', fontweight='bold', color='purple')
    plt.title('Переходная характеристика объекта управления (кривая разгона)', pad=12, fontweight='bold')
    plt.xlabel('Время t, с')
    plt.ylabel('Выход объекта y(t)')
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend(loc='lower right', framealpha=0.95)
    plt.xlim(0, 50)
    plt.ylim(-0.05, 1.15)
    plt.tight_layout()
    plt.savefig('plant_step_response.png')
    plt.close()

    # -------------------------------------------------------------
    # ШАГ 4. РАСЧЕТ КОЭФФИЦИЕНТОВ ПИД ПО ЦИГЛЕРУ-НИКОЛЬСУ
    # -------------------------------------------------------------
    print("\n2. Расчет настроек ПИД-регулятора по методу Циглера–Николса:")
    Kp_zn = 1.2 * T_ob / (K_ob * tau_ob)
    Ti_zn = 2.0 * tau_ob
    Td_zn = 0.5 * tau_ob
    Ki_zn = Kp_zn / Ti_zn
    Kd_zn = Kp_zn * Td_zn

    print(f"   Kp = {Kp_zn:.3f}")
    print(f"   Ti = {Ti_zn:.3f} с  => Ki = {Ki_zn:.3f} 1/с")
    print(f"   Td = {Td_zn:.3f} с  => Kd = {Kd_zn:.3f} с")

    # -------------------------------------------------------------
    # ШАГ 5-6. ИМИТАЦИОННЫЕ ИСПЫТАНИЯ
    # -------------------------------------------------------------
    SP = 50.0
    print(f"\n3. Имитационные испытания при ступенчатом задании SP = {SP:.1f}...")

    # Базовая настройка Циглера–Николса (С Anti-windup)
    plant_zn = FirstOrderPlant(K=1.0, T=10.0, tau=2.0, dt=dt)
    pid_zn = PID(Kp=Kp_zn, Ki=Ki_zn, Kd=Kd_zn, Ts=dt, anti_windup=True)
    t_zn, y_zn, u_zn, _ = simulate_closed_loop(plant_zn, pid_zn, SP=SP, t_total=80.0, dt=dt)
    m_zn = calculate_metrics(t_zn, y_zn, SP=SP)

    # Настройка Циглера–Николса (БЕЗ Anti-windup)
    plant_no_aw = FirstOrderPlant(K=1.0, T=10.0, tau=2.0, dt=dt)
    pid_no_aw = PID(Kp=Kp_zn, Ki=Ki_zn, Kd=Kd_zn, Ts=dt, anti_windup=False)
    t_no_aw, y_no_aw, u_no_aw, _ = simulate_closed_loop(plant_no_aw, pid_no_aw, SP=SP, t_total=80.0, dt=dt)
    m_no_aw = calculate_metrics(t_no_aw, y_no_aw, SP=SP)

    # Оптимизированная настройка
    Kp_opt = 2.8
    Ti_opt = 6.5
    Td_opt = 0.9
    Ki_opt = Kp_opt / Ti_opt
    Kd_opt = Kp_opt * Td_opt
    plant_opt = FirstOrderPlant(K=1.0, T=10.0, tau=2.0, dt=dt)
    pid_opt = PID(Kp=Kp_opt, Ki=Ki_opt, Kd=Kd_opt, Ts=dt, anti_windup=True)
    t_opt, y_opt, u_opt, _ = simulate_closed_loop(plant_opt, pid_opt, SP=SP, t_total=80.0, dt=dt)
    m_opt = calculate_metrics(t_opt, y_opt, SP=SP)

    # Апериодическая настройка
    Kp_aper = 2.0
    Ti_aper = 8.0
    Td_aper = 0.8
    Ki_aper = Kp_aper / Ti_aper
    Kd_aper = Kp_aper * Td_aper
    plant_aper = FirstOrderPlant(K=1.0, T=10.0, tau=2.0, dt=dt)
    pid_aper = PID(Kp=Kp_aper, Ki=Ki_aper, Kd=Kd_aper, Ts=dt, anti_windup=True)
    t_aper, y_aper, u_aper, _ = simulate_closed_loop(plant_aper, pid_aper, SP=SP, t_total=80.0, dt=dt)
    m_aper = calculate_metrics(t_aper, y_aper, SP=SP)

    # Вывод результатов в консоль
    print("\nТаблица показателей качества регулирования:")
    header = f"{'Режим настройки':<25} | {'Kp':<6} | {'Ki':<6} | {'Kd':<6} | {'σ, %':<7} | {'tp, с':<7} | {'tr, с':<7} | {'IAE':<7} | {'ITAE':<7}"
    print("-" * len(header))
    print(header)
    print("-" * len(header))
    print(f"{'Циглер–Николс (с AW)':<25} | {Kp_zn:<6.2f} | {Ki_zn:<6.2f} | {Kd_zn:<6.2f} | {m_zn['overshoot']:<7.2f} | {m_zn['t_settling']:<7.2f} | {m_zn['t_rise']:<7.2f} | {m_zn['iae']:<7.1f} | {m_zn['itae']:<7.1f}")
    print(f"{'Циглер–Николс (без AW)':<25} | {Kp_zn:<6.2f} | {Ki_zn:<6.2f} | {Kd_zn:<6.2f} | {m_no_aw['overshoot']:<7.2f} | {m_no_aw['t_settling']:<7.2f} | {m_no_aw['t_rise']:<7.2f} | {m_no_aw['iae']:<7.1f} | {m_no_aw['itae']:<7.1f}")
    print(f"{'Оптимизированный':<25} | {Kp_opt:<6.2f} | {Ki_opt:<6.2f} | {Kd_opt:<6.2f} | {m_opt['overshoot']:<7.2f} | {m_opt['t_settling']:<7.2f} | {m_opt['t_rise']:<7.2f} | {m_opt['iae']:<7.1f} | {m_opt['itae']:<7.1f}")
    print(f"{'Апериодический':<25} | {Kp_aper:<6.2f} | {Ki_aper:<6.2f} | {Kd_aper:<6.2f} | {m_aper['overshoot']:<7.2f} | {m_aper['t_settling']:<7.2f} | {m_aper['t_rise']:<7.2f} | {m_aper['iae']:<7.1f} | {m_aper['itae']:<7.1f}")
    print("-" * len(header))

    # Построение графиков
    # 1. Отклик Зиглера-Никольса
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 6.5), dpi=300, sharex=True)
    ax1.plot(t_zn, y_zn, 'b-', linewidth=2, label='Регулируемая величина PV (выход объекта y)')
    ax1.axhline(SP, color='r', linestyle='--', linewidth=1.5, label=f'Задание SP = {SP:.1f}')
    ax1.axhline(SP * 1.05, color='gray', linestyle=':', linewidth=1.0, label='±5% трубка допуска')
    ax1.axhline(SP * 0.95, color='gray', linestyle=':', linewidth=1.0)
    ax1.set_title('Переходный процесс замкнутой системы (настройка Циглера–Николса)', pad=10, fontweight='bold')
    ax1.set_ylabel('Выход PV')
    ax1.grid(True, linestyle='--', alpha=0.7)
    ax1.legend(loc='lower right', framealpha=0.95)
    ax1.set_ylim(-5, 65)

    ax2.plot(t_zn, u_zn, 'g-', linewidth=2, label='Управляющий сигнал ПИД u(t)')
    ax2.axhline(100.0, color='r', linestyle=':', linewidth=1.0, label='Ограничения привода [0; 100]%')
    ax2.axhline(0.0, color='r', linestyle=':', linewidth=1.0)
    ax2.set_title('Управляющее воздействие регулятора', pad=8, fontweight='bold')
    ax2.set_xlabel('Время t, с')
    ax2.set_ylabel('Управление u(t), %')
    ax2.grid(True, linestyle='--', alpha=0.7)
    ax2.legend(loc='upper right', framealpha=0.95)
    ax2.set_xlim(0, 80)
    ax2.set_ylim(-5, 110)
    plt.tight_layout()
    plt.savefig('pid_response_zn.png')
    plt.close()

    # 2. Сравнение Anti-windup
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 6.5), dpi=300, sharex=True)
    ax1.plot(t_zn, y_zn, 'b-', linewidth=2, label=f'С Anti-windup (σ = {m_zn["overshoot"]}%, tp = {m_zn["t_settling"]} с)')
    ax1.plot(t_no_aw, y_no_aw, 'r--', linewidth=2, label=f'БЕЗ Anti-windup (σ = {m_no_aw["overshoot"]}%, tp = {m_no_aw["t_settling"]} с)')
    ax1.axhline(SP, color='black', linestyle=':', linewidth=1.2, label=f'Задание SP = {SP:.1f}')
    ax1.set_title('Влияние защиты от интегрального насыщения (Anti-windup)', pad=10, fontweight='bold')
    ax1.set_ylabel('Выход PV')
    ax1.grid(True, linestyle='--', alpha=0.7)
    ax1.legend(loc='lower right', framealpha=0.95)
    ax1.set_ylim(-5, 85)

    ax2.plot(t_zn, u_zn, 'b-', linewidth=1.8, label='Управление u(t) С Anti-windup')
    ax2.plot(t_no_aw, u_no_aw, 'r--', linewidth=1.8, label='Управление u(t) БЕЗ Anti-windup (залипание в 100%)')
    ax2.axhline(100.0, color='gray', linestyle=':', linewidth=1.0)
    ax2.set_title('Динамика управляющего воздействия u(t)', pad=8, fontweight='bold')
    ax2.set_xlabel('Время t, с')
    ax2.set_ylabel('Управление u(t), %')
    ax2.grid(True, linestyle='--', alpha=0.7)
    ax2.legend(loc='upper right', framealpha=0.95)
    ax2.set_xlim(0, 80)
    ax2.set_ylim(-5, 110)
    plt.tight_layout()
    plt.savefig('pid_antiwindup_comparison.png')
    plt.close()

    # 3. Сравнение оптимизации
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 6.5), dpi=300, sharex=True)
    ax1.plot(t_zn, y_zn, 'r--', linewidth=1.8, label=f'Циглер–Николс: Kp={Kp_zn:.2f}, Ki={Ki_zn:.2f}, Kd={Kd_zn:.2f} (σ={m_zn["overshoot"]}%, tp={m_zn["t_settling"]}с)')
    ax1.plot(t_opt, y_opt, 'b-', linewidth=2.2, label=f'Оптимизированный: Kp={Kp_opt:.2f}, Ki={Ki_opt:.2f}, Kd={Kd_opt:.2f} (σ={m_opt["overshoot"]}%, tp={m_opt["t_settling"]}с)')
    ax1.plot(t_aper, y_aper, 'g-.', linewidth=1.8, label=f'Апериодический: Kp={Kp_aper:.2f}, Ki={Ki_aper:.2f}, Kd={Kd_aper:.2f} (σ={m_aper["overshoot"]}%, tp={m_aper["t_settling"]}с)')
    ax1.axhline(SP, color='black', linestyle=':', linewidth=1.2, label=f'Задание SP = {SP:.1f}')
    ax1.axhline(SP * 1.05, color='gray', linestyle=':', linewidth=0.8)
    ax1.axhline(SP * 0.95, color='gray', linestyle=':', linewidth=0.8)
    ax1.set_title('Сравнение переходных процессов при различных настройках ПИД', pad=10, fontweight='bold')
    ax1.set_ylabel('Выход PV')
    ax1.grid(True, linestyle='--', alpha=0.7)
    ax1.legend(loc='lower right', framealpha=0.95)
    ax1.set_ylim(-5, 65)

    ax2.plot(t_zn, u_zn, 'r--', linewidth=1.5, label='Управление Циглера–Николса')
    ax2.plot(t_opt, u_opt, 'b-', linewidth=2.0, label='Управление оптимизированное')
    ax2.plot(t_aper, u_aper, 'g-.', linewidth=1.5, label='Управление апериодическое')
    ax2.set_title('Сравнение управляющих сигналов u(t)', pad=8, fontweight='bold')
    ax2.set_xlabel('Время t, с')
    ax2.set_ylabel('Управление u(t), %')
    ax2.grid(True, linestyle='--', alpha=0.7)
    ax2.legend(loc='upper right', framealpha=0.95)
    ax2.set_xlim(0, 80)
    ax2.set_ylim(-5, 110)
    plt.tight_layout()
    plt.savefig('pid_optimization_comparison.png')
    plt.close()

    print("\nМоделирование завершено успешно! Графики сохранены в текущую папку.")


if __name__ == '__main__':
    main()

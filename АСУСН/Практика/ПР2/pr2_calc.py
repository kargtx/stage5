# -*- coding: utf-8 -*-
"""
Практическое занятие № 2
Дисциплина: «Автоматизированные системы специального назначения» (АССН)
Тема: Расчёт показателей надёжности системы с резервированием
Вариант 2: Система противоаварийной защиты (ПАЗ) по мажоритарной схеме «2 из 3» (2oo3)
Студенты: Гайнанов Д.И., Гарифуллин К.Р., группа ЭАС-514С
Преподаватель: Антонов В.В.
Уфа, УУНиТ, 2026 г.
"""

import os
import numpy as np
import matplotlib.pyplot as plt

# Настройка параметров шрифта для графиков
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 10
plt.rcParams['axes.titlesize'] = 12
plt.rcParams['axes.labelsize'] = 11
plt.rcParams['xtick.labelsize'] = 10
plt.rcParams['ytick.labelsize'] = 10
plt.rcParams['legend.fontsize'] = 10
plt.rcParams['figure.titlesize'] = 13


def main():
    print("=" * 75)
    print("ПРАКТИЧЕСКОЕ ЗАНЯТИЕ № 2. РАСЧЁТ НАДЁЖНОСТИ СИСТЕМЫ С РЕЗЕРВИРОВАНИЕМ")
    print("ВАРИАНТ № 2: Мажоритарная система противоаварийной защиты «2 из 3» (2oo3)")
    print("=" * 75)

    # Исходные данные по Варианту 2:
    lam = 5e-5       # Интенсивность отказов одного канала, 1/ч
    t_work = 5000.0  # Заданное время непрерывной работы, ч
    t_max = 20000.0  # Верхняя граница временного диапазона анализа, ч

    print(f"\nИСХОДНЫЕ ДАННЫЕ:")
    print(f"• Интенсивность отказов одного канала: lambda = {lam:.1e} 1/ч")
    print(f"• Заданное время непрерывной работы: t = {t_work:.0f} ч")
    print(f"• Структура резервирования: мажоритарная схема «2 из 3» (2oo3, TMR)")

    # 1. Вероятность безотказной работы одиночного канала
    P_channel_5000 = np.exp(-lam * t_work)
    Q_channel_5000 = 1.0 - P_channel_5000

    print(f"\n1. ВЕРОЯТНОСТЬ БЕЗОТКАЗНОЙ РАБОТЫ ОДИНОЧНОГО КАНАЛА (t = {t_work:.0f} ч):")
    print(f"   P_кан(t) = exp(-lambda*t) = exp(-{lam * t_work:.2f})")
    print(f"   P_кан(5000) = {P_channel_5000:.6f} ~= {P_channel_5000:.4f}")
    print(f"   Q_кан(5000) = 1 - P_кан(5000) = {Q_channel_5000:.6f} ~= {Q_channel_5000:.4f}")

    # 2. Вероятность безотказной работы мажоритарной системы 2 из 3
    # P_23(t) = 3 * P(t)^2 - 2 * P(t)^3
    P_23_5000 = 3.0 * (P_channel_5000 ** 2) - 2.0 * (P_channel_5000 ** 3)
    Q_23_5000 = 1.0 - P_23_5000

    print(f"\n2. ВЕРОЯТНОСТЬ БЕЗОТКАЗНОЙ РАБОТЫ СИСТЕМЫ «2 ИЗ 3» (t = {t_work:.0f} ч):")
    print(f"   P_2/3(t) = 3*P^2(t) - 2*P^3(t)")
    print(f"   P_2/3(5000) = 3*({P_channel_5000:.4f})^2 - 2*({P_channel_5000:.4f})^3")
    print(f"   P_2/3(5000) = {P_23_5000:.6f} ~= {P_23_5000:.4f}")
    print(f"   Q_2/3(5000) = 1 - P_2/3(5000) = {Q_23_5000:.6f} ~= {Q_23_5000:.4f}")

    # 3. Средняя наработка на отказ (T0)
    T0_channel = 1.0 / lam
    T0_23 = (5.0 / 6.0) / lam

    print(f"\n3. СРЕДНЯЯ НАРАБОТКА НА ОТКАЗ СИСТЕМЫ (T0 / MTTF):")
    print(f"   Одиночный канал: T0_кан = 1 / lambda = {T0_channel:.0f} ч")
    print(f"   Система «2 из 3»: T0_2/3 = 5 / (6*lambda) = {T0_23:.2f} ч ~= {round(T0_23)} ч")

    # 4. Выигрыш в надежности
    ratio_P = P_23_5000 / P_channel_5000
    gain_P_percent = (ratio_P - 1.0) * 100.0
    ratio_Q = Q_channel_5000 / Q_23_5000
    reduction_Q_percent = (1.0 - Q_23_5000 / Q_channel_5000) * 100.0

    print(f"\n4. ОЦЕНКА ПОВЫШЕНИЯ НАДЁЖНОСТИ:")
    print(f"   • Отношение вероятностей безотказности: P_2/3 / P_кан = {ratio_P:.4f} (+{gain_P_percent:.2f}%)")
    print(f"   • Коэффициент снижения вероятности отказа: K_Q = Q_кан / Q_2/3 = {ratio_Q:.2f} раза")
    print(f"   • Снижение риска аварийного отказа: на {reduction_Q_percent:.2f}%")

    # Критическая точка эквивалентности
    t_crit = np.log(2.0) / lam
    print(f"\n   Критическая точка эквивалентности (P = 0.5): t_кр = ln(2) / lambda = {t_crit:.1f} ч")
    print(f"   При t < {t_crit:.0f} ч мажоритарная система надежнее одиночного канала.")

    # 5. Коэффициент готовности
    Tv = 24.0
    Kg_channel = T0_channel / (T0_channel + Tv)
    Kg_23 = T0_23 / (T0_23 + Tv)
    print(f"\n   Коэффициент готовности при Tв = {Tv:.0f} ч:")
    print(f"   • Одиночный канал: Kг_кан = {Kg_channel:.5f}")
    print(f"   • Система «2 из 3»: Kг_2/3 = {Kg_23:.5f}")

    # -------------------------------------------------------------
    # 5. ПОСТРОЕНИЕ ГРАФИКОВ
    # -------------------------------------------------------------
    print(f"\n5. Построение графиков зависимостей P(t) и Q(t)...")
    t_arr = np.linspace(0, t_max, 1000)
    P_ch_arr = np.exp(-lam * t_arr)
    P_23_arr = 3.0 * (P_ch_arr ** 2) - 2.0 * (P_ch_arr ** 3)

    # График 1: P(t)
    plt.figure(figsize=(9, 5.5), dpi=300)
    plt.plot(t_arr, P_ch_arr, 'b--', linewidth=2.0, label='Одиночный канал: P(t) = exp(-λ·t)')
    plt.plot(t_arr, P_23_arr, 'r-', linewidth=2.5, label='Мажоритарная система «2 из 3»: P_2/3(t) = 3P² - 2P³')
    plt.plot([t_work, t_work], [0, P_23_5000], 'gray', linestyle=':', linewidth=1.2)
    plt.plot(t_work, P_channel_5000, 'bo', markersize=7)
    plt.plot(t_work, P_23_5000, 'ro', markersize=7)
    plt.annotate(f'P_кан(5000) = {P_channel_5000:.4f}', xy=(t_work, P_channel_5000), xytext=(t_work + 500, P_channel_5000 - 0.08),
                 arrowprops=dict(arrowstyle='->', color='blue', lw=1.2), fontweight='bold', color='blue')
    plt.annotate(f'P_2/3(5000) = {P_23_5000:.4f}\n(+{gain_P_percent:.1f}%)', xy=(t_work, P_23_5000), xytext=(t_work + 500, P_23_5000 + 0.04),
                 arrowprops=dict(arrowstyle='->', color='red', lw=1.2), fontweight='bold', color='red')
    plt.plot(t_crit, 0.5, 'go', markersize=8)
    plt.annotate(f'Точка эквивалентности\nt_кр = {t_crit:.0f} ч (P = 0.5)', xy=(t_crit, 0.5), xytext=(t_crit - 4500, 0.35),
                 arrowprops=dict(arrowstyle='->', color='green', lw=1.2), fontweight='bold', color='darkgreen')
    plt.axvspan(0, t_crit, color='yellow', alpha=0.1, label=f'Зона выигрыша TMR (t < {t_crit:.0f} ч)')
    plt.title('Вероятность безотказной работы P(t) одиночного канала и системы «2 из 3»', pad=12, fontweight='bold')
    plt.xlabel('Время непрерывной работы t, ч')
    plt.ylabel('Вероятность безотказной работы P(t)')
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend(loc='upper right', framealpha=0.95)
    plt.xlim(0, t_max)
    plt.ylim(0, 1.05)
    plt.tight_layout()
    plt.savefig('reliability_2oo3.png')
    plt.close()

    # График 2: Q(t) и кратность
    Q_ch_arr = 1.0 - P_ch_arr
    Q_23_arr = 1.0 - P_23_arr
    ratio_Q_arr = np.zeros_like(t_arr)
    mask = t_arr > 10
    ratio_Q_arr[mask] = Q_ch_arr[mask] / Q_23_arr[mask]

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(9.5, 6.5), dpi=300, sharex=True)
    ax1.plot(t_arr, Q_ch_arr, 'b--', linewidth=2.0, label='Вероятность отказа канала Q_кан(t)')
    ax1.plot(t_arr, Q_23_arr, 'r-', linewidth=2.2, label='Вероятность отказа системы «2 из 3» Q_2/3(t)')
    ax1.plot(t_work, Q_channel_5000, 'bo', markersize=6)
    ax1.plot(t_work, Q_23_5000, 'ro', markersize=6)
    ax1.annotate(f'Q_кан = {Q_channel_5000:.4f}', xy=(t_work, Q_channel_5000), xytext=(t_work + 600, Q_channel_5000 + 0.05),
                 arrowprops=dict(arrowstyle='->', color='blue'), fontweight='bold', color='blue')
    ax1.annotate(f'Q_2/3 = {Q_23_5000:.4f} (снижение в 1.77 раза)', xy=(t_work, Q_23_5000), xytext=(t_work + 600, Q_23_5000 - 0.07),
                 arrowprops=dict(arrowstyle='->', color='red'), fontweight='bold', color='red')
    ax1.set_title('Вероятность отказа Q(t) = 1 - P(t)', pad=10, fontweight='bold')
    ax1.set_ylabel('Вероятность отказа Q(t)')
    ax1.grid(True, linestyle='--', alpha=0.7)
    ax1.legend(loc='upper left', framealpha=0.95)
    ax1.set_ylim(-0.02, 1.0)

    mask2 = t_arr >= 200
    ax2.plot(t_arr[mask2], ratio_Q_arr[mask2], 'purple', linewidth=2.2, label='Кратность снижения опасности отказа K_Q = Q_кан / Q_2/3')
    ax2.axhline(1.0, color='gray', linestyle=':', label='K_Q = 1 (граница эффективности)')
    ax2.axvline(t_crit, color='green', linestyle='--', label=f't_кр = {t_crit:.0f} ч')
    ax2.plot(t_work, ratio_Q, 's', color='darkred', markersize=7)
    ax2.annotate(f'K_Q(5000) = {ratio_Q:.2f} раза', xy=(t_work, ratio_Q), xytext=(t_work + 1000, ratio_Q + 1.5),
                 arrowprops=dict(arrowstyle='->', color='darkred'), fontweight='bold', color='darkred')
    ax2.set_title('Коэффициент снижения вероятности отказа (фактор снижения риска RRF)', pad=8, fontweight='bold')
    ax2.set_xlabel('Время непрерывной работы t, ч')
    ax2.set_ylabel('Кратность K_Q')
    ax2.grid(True, linestyle='--', alpha=0.7)
    ax2.legend(loc='upper right', framealpha=0.95)
    ax2.set_xlim(0, t_max)
    ax2.set_ylim(0, 15)
    plt.tight_layout()
    plt.savefig('unreliability_comparison.png')
    plt.close()

    print("\nРасчет завершен успешно! Графики сохранены в текущей папке.")


if __name__ == '__main__':
    main()

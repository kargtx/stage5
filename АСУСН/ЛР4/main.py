# -*- coding: utf-8 -*-
"""
Основной скрипт проведения эксперимента по регламенту Лабораторной работы №4:
«Реализация алгоритма диагностики и обнаружения отказа датчика»

Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.
"""

import os
import sys
import argparse
import time
import numpy as np
import matplotlib.pyplot as plt

# Импорт разработанных модулей
from sensor_simulator import SensorSimulator
from diagnostics import SensorDiagnostics, MedianFilter, KalmanFilter1D

# Настройка шрифтов matplotlib
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Segoe UI', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False


def run_experiment(realtime=False):
    """
    Запуск регламентного вычислительного эксперимента.
    
    :param realtime: если True, выдерживается реальная задержка time.sleep(dt).
                     если False, симуляция выполняется ускоренно (синхронное виртуальное время).
    """
    print("=" * 70)
    print("  ЛАБОРАТОРНАЯ РАБОТА №4. АСУСН")
    print("  Диагностика и обнаружение отказов датчика технологических параметров")
    print(f"  Режим выполнения: {'Реальное время (wall-clock)' if realtime else 'Ускоренное моделирование (virtual time)'}")
    print("=" * 70)

    # 1. Инициализация подсистем
    sensor = SensorSimulator(nominal=50.0, noise=0.5, min_val=-40.0, max_val=150.0, drift_rate=0.5)
    diag = SensorDiagnostics(min_val=-40.0, max_val=150.0, max_rate=15.0, stuck_threshold=10)
    median = MedianFilter(window_size=5)

    # 2. Параметры эксперимента по методическим указаниям
    duration = 60.0         # Длительность эксперимента (с)
    dt = 0.1                # Период дискретизации опроса датчика (с)

    # Расписание внесения регламентированных отказов:
    # (время_начала, тип_отказа, длительность)
    schedule = [
        (15.0, 'drift', 10.0),   # На 15-й секунде – дрейф нуля на 10 сек (скорость 0.5 °C/с)
        (35.0, 'open', 5.0),     # На 35-й секунде – обрыв цепи на 5 сек (-50 °C)
        (45.0, 'spike', 2.0),    # На 45-й секунде – импульсный выброс (+/- 30 °C)
    ]

    # Журналы измерений и событий
    times = []
    values_raw = []
    values_filtered = []
    faults_log = []

    start_real_time = time.time()
    next_fault_idx = 0
    num_steps = int(duration / dt) + 1

    for step in range(num_steps):
        t = round(step * dt, 3)
        if t > duration:
            break

        # Проверка расписания внесения отказов
        if next_fault_idx < len(schedule):
            fault_time, fault_type, fault_dur = schedule[next_fault_idx]
            if t >= fault_time:
                sensor.set_fault(fault_type, duration=fault_dur, current_time=t)
                print(f"[{t:5.1f} s] >>> ВНЕСЁН ОТКАЗ: {fault_type.upper()} (длительность {fault_dur} с)")
                next_fault_idx += 1

        # Опрос датчика
        raw = sensor.read(sim_time=t)
        filtered = median.filter(raw)

        # Диагностический контроль отфильтрованного сигнала
        problems = diag.check(filtered, t)
        if problems:
            print(f"[{t:5.1f} s] !!! ОБНАРУЖЕН ОТКАЗ: {problems}, значение = {filtered:6.2f} °C")
            faults_log.append((t, list(problems), filtered))

        # Накопление временных рядов
        times.append(t)
        values_raw.append(raw)
        values_filtered.append(filtered)

        if realtime:
            time.sleep(dt)

    total_sim_time = time.time() - start_real_time
    print("-" * 70)
    print(f"Моделирование завершено за {total_sim_time:.2f} сек. Всего тактов: {len(times)}")
    print(f"Всего зафиксировано событий отказа: {len(faults_log)}")

    # 3. Визуализация результатов (согласно разделу 4.5 методички)
    fig, ax = plt.subplots(2, 1, figsize=(12, 8), dpi=150)

    # Верхний график: сигналы и технологические границы
    ax[0].plot(times, values_raw, 'b-', alpha=0.45, linewidth=1.0, label='Сырой сигнал (с шумом)')
    ax[0].plot(times, values_filtered, 'r-', linewidth=1.5, label='Медианный фильтр (окно N=5)')
    ax[0].axhline(y=150.0, color='darkred', linestyle='--', linewidth=1.2, label='Верхняя граница (150 °C)')
    ax[0].axhline(y=-40.0, color='darkblue', linestyle='--', linewidth=1.2, label='Нижняя граница (-40 °C)')
    ax[0].set_title('Реализация алгоритма диагностики и обнаружения отказа датчика', fontsize=12, fontweight='bold', pad=10)
    ax[0].set_xlabel('Время, с', fontsize=10)
    ax[0].set_ylabel('Температура, °C', fontsize=10)
    ax[0].legend(loc='upper right', framealpha=0.9)
    ax[0].grid(True, linestyle=':', alpha=0.7)

    # Отметки моментов обнаружения отказов
    for t_fault, probs, val in faults_log:
        ax[0].axvline(x=t_fault, color='red', alpha=0.35, linestyle=':')

    # Нижний график: бинарный флаг диагностированного отказа
    fault_flags = [1 if any(abs(t - f[0]) < 1e-4 for f in faults_log) else 0 for t in times]
    ax[1].step(times, fault_flags, 'r-', where='mid', linewidth=1.3, label='Флаг обнаруженного отказа')
    ax[1].set_xlabel('Время, с', fontsize=10)
    ax[1].set_ylabel('Флаг отказа (0/1)', fontsize=10)
    ax[1].set_ylim(-0.1, 1.2)
    ax[1].set_yticks([0, 1])
    ax[1].set_yticklabels(['Норма (0)', 'Отказ (1)'])
    ax[1].grid(True, linestyle=':', alpha=0.7)
    ax[1].legend(loc='upper right')

    plt.tight_layout()
    output_png = os.path.join(os.path.dirname(__file__), 'diagnostics_result.png')
    plt.savefig(output_png, dpi=150)
    print(f"График сохранён в файл: {output_png}")

    # Сводная таблица журнала отказов
    print("\nСводный журнал зарегистрированных отказов:")
    print("  Время (с) | Обнаруженный дефект        | Измеренное значение")
    print("  " + "-" * 55)
    for t, probs, val in faults_log:
        print(f"   {t:6.1f}  | {', '.join(probs):<25} | {val:8.2f} °C")

    return times, values_raw, values_filtered, faults_log


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Запуск лабораторной работы №4 АСУСН")
    parser.add_argument('--realtime', action='store_true', help="Выполнять симуляцию с реальной секундной задержкой")
    args = parser.parse_args()

    run_experiment(realtime=args.realtime)

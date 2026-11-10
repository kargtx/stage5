# -*- coding: utf-8 -*-
"""
Модуль программного клиента Modbus TCP (эмулятор SCADA-станции).
Дисциплина: АСУСН, Лабораторная работа №5.
Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.

Выполняет:
- Периодический циклический опрос регистров ПЛК (Holding Registers 0..2, Discrete Input 0)
- Автоматическое диспетчерское управление насосным агрегатом (Write Coil 0 / Write Coil 1)
- Построение и сохранение графических трендов параметров процесса (modbus_result.png)
"""

import os
import sys
import time
import argparse
import matplotlib.pyplot as plt
from pymodbus.client import ModbusTcpClient

# Настройка шрифтов matplotlib
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Segoe UI', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False


def run_modbus_client(host='127.0.0.1', port=5020, duration=30.0, poll_interval=0.5, save_plot=True):
    """
    Запуск цикла опроса и управления Modbus TCP клиента.
    
    :param host: IP адрес Modbus TCP сервера
    :param port: TCP порт сервера
    :param duration: длительность опроса (с)
    :param poll_interval: период опроса регистров (с)
    :param save_plot: сохранять ли графический файл modbus_result.png
    """
    print("=" * 70)
    print("  ЭМУЛЯТОР SCADA-КЛИЕНТА MODBUS TCP")
    print(f"  Подключение к серверу: {host}:{port}")
    print(f"  Длительность сессии: {duration} с (шаг опроса: {poll_interval} с)")
    print("=" * 70)

    client = ModbusTcpClient(host, port=port)
    if not client.connect():
        print(f"ОШИБКА: Не удалось подключиться к серверу {host}:{port}!")
        return None

    # Журналы измерений
    times = []
    temps = []
    pressures = []
    levels = []
    alarms = []
    commands_log = []

    start_time = time.time()
    try:
        while time.time() - start_time < duration:
            t = round(time.time() - start_time, 2)

            # 1. Чтение Holding Registers (Функция 0x03)
            # HR[0] - Температура x10, HR[1] - Давление x100, HR[2] - Уровень
            try:
                hr_response = client.read_holding_registers(0, count=3)
            except TypeError:
                hr_response = client.read_holding_registers(0, 3)

            if not hr_response.isError():
                temp = hr_response.registers[0] / 10.0
                pressure = hr_response.registers[1] / 100.0
                level = hr_response.registers[2]
            else:
                print(f"[{t:5.1f} с] Ошибка чтения Holding Registers (Func 0x03)")
                temp = pressure = level = None

            # 2. Чтение Discrete Inputs (Функция 0x02) - флаг аварии
            try:
                di_response = client.read_discrete_inputs(0, count=1)
            except TypeError:
                di_response = client.read_discrete_inputs(0, 1)

            if not di_response.isError():
                alarm = 1 if di_response.bits[0] else 0
            else:
                alarm = None

            # 3. Алгоритм телеуправления исполнительным органом (Функция 0x05 - Write Single Coil)
            # Правило автоматизации:
            # - Если Уровень < 40% -> Отправить команду «Пуск насоса» (Coil 0 = True)
            # - Если Уровень > 60% -> Отправить команду «Стоп насоса» (Coil 1 = True)
            cmd_info = "–"
            if level is not None:
                if level < 40:
                    client.write_coil(0, True)
                    cmd_info = "ПУСК НАСОСА (CO[0]=1)"
                    commands_log.append((t, "START", level))
                    print(f"[{t:5.1f} с] >>> ТЕЛЕУПРАВЛЕНИЕ: Уровень {level}% < 40% -> Отправлена команда ПУСК")
                elif level > 60:
                    client.write_coil(1, True)
                    cmd_info = "СТОП НАСОСА (CO[1]=1)"
                    commands_log.append((t, "STOP", level))
                    print(f"[{t:5.1f} с] >>> ТЕЛЕУПРАВЛЕНИЕ: Уровень {level}% > 60% -> Отправлена команда СТОП")

            # Фиксация в протоколе
            if temp is not None:
                times.append(t)
                temps.append(temp)
                pressures.append(pressure)
                levels.append(level)
                alarms.append(alarm)
                print(f"[{t:5.1f} с] T={temp:5.1f} °C, P={pressure:4.2f} бар, L={level:3d} %, Авария={alarm} | {cmd_info}")

            time.sleep(poll_interval)

    finally:
        client.close()
        print("-" * 70)
        print("Сессия опроса Modbus TCP завершена.")

    # Построение графических трендов
    if save_plot and times:
        fig, axes = plt.subplots(4, 1, figsize=(11, 9), sharex=True, dpi=150)

        # 1. Температура
        axes[0].plot(times, temps, color='#D0021B', lw=1.5, marker='o', markersize=3, label='Температура (°C)')
        axes[0].axhline(y=55.0, color='darkred', linestyle='--', lw=1.2, label='Порог аварии (55 °C)')
        axes[0].set_ylabel('T, °C')
        axes[0].set_title('Моделирование системы сбора данных и управления по протоколу Modbus TCP', fontweight='bold', pad=10)
        axes[0].grid(True, linestyle=':', alpha=0.6)
        axes[0].legend(loc='upper right', fontsize=8.5)

        # 2. Давление
        axes[1].plot(times, pressures, color='#4A90E2', lw=1.5, marker='s', markersize=3, label='Давление (бар)')
        axes[1].axhline(y=7.5, color='darkblue', linestyle='--', lw=1.2, label='Порог аварии (7.5 бар)')
        axes[1].set_ylabel('P, бар')
        axes[1].grid(True, linestyle=':', alpha=0.6)
        axes[1].legend(loc='upper right', fontsize=8.5)

        # 3. Уровень
        axes[2].plot(times, levels, color='#27AE60', lw=1.5, marker='^', markersize=3, label='Уровень в резервуаре (%)')
        axes[2].axhline(y=40.0, color='darkgreen', linestyle=':', lw=1.2, label='Уставка включения (40 %)')
        axes[2].axhline(y=60.0, color='orange', linestyle=':', lw=1.2, label='Уставка отключения (60 %)')
        axes[2].set_ylabel('Уровень, %')
        axes[2].grid(True, linestyle=':', alpha=0.6)
        axes[2].legend(loc='upper right', fontsize=8.5)

        # 4. Флаг аварии
        axes[3].step(times, alarms, color='#8E44AD', where='mid', lw=1.6, label='Дискретный вход: Авария (DI[0])')
        axes[3].set_ylabel('Авария')
        axes[3].set_xlabel('Время эксперимента, с')
        axes[3].set_ylim(-0.15, 1.25)
        axes[3].set_yticks([0, 1])
        axes[3].set_yticklabels(['Норма (0)', 'Авария (1)'])
        axes[3].grid(True, linestyle=':', alpha=0.6)
        axes[3].legend(loc='upper right', fontsize=8.5)

        plt.tight_layout()
        out_png = os.path.join(os.path.dirname(__file__), 'modbus_result.png')
        plt.savefig(out_png, dpi=150)
        plt.close()
        print(f"График сохранён в файл: {out_png}")

    return {
        'times': times,
        'temps': temps,
        'pressures': pressures,
        'levels': levels,
        'alarms': alarms,
        'commands': commands_log
    }


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Modbus TCP Client (Эмулятор SCADA)")
    parser.add_argument('--host', type=str, default='127.0.0.1', help="IP адрес сервера")
    parser.add_argument('--port', type=int, default=5020, help="TCP порт сервера")
    parser.add_argument('--duration', type=float, default=20.0, help="Длительность опроса в секундах")
    args = parser.parse_args()

    run_modbus_client(host=args.host, port=args.port, duration=args.duration)

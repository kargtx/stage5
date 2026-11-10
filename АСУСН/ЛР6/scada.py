# -*- coding: utf-8 -*-
"""
Модуль программного SCADA-эмулятора распределенной системы управления.
Дисциплина: АСУСН, Лабораторная работа №6.
Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.

Выполняет:
- Подключение по Modbus TCP к ПЛК №1 (порт 5020) и ПЛК №2 (порт 5021)
- Сбор данных технологического процесса в реальном времени
- Отображение мнемосхемы и трендов
- Возможность оперативного изменения технологической уставки
- Сохранение графиков в scada_result.png
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


class ScadaEmulator:
    """Диспетчерский терминал SCADA для распределенной системы управления."""

    def __init__(self, host='127.0.0.1', plc1_port=5020, plc2_port=5021):
        self.host = host
        self.plc1_port = int(plc1_port)
        self.plc2_port = int(plc2_port)

        self.client1 = ModbusTcpClient(self.host, port=self.plc1_port)
        self.client2 = ModbusTcpClient(self.host, port=self.plc2_port)

        self.history = {
            't': [],
            'temp': [],
            'setpoint': [],
            'heater': [],
            'level': [],
            'pressure': [],
            'pump': [],
            'valve': [],
            'alarm': []
        }

    def connect(self):
        c1 = self.client1.connect()
        c2 = self.client2.connect()
        return c1 and c2

    def disconnect(self):
        self.client1.close()
        self.client2.close()

    def set_setpoint(self, new_sp):
        """Отправка новой уставки в ПЛК №1 (HR[2])."""
        val = int(float(new_sp) * 10)
        try:
            self.client1.write_register(2, val)
            print(f"[SCADA] Отправлена новая уставка температуры: {new_sp:.1f} °C (HR[2]={val})")
        except Exception as e:
            print(f"[SCADA] Ошибка записи уставки: {e}")

    def read_telemetry(self):
        """Опрос обоих контроллеров."""
        data = {}

        # 1. Опрос ПЛК №1
        try:
            hr1 = self.client1.read_holding_registers(0, count=3)
            if not hr1.isError() and len(hr1.registers) >= 3:
                data['temp'] = hr1.registers[0] / 10.0
                data['heater'] = hr1.registers[1] / 10.0
                data['setpoint'] = hr1.registers[2] / 10.0
        except Exception:
            pass

        # 2. Опрос ПЛК №2
        try:
            hr2 = self.client2.read_holding_registers(0, count=3)
            if not hr2.isError() and len(hr2.registers) >= 3:
                data['level'] = hr2.registers[1]
                data['pressure'] = hr2.registers[2] / 100.0

            di2 = self.client2.read_discrete_inputs(0, count=3)
            if not di2.isError() and len(di2.bits) >= 3:
                data['pump'] = 1 if di2.bits[0] else 0
                data['valve'] = 1 if di2.bits[1] else 0
                data['alarm'] = 1 if di2.bits[2] else 0
        except Exception:
            pass

        return data

    def run_session(self, duration=20.0, poll_interval=0.5, save_plot=True):
        """Проведение измерительной сессии SCADA."""
        print("=" * 70)
        print("  SCADA ДИСПЕТЧЕРСКАЯ СТАНЦИЯ (СБОР ДАННЫХ И ТЕЛЕУПРАВЛЕНИЕ)")
        print(f"  Опрос ПЛК №1 (порт {self.plc1_port}) и ПЛК №2 (порт {self.plc2_port})")
        print(f"  Длительность: {duration} с (период {poll_interval} с)")
        print("=" * 70)

        self.connect()
        start = time.time()

        try:
            while time.time() - start < duration:
                t = round(time.time() - start, 2)
                d = self.read_telemetry()

                if d and 'temp' in d:
                    self.history['t'].append(t)
                    self.history['temp'].append(d.get('temp', 20.0))
                    self.history['setpoint'].append(d.get('setpoint', 60.0))
                    self.history['heater'].append(d.get('heater', 0.0))
                    self.history['level'].append(d.get('level', 50.0))
                    self.history['pressure'].append(d.get('pressure', 1.0))
                    self.history['pump'].append(d.get('pump', 0))
                    self.history['valve'].append(d.get('valve', 0))
                    self.history['alarm'].append(d.get('alarm', 0))

                    print(f"[{t:5.1f} с] T={d.get('temp', 0):4.1f}°C (SP={d.get('setpoint', 0):.0f}), "
                          f"ТЭН={d.get('heater', 0):3.0f}%, Уровень={d.get('level', 0):2.0f}%, "
                          f"P={d.get('pressure', 0):4.2f}бар | Насос={d.get('pump', 0)}, Клапан={d.get('valve', 0)}, Авария={d.get('alarm', 0)}")

                time.sleep(poll_interval)
        finally:
            self.disconnect()

        if save_plot and self.history['t']:
            self.plot_results()

        return self.history

    def plot_results(self, output_file='scada_result.png'):
        """Построение комплексных технологических трендов."""
        t = self.history['t']
        fig, axes = plt.subplots(3, 1, figsize=(11, 8.5), sharex=True, dpi=150)

        # 1. Температура и ПИД
        axes[0].plot(t, self.history['temp'], color='#C0392B', lw=1.6, label='Температура T (°C)')
        axes[0].plot(t, self.history['setpoint'], color='darkgreen', linestyle='--', lw=1.3, label='Уставка SP (°C)')
        ax0_t = axes[0].twinx()
        ax0_t.plot(t, self.history['heater'], color='#F39C12', lw=1.2, linestyle=':', label='Мощность ТЭНа (%)')
        ax0_t.set_ylabel('ТЭН, %', color='#B9770E')
        ax0_t.set_ylim(-5, 105)
        axes[0].set_ylabel('T, °C')
        axes[0].set_title('Мнемосхема трендов SCADA: Температурный контур и исполнительные органы', fontweight='bold', pad=8)
        axes[0].grid(True, linestyle=':', alpha=0.6)
        axes[0].legend(loc='upper left', fontsize=8.5)
        ax0_t.legend(loc='upper right', fontsize=8.5)

        # 2. Уровень и состояние насоса/клапана
        axes[1].plot(t, self.history['level'], color='#2980B9', lw=1.6, label='Уровень жидкости (%)')
        axes[1].axhline(80.0, color='darkred', linestyle=':', lw=1.0, label='Верхняя граница (80 %)')
        axes[1].axhline(30.0, color='darkblue', linestyle=':', lw=1.0, label='Нижняя граница (30 %)')
        axes[1].set_ylabel('Уровень, %')
        axes[1].grid(True, linestyle=':', alpha=0.6)
        axes[1].legend(loc='upper left', fontsize=8.5)

        # 3. Давление и статус сигнализации
        axes[2].plot(t, self.history['pressure'], color='#27AE60', lw=1.5, label='Давление (бар)')
        ax2_t = axes[2].twinx()
        ax2_t.step(t, self.history['alarm'], color='#8E44AD', where='mid', lw=1.6, label='Аварийная сигнализация')
        ax2_t.set_ylabel('Авария (0/1)', color='#8E44AD')
        ax2_t.set_ylim(-0.15, 1.25)
        ax2_t.set_yticks([0, 1])
        axes[2].set_ylabel('P, бар')
        axes[2].set_xlabel('Время эксперимента, с')
        axes[2].grid(True, linestyle=':', alpha=0.6)
        axes[2].legend(loc='upper left', fontsize=8.5)
        ax2_t.legend(loc='upper right', fontsize=8.5)

        plt.tight_layout()
        out_path = os.path.join(os.path.dirname(__file__), output_file)
        plt.savefig(out_path, dpi=150)
        plt.close()
        print(f"График сохранён в файл: {out_path}")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="SCADA-эмулятор")
    parser.add_argument('--duration', type=float, default=15.0, help="Длительность работы")
    parser.add_argument('--setpoint', type=float, default=None, help="Установить новую уставку")
    args = parser.parse_args()

    scada = ScadaEmulator()
    if args.setpoint is not None:
        scada.connect()
        scada.set_setpoint(args.setpoint)
        scada.disconnect()
    else:
        scada.run_session(duration=args.duration)

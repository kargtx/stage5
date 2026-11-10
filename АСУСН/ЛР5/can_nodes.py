# -*- coding: utf-8 -*-
"""
Модуль программной эмуляции промышленной шины CAN (Controller Area Network).
Дисциплина: АСУСН, Лабораторная работа №5.
Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.

Реализует:
- Узел 1 (Датчик температуры): периодическая передача телеметрии (ID = 0x100)
- Узел 2 (Исполнительный механизм): прием телеметрии и посылка аварийных команд (ID = 0x200 / 0x050)
- Построение графика CAN-трафика (can_result.png)
- Демонстрация побитового неразрушающего арбитража шины CAN
"""

import os
import sys
import time
import random
import threading
import argparse
import can
import matplotlib.pyplot as plt

# Настройка шрифтов matplotlib
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Segoe UI', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False


class CanNetworkSimulation:
    """Эмулятор распределенной системы на шине CAN с поддержкой Windows Virtual Bus."""

    def __init__(self, channel=0, duration=20.0, alarm_threshold=55.0):
        self.channel = channel
        self.duration = float(duration)
        self.alarm_threshold = float(alarm_threshold)

        # Журналы
        self.received_messages = []
        self.times = []
        self.temps_from_can = []
        self.alarms_from_can = []
        self.running = False

    def sensor_node_loop(self):
        """Узел 1: Датчик температуры (ID = 0x100, период 500 мс)."""
        # Отдельная шина для передающего узла на виртуальном интерфейсе
        bus_sensor = can.interface.Bus(interface='virtual', channel=self.channel)
        msg_id = 0x100
        start = time.time()

        try:
            while self.running and (time.time() - start < self.duration):
                temp = random.uniform(40.0, 60.0)
                # Упаковка температуры в 2 байта (Big-Endian, x10)
                data = list(int(temp * 10).to_bytes(2, byteorder='big'))
                
                # Дополняем кадр счетчиком пакетов и статусом
                data.append(int(time.time() * 10) % 256)
                data.append(0xAA)  # Контрольный маркер

                msg = can.Message(
                    arbitration_id=msg_id,
                    data=data,
                    is_extended_id=False
                )
                bus_sensor.send(msg)
                print(f"[Узел 1 (Датчик)]   >>> Отправлено: ID=0x{msg_id:03X}, T={temp:4.1f} °C, DLC={len(data)}")
                time.sleep(0.5)
        finally:
            bus_sensor.shutdown()

    def actuator_node_loop(self):
        """Узел 2: Исполнительный механизм (прием ID=0x100, отправка ID=0x200)."""
        bus_actuator = can.interface.Bus(interface='virtual', channel=self.channel)
        start = time.time()

        try:
            while self.running and (time.time() - start < self.duration):
                msg = bus_actuator.recv(timeout=0.6)
                if msg is None:
                    continue

                t = round(time.time() - start, 2)
                self.received_messages.append(msg)

                if msg.arbitration_id == 0x100:
                    temp = int.from_bytes(msg.data[:2], byteorder='big') / 10.0
                    self.times.append(t)
                    self.temps_from_can.append(temp)

                    is_alarm = 1 if temp > self.alarm_threshold else 0
                    self.alarms_from_can.append(is_alarm)

                    print(f"[Узел 2 (Механизм)] <<< Принято:    ID=0x{msg.arbitration_id:03X}, T={temp:4.1f} °C")

                    # Если температура превышает порог - генерация высокоприоритетного аварийного кадра
                    if temp > self.alarm_threshold:
                        alarm_msg = can.Message(
                            arbitration_id=0x200,
                            data=[0x01, int(temp), 0xFF],
                            is_extended_id=False
                        )
                        bus_actuator.send(alarm_msg)
                        print(f"[Узел 2 (Механизм)] !!! АВАРИЙНЫЙ КАДР: ID=0x200 (T={temp:.1f} °C > {self.alarm_threshold} °C)")

                elif msg.arbitration_id == 0x200:
                    print(f"[Узел 2 (Механизм)] <<< Подтверждение аварии: ID=0x200, код={list(msg.data)}")

        finally:
            bus_actuator.shutdown()

    def run(self, save_plot=True):
        """Запуск параллельной симуляции узлов."""
        print("=" * 70)
        print("  ЭМУЛЯЦИЯ ПРОМЫШЛЕННОЙ СЕТИ CAN (Controller Area Network)")
        print(f"  Виртуальный канал: {self.channel}, Длительность: {self.duration} с")
        print("=" * 70)

        self.running = True
        t_sensor = threading.Thread(target=self.sensor_node_loop, daemon=True)
        t_actuator = threading.Thread(target=self.actuator_node_loop, daemon=True)

        t_actuator.start()
        time.sleep(0.05)
        t_sensor.start()

        t_sensor.join()
        t_actuator.join()
        self.running = False

        print("-" * 70)
        print(f"Симуляция CAN завершена. Всего обработано кадров: {len(self.received_messages)}")

        if save_plot and self.times:
            fig, ax = plt.subplots(2, 1, figsize=(11, 7), sharex=True, dpi=150)

            # Верхний график: данные температуры
            ax[0].plot(self.times, self.temps_from_can, color='#C0392B', lw=1.6, marker='o', markersize=4, label='Температура из CAN-кадра (ID=0x100)')
            ax[0].axhline(y=self.alarm_threshold, color='darkred', linestyle='--', lw=1.3, label=f'Порог аварии ({self.alarm_threshold} °C)')
            ax[0].set_ylabel('Температура, °C')
            ax[0].set_title('Эмуляция обмена сообщениями в промышленной сети CAN', fontweight='bold', pad=10)
            ax[0].grid(True, linestyle=':', alpha=0.6)
            ax[0].legend(loc='upper right', fontsize=8.5)

            # Нижний график: флаг аварийного кадра ID=0x200
            ax[1].step(self.times, self.alarms_from_can, color='#8E44AD', where='mid', lw=1.8, label='Аварийное сообщение (ID=0x200)')
            ax[1].set_xlabel('Время эксперимента, с')
            ax[1].set_ylabel('Статус тревоги')
            ax[1].set_ylim(-0.15, 1.25)
            ax[1].set_yticks([0, 1])
            ax[1].set_yticklabels(['Норма (0)', 'Авария (1)'])
            ax[1].grid(True, linestyle=':', alpha=0.6)
            ax[1].legend(loc='upper right', fontsize=8.5)

            plt.tight_layout()
            out_png = os.path.join(os.path.dirname(__file__), 'can_result.png')
            plt.savefig(out_png, dpi=150)
            plt.close()
            print(f"График сохранён в файл: {out_png}")

        return {
            'times': self.times,
            'temps': self.temps_from_can,
            'alarms': self.alarms_from_can,
            'total_messages': len(self.received_messages)
        }


def demonstrate_can_arbitration():
    """Демонстрация побитового неразрушающего арбитража шины CAN."""
    print("\n--- ДЕМОНСТРАЦИЯ ПОБИТОВОГО АРБИТРАЖА В СЕТИ CAN ---")
    id1 = 0x100  # Двоичное: 001 0000 0000 (Датчик)
    id2 = 0x050  # Двоичное: 000 0101 0000 (Экстренная авария)
    id3 = 0x200  # Двоичное: 010 0000 0000 (Команда)

    print(f"Узел 1 (Телеметрия): ID = 0x{id1:03X} -> биты: {id1:011b}")
    print(f"Узел 2 (Авария):     ID = 0x{id2:03X} -> биты: {id2:011b}")
    print(f"Узел 3 (Команда):    ID = 0x{id3:03X} -> биты: {id3:011b}")
    print("Правило CAN: Доминантный бит '0' подавляет рецессивный бит '1'.")
    print("В бит-позиции 3 узел 0x050 выдает '0', тогда как 0x100 и 0x200 выдают '1'.")
    print("ИТОГ: Узел 0x050 безусловно выигрывает арбитраж без разрушения кадра!")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Эмуляция CAN узлов")
    parser.add_argument('--duration', type=float, default=15.0, help="Длительность симуляции в секундах")
    args = parser.parse_args()

    sim = CanNetworkSimulation(duration=args.duration)
    sim.run()
    demonstrate_can_arbitration()

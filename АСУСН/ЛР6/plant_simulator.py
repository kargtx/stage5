# -*- coding: utf-8 -*-
"""
Модуль имитатора технологического объекта управления (Plant Simulator).
Дисциплина: АСУСН, Лабораторная работа №6.
Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.

Моделирует физические процессы резервуара с жидкостью:
- Тепловой баланс: нагрев ТЭНом и охлаждение в окружающую среду
- Гидродинамический баланс: приток от насоса и слив через регулирующий клапан
- Барометрическое давление в емкости
- Поддержка внесения отказов датчиков и исполнительных органов
- Сетевой TCP-интерфейс (порт 9999) для сопряжения с ПЛК №2
"""

import socket
import json
import time
import argparse


class PlantSimulator:
    """Математическая модель технологического аппарата с жидкостным теплообменом."""

    def __init__(self, init_temp=20.0, init_level=50.0, env_temp=20.0):
        self.temperature = float(init_temp)
        self.level = float(init_level)
        self.pressure = 1.0
        self.env_temp = float(env_temp)

        # Конструктивные коэффициенты
        self.K_heat = 0.5    # Коэффициент нагрева (°C / (% · с))
        self.K_cool = 0.05   # Коэффициент теплоотдачи в среду (1/с)
        self.K_in = 2.0      # Скорость наполнения от насоса (%/с)
        self.K_out = 1.0     # Скорость опорожнения через клапан (%/с)

        # Входные управляющие сигналы (от ПЛК №2)
        self.heater_power = 0.0  # Мощность нагревателя (0..100 %)
        self.pump_on = False     # Состояние насоса подачи
        self.valve_open = False  # Состояние клапана слива

        # Режим внесения отказов
        self.fault_mode = None
        self.fault_param = None

    def set_fault(self, fault_mode, param=None):
        """Внесение искусственного отказа (sensor_open, pump_jammed, heater_short)."""
        self.fault_mode = fault_mode
        self.fault_param = param

    def clear_fault(self):
        """Сброс активного отказа."""
        self.fault_mode = None
        self.fault_param = None

    def step(self, dt=0.1):
        """
        Один такт численного интегрирования физической модели.
        
        :param dt: шаг интегрирования по времени (с)
        :return: словарь текущего состояния {'temperature', 'level', 'pressure'}
        """
        # 1. Тепловой баланс (нагрев ТЭНом с учетом естественного охлаждения)
        effective_power = self.heater_power
        if self.fault_mode == 'heater_short':
            effective_power = 100.0  # Заклинивание нагревателя на максимуме

        dT = (effective_power * self.K_heat - self.K_cool * (self.temperature - self.env_temp)) * dt
        self.temperature += dT

        # 2. Гидродинамика (уровень жидкости)
        eff_pump = self.pump_on
        if self.fault_mode == 'pump_jammed':
            eff_pump = False

        q_in = self.K_in if eff_pump else 0.0
        q_out = self.K_out if self.valve_open else 0.0
        self.level += (q_in - q_out) * dt

        # Физические ограничения
        self.level = max(0.0, min(100.0, self.level))
        self.temperature = max(0.0, min(200.0, self.temperature))

        # 3. Давление в емкости (гидростатическое + температурное расширение пара)
        self.pressure = 0.5 + 0.05 * self.level + 0.02 * self.temperature

        # Имитация отказа измерительного датчика температуры (обрыв цепи)
        reported_temp = self.temperature
        if self.fault_mode == 'sensor_open':
            reported_temp = -50.0  # Падение сигнала ниже диапазона

        return {
            'temperature': round(reported_temp, 2),
            'real_temperature': round(self.temperature, 2),
            'level': round(self.level, 2),
            'pressure': round(self.pressure, 3)
        }


def run_plant_server(host='127.0.0.1', port=9999):
    """Запуск TCP-сервера имитатора объекта для подключения ПЛК №2."""
    plant = PlantSimulator()
    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server.bind((host, port))
    server.listen(1)

    print("=" * 70)
    print("  ИМИТАТОР ТЕХНОЛОГИЧЕСКОГО ОБЪЕКТА (PLANT SIMULATOR)")
    print(f"  Сервер запущен и ожидает подключение ПЛК №2 на {host}:{port}")
    print("=" * 70)

    try:
        conn, addr = server.accept()
        print(f"[TCP Server] Успешное подключение от ПЛК №2: {addr}")
        while True:
            data = conn.recv(1024)
            if not data:
                break
            commands = json.loads(data.decode('utf-8'))
            plant.heater_power = float(commands.get('heater_power', 0.0))
            plant.pump_on = bool(commands.get('pump_on', False))
            plant.valve_open = bool(commands.get('valve_open', False))

            state = plant.step(dt=0.1)
            conn.sendall(json.dumps(state).encode('utf-8'))
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("\nОстановка имитатора по команде оператора.")
    finally:
        server.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Имитатор технологического объекта")
    parser.add_argument('--port', type=int, default=9999, help="TCP порт сервера")
    args = parser.parse_args()

    run_plant_server(port=args.port)

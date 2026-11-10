# -*- coding: utf-8 -*-
"""
Модуль программного эмулятора ПЛК (Modbus TCP Server).
Дисциплина: АСУСН, Лабораторная работа №5.
Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.

Реализует карту памяти Modbus:
- Holding Registers (HR):
    HR[0] - Температура x10 (°C)
    HR[1] - Давление x100 (бар)
    HR[2] - Уровень (%)
- Discrete Inputs (DI):
    DI[0] - Флаг аварийного состояния (Alarm)
- Coils (CO):
    CO[0] - Команда телеуправления «Пуск насоса»
    CO[1] - Команда телеуправления «Стоп насоса»
"""

import time
import random
import threading
import argparse

# Совместимость с различными версиями pymodbus (2.x / 3.x)
try:
    from pymodbus.datastore import ModbusSlaveContext
except ImportError:
    from pymodbus.datastore import ModbusDeviceContext as ModbusSlaveContext

from pymodbus.datastore import ModbusSequentialDataBlock, ModbusServerContext
from pymodbus.server import StartTcpServer


class ModbusPlcServer:
    """Эмулятор программируемого логического контроллера (ПЛК) с протоколом Modbus TCP."""

    def __init__(self, host="127.0.0.1", port=5020, variant=1):
        self.host = host
        self.port = int(port)
        self.variant = int(variant)
        self.running = False

        # Инициализация блоков данных Modbus
        # Используем адрес 1 для совместимости с pymodbus 3.x
        self.di_block = ModbusSequentialDataBlock(1, [0] * 100)
        self.co_block = ModbusSequentialDataBlock(1, [0] * 100)
        self.hr_block = ModbusSequentialDataBlock(1, [0] * 100)
        self.ir_block = ModbusSequentialDataBlock(1, [0] * 100)

        # Контекст устройства
        try:
            self.store = ModbusSlaveContext(
                di=self.di_block,
                co=self.co_block,
                hr=self.hr_block,
                ir=self.ir_block
            )
            self.context = ModbusServerContext(devices=self.store)
        except Exception:
            self.store = ModbusSlaveContext(
                di=self.di_block,
                co=self.co_block,
                hr=self.hr_block,
                ir=self.ir_block
            )
            self.context = ModbusServerContext(slaves=self.store, single=True)

        self.server_thread = None
        self.update_thread = None

    def _get_target_list(self, block, block_idx=2):
        """Вспомогательный метод для прямого доступа к списку значений блока данных (для версий pymodbus 2.x/3.x)."""
        if hasattr(self.store, 'simdevice') and hasattr(self.store.simdevice, 'simdata'):
            simdata_tuple = self.store.simdevice.simdata
            if len(simdata_tuple) > block_idx and len(simdata_tuple[block_idx]) > 0:
                return simdata_tuple[block_idx][0].values
        if hasattr(block, 'values'):
            return block.values
        return None

    def update_simulation_loop(self):
        """Фоновый поток: генерация технологического процесса и обновление регистров."""
        while self.running:
            di_vals = self._get_target_list(self.di_block, block_idx=0)
            co_vals = self._get_target_list(self.co_block, block_idx=1)
            hr_vals = self._get_target_list(self.hr_block, block_idx=2)

            # Генерация данных по Варианту 1 (базовый)
            # HR[0] - Температура (40..60 °C), масштабирована x10
            temp = random.uniform(40.0, 60.0)
            # HR[1] - Давление (5..8 бар), масштабировано x100
            pressure = random.uniform(5.0, 8.0)
            # HR[2] - Уровень (30..70 %)
            level = random.uniform(30.0, 70.0)

            alarm = 1 if (temp > 55.0 or pressure > 7.5) else 0

            if hr_vals is not None and len(hr_vals) >= 3:
                hr_vals[0] = int(temp * 10)
                hr_vals[1] = int(pressure * 100)
                hr_vals[2] = int(level)

            if di_vals is not None and len(di_vals) >= 1:
                di_vals[0] = bool(alarm)

            # Обработка команд телеуправления из катушек CO
            if co_vals is not None and len(co_vals) >= 2:
                if co_vals[0]:  # Пуск
                    co_vals[0] = False
                if co_vals[1]:  # Стоп
                    co_vals[1] = False

            time.sleep(0.5)

    def start(self, blocking=True):
        """Запуск сервера Modbus TCP."""
        self.running = True
        self.update_thread = threading.Thread(target=self.update_simulation_loop, daemon=True)
        self.update_thread.start()

        print(f"Modbus TCP сервер запущен на {self.host}:{self.port}")
        print("  HR[0] = Температура x10 (°C)")
        print("  HR[1] = Давление x100 (бар)")
        print("  HR[2] = Уровень (%)")
        print("  DI[0] = Сигнал аварии (Alarm)")
        print("  CO[0] = Пуск насоса, CO[1] = Стоп насоса")

        if blocking:
            try:
                StartTcpServer(context=self.context, address=(self.host, self.port))
            except KeyboardInterrupt:
                self.stop()
        else:
            self.server_thread = threading.Thread(
                target=StartTcpServer,
                kwargs={'context': self.context, 'address': (self.host, self.port)},
                daemon=True
            )
            self.server_thread.start()

    def stop(self):
        """Остановка эмулятора."""
        self.running = False
        print("Modbus TCP сервер остановлен.")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Modbus TCP Server (Эмулятор ПЛК)")
    parser.add_argument('--port', type=int, default=5020, help="TCP порт (по умолчанию 5020)")
    parser.add_argument('--variant', type=int, default=1, help="Номер варианта (1-6)")
    args = parser.parse_args()

    server = ModbusPlcServer(port=args.port, variant=args.variant)
    server.start(blocking=True)

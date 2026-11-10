# -*- coding: utf-8 -*-
"""
Модуль виртуального контроллера ПЛК №1 (ПИД-регулятор температуры).
Дисциплина: АСУСН, Лабораторная работа №6.
Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.

Функции:
- Modbus TCP Server на порту 5020
- Карта регистров:
    HR[0] - Температура x10 (°C)
    HR[1] - Мощность ТЭНа x10 (%)
    HR[2] - Заданная уставка x10 (°C)
    HR[3] - Уровень жидкости (%)
    HR[4] - Давление x100 (бар)
- Дискретный ПИД-регулятор температуры с антисатурацией (Anti-Windup Clamping)
"""

import time
import threading
import argparse

try:
    from pymodbus.datastore import ModbusSlaveContext
except ImportError:
    from pymodbus.datastore import ModbusDeviceContext as ModbusSlaveContext

from pymodbus.datastore import ModbusSequentialDataBlock, ModbusServerContext
from pymodbus.server import StartTcpServer


class VirtualPlc1:
    """Виртуальный программируемый логический контроллер ПЛК №1."""

    def __init__(self, port=5020, setpoint=60.0):
        self.port = int(port)
        self.running = False

        # Параметры ПИД-регулятора
        self.state = {
            'temperature': 20.0,
            'setpoint': float(setpoint),
            'heater_power': 0.0,
            'integral': 0.0,
            'prev_error': 0.0,
            'Kp': 2.0,
            'Ki': 0.5,
            'Kd': 0.1
        }

        # Блоки данных Modbus
        self.di_block = ModbusSequentialDataBlock(1, [0] * 100)
        self.co_block = ModbusSequentialDataBlock(1, [0] * 100)
        self.hr_block = ModbusSequentialDataBlock(1, [0] * 100)
        self.ir_block = ModbusSequentialDataBlock(1, [0] * 100)

        try:
            self.store = ModbusSlaveContext(di=self.di_block, co=self.co_block, hr=self.hr_block, ir=self.ir_block)
            self.context = ModbusServerContext(devices=self.store)
        except Exception:
            self.store = ModbusSlaveContext(di=self.di_block, co=self.co_block, hr=self.hr_block, ir=self.ir_block)
            self.context = ModbusServerContext(slaves=self.store, single=True)

        self._init_registers()

    def _get_target_list(self, block, block_idx=2):
        if hasattr(self.store, 'simdevice') and hasattr(self.store.simdevice, 'simdata'):
            simdata_tuple = self.store.simdevice.simdata
            if len(simdata_tuple) > block_idx and len(simdata_tuple[block_idx]) > 0:
                return simdata_tuple[block_idx][0].values
        if hasattr(block, 'values'):
            return block.values
        return None

    def _init_registers(self):
        hr = self._get_target_list(self.hr_block, block_idx=2)
        if hr is not None and len(hr) >= 5:
            hr[0] = int(self.state['temperature'] * 10)
            hr[1] = int(self.state['heater_power'] * 10)
            hr[2] = int(self.state['setpoint'] * 10)

    def pid_step(self, dt=0.1):
        """Один такт дискретного ПИД-регулирования с защитой от интегрального насыщения."""
        error = self.state['setpoint'] - self.state['temperature']

        # Пропорциональная и дифференциальная составляющие
        p_term = self.state['Kp'] * error
        derivative = (error - self.state['prev_error']) / dt if dt > 0 else 0.0
        d_term = self.state['Kd'] * derivative

        # Интегральная составляющая с Anti-Windup Clamping
        tentative_integral = self.state['integral'] + error * dt
        i_term = self.state['Ki'] * tentative_integral

        raw_output = p_term + i_term + d_term

        # Насыщение выхода 0..100%
        if raw_output > 100.0:
            output = 100.0
            # Если ошибка одного знака с насыщением - заморозка интегрирования
            if error <= 0:
                self.state['integral'] = tentative_integral
        elif raw_output < 0.0:
            output = 0.0
            if error >= 0:
                self.state['integral'] = tentative_integral
        else:
            output = raw_output
            self.state['integral'] = tentative_integral

        self.state['prev_error'] = error
        self.state['heater_power'] = output
        return output

    def control_loop(self):
        """Основной цикл управления ПЛК №1 (период 100 мс)."""
        while self.running:
            hr = self._get_target_list(self.hr_block, block_idx=2)

            # Проверка изменения уставки из SCADA (регистр HR[2])
            if hr is not None and len(hr) >= 3:
                scada_sp = hr[2] / 10.0
                if abs(scada_sp - self.state['setpoint']) > 0.01 and scada_sp > 0:
                    self.state['setpoint'] = scada_sp

            heater = self.pid_step(dt=0.1)

            # Обновление регистров
            if hr is not None and len(hr) >= 3:
                hr[0] = int(self.state['temperature'] * 10)
                hr[1] = int(heater * 10)
                hr[2] = int(self.state['setpoint'] * 10)

            time.sleep(0.1)

    def start(self, blocking=True):
        self.running = True
        t_ctrl = threading.Thread(target=self.control_loop, daemon=True)
        t_ctrl.start()

        print(f"ПЛК №1 (ПИД-регулятор) запущен на порту {self.port}")
        print("  HR[0] = T x10, HR[1] = Heater x10, HR[2] = Setpoint x10")

        if blocking:
            try:
                StartTcpServer(context=self.context, address=("127.0.0.1", self.port))
            except KeyboardInterrupt:
                self.stop()
        else:
            self.srv_th = threading.Thread(
                target=StartTcpServer,
                kwargs={'context': self.context, 'address': ("127.0.0.1", self.port)},
                daemon=True
            )
            self.srv_th.start()

    def stop(self):
        self.running = False
        print("ПЛК №1 остановлен.")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Виртуальный ПЛК №1")
    parser.add_argument('--port', type=int, default=5020, help="Modbus TCP порт")
    parser.add_argument('--setpoint', type=float, default=60.0, help="Начальная уставка температуры")
    args = parser.parse_args()

    plc = VirtualPlc1(port=args.port, setpoint=args.setpoint)
    plc.start(blocking=True)

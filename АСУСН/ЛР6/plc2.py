# -*- coding: utf-8 -*-
"""
Модуль виртуального контроллера ПЛК №2 (Управление насосом, клапаном и связь с объектом).
Дисциплина: АСУСН, Лабораторная работа №6.
Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.

Функции:
- Взаимодействие с физическим объектом по TCP-сокету (порт 9999)
- Межконтроллерный опрос ПЛК №1 по Modbus TCP (порт 5020)
- Реализация противоаварийной блокировки и двухпозиционного регулирования
- Modbus TCP Server для SCADA-системы на порту 5021:
    HR[0] - Температура x10 (°C)
    HR[1] - Уровень (%)
    HR[2] - Давление x100 (бар)
    DI[0] - Состояние насоса (Pump On)
    DI[1] - Состояние клапана (Valve Open)
    DI[2] - Сигнал аварии (Alarm)
"""

import socket
import json
import time
import threading
import argparse

try:
    from pymodbus.datastore import ModbusSlaveContext
except ImportError:
    from pymodbus.datastore import ModbusDeviceContext as ModbusSlaveContext

from pymodbus.datastore import ModbusSequentialDataBlock, ModbusServerContext
from pymodbus.server import StartTcpServer
from pymodbus.client import ModbusTcpClient


class VirtualPlc2:
    """Виртуальный программируемый логический контроллер ПЛК №2."""

    def __init__(self, port=5021, plc1_port=5020, plant_port=9999, variant=1):
        self.port = int(port)
        self.plc1_port = int(plc1_port)
        self.plant_port = int(plant_port)
        self.variant = int(variant)
        self.running = False

        self.state = {
            'temperature': 20.0,
            'level': 50.0,
            'pressure': 1.0,
            'pump_on': False,
            'valve_open': False,
            'heater_power': 0.0,
            'alarm': False,
            'plc1_connected': False
        }

        # Блоки Modbus для ПЛК №2
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

        self.plant_socket = None

    def _get_target_list(self, block, block_idx=2):
        if hasattr(self.store, 'simdevice') and hasattr(self.store.simdevice, 'simdata'):
            simdata_tuple = self.store.simdevice.simdata
            if len(simdata_tuple) > block_idx and len(simdata_tuple[block_idx]) > 0:
                return simdata_tuple[block_idx][0].values
        if hasattr(block, 'values'):
            return block.values
        return None

    def connect_to_plant(self):
        """Подключение к TCP-серверу имитатора объекта."""
        try:
            self.plant_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.plant_socket.connect(('127.0.0.1', self.plant_port))
            print(f"[ПЛК №2] Успешное подключение к имитатору объекта на 127.0.0.1:{self.plant_port}")
            return True
        except Exception as e:
            print(f"[ПЛК №2] Предупреждение: имитатор объекта пока недоступен ({e})")
            return False

    def read_from_plc1(self):
        """Чтение регуляторных сигналов от ПЛК №1 по Modbus TCP."""
        client = ModbusTcpClient('127.0.0.1', port=self.plc1_port)
        if not client.connect():
            self.state['plc1_connected'] = False
            return False

        try:
            res = client.read_holding_registers(0, count=3)
            if not res.isError() and len(res.registers) >= 2:
                self.state['heater_power'] = res.registers[1] / 10.0
                self.state['plc1_connected'] = True
                return True
        except Exception:
            self.state['plc1_connected'] = False
        finally:
            client.close()
        return False

    def control_loop(self):
        """Цикл управления ПЛК №2 (период 100 мс)."""
        while self.running:
            # 1. Считывание регуляторной мощности от ПЛК №1
            self.read_from_plc1()

            # 2. Логика управления агрегатами по Варианту 1
            t_cur = self.state['temperature']
            l_cur = self.state['level']

            # Контур охлаждения и откачки
            if t_cur > 70.0:
                self.state['pump_on'] = True
            elif t_cur < 50.0:
                self.state['pump_on'] = False

            # Контур защиты по уровню
            if l_cur > 80.0:
                self.state['valve_open'] = True
            elif l_cur < 30.0:
                self.state['valve_open'] = False

            # Сигнал аварии
            self.state['alarm'] = bool(t_cur > 85.0 or self.state['pressure'] > 7.5 or not self.state['plc1_connected'])

            # 3. Обмен данными с физическим объектом через сокет
            if self.plant_socket:
                cmd = {
                    'heater_power': self.state['heater_power'] if self.state['plc1_connected'] else 0.0,
                    'pump_on': self.state['pump_on'],
                    'valve_open': self.state['valve_open']
                }
                try:
                    self.plant_socket.sendall(json.dumps(cmd).encode('utf-8'))
                    raw_data = self.plant_socket.recv(1024)
                    if raw_data:
                        plant_state = json.loads(raw_data.decode('utf-8'))
                        self.state['temperature'] = plant_state['temperature']
                        self.state['level'] = plant_state['level']
                        self.state['pressure'] = plant_state['pressure']
                except Exception:
                    self.plant_socket = None

            # 4. Обновление регистров Modbus TCP для SCADA
            hr = self._get_target_list(self.hr_block, block_idx=2)
            di = self._get_target_list(self.di_block, block_idx=0)

            if hr is not None and len(hr) >= 3:
                hr[0] = int(self.state['temperature'] * 10)
                hr[1] = int(self.state['level'])
                hr[2] = int(self.state['pressure'] * 100)

            if di is not None and len(di) >= 3:
                di[0] = bool(self.state['pump_on'])
                di[1] = bool(self.state['valve_open'])
                di[2] = bool(self.state['alarm'])

            time.sleep(0.1)

    def start(self, blocking=True):
        self.running = True
        self.connect_to_plant()

        t_ctrl = threading.Thread(target=self.control_loop, daemon=True)
        t_ctrl.start()

        print(f"ПЛК №2 (Логика насоса и клапана) запущен на порту {self.port}")
        print("  HR[0]=T, HR[1]=Level, HR[2]=Pressure; DI[0]=Pump, DI[1]=Valve, DI[2]=Alarm")

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
        if self.plant_socket:
            try:
                self.plant_socket.close()
            except Exception:
                pass
        print("ПЛК №2 остановлен.")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Виртуальный ПЛК №2")
    parser.add_argument('--port', type=int, default=5021, help="Modbus TCP порт для SCADA")
    parser.add_argument('--plc1_port', type=int, default=5020, help="Modbus TCP порт ПЛК №1")
    parser.add_argument('--plant_port', type=int, default=9999, help="TCP порт имитатора объекта")
    args = parser.parse_args()

    plc2 = VirtualPlc2(port=args.port, plc1_port=args.plc1_port, plant_port=args.plant_port)
    plc2.start(blocking=True)

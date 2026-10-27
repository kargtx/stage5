# -*- coding: utf-8 -*-
"""
Имитатор виртуального ПЛК Modbus TCP Server (Slave) для Лабораторной работы №3 по АСУСН
Кафедра АСУ, УУНиТ, 2026 г.

Скрипт эмулирует работу контроллера полевого уровня (WinPLC / CODESYS):
- Протокол: Modbus TCP (по умолчанию IP 127.0.0.1, порт 502)
- Предоставляет технологические параметры для всех вариантов лабораторной работы
- Поддерживает команды телеуправления от SCADA (Пуск/Стоп насоса)
- Позволяет мгновенно активировать аварийный режим для проверки тревог и архивов в SCADA
"""

import sys
import time
import math
import random
import socket
import struct
import threading

# Настройка кодировки консоли Windows
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Параметры вариантов заданий
VARIANTS_CONFIG = {
    1: {
        "name": "Вариант 1 (Температура, Давление, Уровень)",
        "p1_name": "Температура", "p1_unit": "°C", "p1_range": (0, 100), "p1_norm": (45.0, 55.0), "p1_alarm": 88.5, "p1_thresh": 80.0, "p1_scale": 10.0,
        "p2_name": "Давление",    "p2_unit": "бар", "p2_range": (0, 10),  "p2_norm": (4.5, 5.5),   "p2_alarm": 9.2,  "p2_thresh": 8.0,  "p2_scale": 100.0,
        "p3_name": "Уровень",     "p3_unit": "%",   "p3_range": (0, 100), "p3_norm": (60.0, 70.0), "p3_alarm": 92.0, "p3_thresh": 90.0, "p3_scale": 10.0,
    },
    2: {
        "name": "Вариант 2 (Расход, Уровень, Вибрация)",
        "p1_name": "Расход",      "p1_unit": "м³/ч", "p1_range": (0, 500), "p1_norm": (250.0, 320.0), "p1_alarm": 445.0, "p1_thresh": 400.0, "p1_scale": 1.0,
        "p2_name": "Уровень",     "p2_unit": "%",    "p2_range": (0, 100), "p2_norm": (55.0, 65.0),   "p2_alarm": 85.0,  "p2_thresh": 80.0,  "p2_scale": 10.0,
        "p3_name": "Вибрация",    "p3_unit": "мм/с", "p3_range": (0, 20),  "p3_norm": (4.0, 7.0),     "p3_alarm": 17.5,  "p3_thresh": 15.0,  "p3_scale": 100.0,
    },
    3: {
        "name": "Вариант 3 (Напряжение, Ток, Мощность)",
        "p1_name": "Напряжение",  "p1_unit": "В",   "p1_range": (200, 250), "p1_norm": (220.0, 230.0), "p1_alarm": 248.0, "p1_thresh": 245.0, "p1_scale": 10.0,
        "p2_name": "Ток",         "p2_unit": "А",   "p2_range": (0, 100),   "p2_norm": (40.0, 60.0),   "p2_alarm": 95.0,  "p2_thresh": 90.0,  "p2_scale": 10.0,
        "p3_name": "Мощность",    "p3_unit": "кВт", "p3_range": (0, 50),    "p3_norm": (20.0, 30.0),   "p3_alarm": 46.0,  "p3_thresh": 45.0,  "p3_scale": 10.0,
    },
    4: {
        "name": "Вариант 4 (pH, Мутность, Температура)",
        "p1_name": "pH",          "p1_unit": "pH",  "p1_range": (0, 14),   "p1_norm": (7.0, 7.6),    "p1_alarm": 9.5,   "p1_thresh": 9.0,   "p1_scale": 100.0,
        "p2_name": "Мутность",    "p2_unit": "NTU", "p2_range": (0, 100),  "p2_norm": (15.0, 30.0),  "p2_alarm": 85.0,  "p2_thresh": 80.0,  "p2_scale": 10.0,
        "p3_name": "Температура", "p3_unit": "°C",  "p3_range": (0, 60),   "p3_norm": (25.0, 35.0),  "p3_alarm": 55.0,  "p3_thresh": 50.0,  "p3_scale": 10.0,
    },
    5: {
        "name": "Вариант 5 (Давление, Расход, Уровень)",
        "p1_name": "Давление",    "p1_unit": "бар",  "p1_range": (0, 16),  "p1_norm": (7.0, 9.0),     "p1_alarm": 15.2,  "p1_thresh": 14.0,  "p1_scale": 100.0,
        "p2_name": "Расход",      "p2_unit": "м³/ч", "p2_range": (0, 200), "p2_norm": (90.0, 120.0), "p2_alarm": 42.0,  "p2_thresh": 50.0,  "p2_scale": 1.0,
        "p3_name": "Уровень",     "p3_unit": "%",    "p3_range": (0, 100), "p3_norm": (60.0, 75.0),   "p3_alarm": 93.0,  "p3_thresh": 90.0,  "p3_scale": 10.0,
    }
}

class ModbusPLCSimulator:
    def __init__(self, host='0.0.0.0', port=502, variant=1):
        self.host = host
        self.port = port
        self.variant = variant
        self.running = False
        self.lock = threading.Lock()
        
        # Режим работы модели: 'NORMAL' или 'ALARM'
        self.mode = 'NORMAL'
        self.pump_state = 1      # 1 = Работает, 0 = Остановлен
        self.pump_control = 1    # Значение регистра управления
        
        # Внутренние физические величины
        self.val1 = 50.0
        self.val2 = 5.0
        self.val3 = 65.0
        
        # Карта регистров хранения Holding Registers (адреса 0..19)
        self.holding_registers = [0] * 30
        
        # Карта дискретных флагов Coils (адреса 0..9)
        self.coils = [False] * 10
        self.coils[0] = True
        
        self.update_registers()

    def float_to_registers(self, val):
        """Преобразует 32-битный float в два 16-битных регистра (Modbus standard Big-Endian)"""
        raw = struct.pack('>f', float(val))
        reg1, reg2 = struct.unpack('>HH', raw)
        return reg1, reg2

    def update_registers(self):
        """Обновляет значения регистров Modbus на основе текущего состояния модели"""
        cfg = VARIANTS_CONFIG[self.variant]
        with self.lock:
            # Целочисленные масштабированные значения (Holding Registers 0..4)
            # Регистр 0: Параметр 1
            r0 = int(round(self.val1 * cfg['p1_scale']))
            # Регистр 1: Параметр 2
            r1 = int(round(self.val2 * cfg['p2_scale']))
            # Регистр 2: Параметр 3
            r2 = int(round(self.val3 * cfg['p3_scale']))
            # Регистр 3: Состояние агрегата (1 - работа, 0 - останов)
            r3 = int(self.pump_state)
            # Регистр 4: Канал управления от SCADA
            r4 = int(self.pump_control)
            
            self.holding_registers[0] = max(0, min(65535, r0))
            self.holding_registers[1] = max(0, min(65535, r1))
            self.holding_registers[2] = max(0, min(65535, r2))
            self.holding_registers[3] = r3
            self.holding_registers[4] = r4
            
            # Регистры 10..15: вещественные значения Float32 (для гибкости SCADA)
            f1_hi, f1_lo = self.float_to_registers(self.val1)
            f2_hi, f2_lo = self.float_to_registers(self.val2)
            f3_hi, f3_lo = self.float_to_registers(self.val3)
            
            self.holding_registers[10] = f1_hi
            self.holding_registers[11] = f1_lo
            self.holding_registers[12] = f2_hi
            self.holding_registers[13] = f2_lo
            self.holding_registers[14] = f3_hi
            self.holding_registers[15] = f3_lo

    def simulation_step(self):
        """Шаг физической модели (каждые 500 мс)"""
        cfg = VARIANTS_CONFIG[self.variant]
        noise1 = random.uniform(-0.3, 0.3)
        noise2 = random.uniform(-0.05, 0.05)
        noise3 = random.uniform(-0.2, 0.2)
        
        # Если насос отключен, давление/расход падают
        if self.pump_state == 0:
            target_p2 = cfg['p2_norm'][0] * 0.1
        else:
            target_p2 = (cfg['p2_norm'][0] + cfg['p2_norm'][1]) / 2.0
            
        if self.mode == 'NORMAL':
            norm1 = (cfg['p1_norm'][0] + cfg['p1_norm'][1]) / 2.0
            norm3 = (cfg['p3_norm'][0] + cfg['p3_norm'][1]) / 2.0
            
            self.val1 += (norm1 - self.val1) * 0.1 + noise1
            self.val2 += (target_p2 - self.val2) * 0.15 + noise2
            self.val3 += (norm3 - self.val3) * 0.05 + noise3
        else:
            # Аварийный режим - значения уходят в зону срабатывания тревоги
            self.val1 += (cfg['p1_alarm'] - self.val1) * 0.15 + noise1
            self.val2 += (cfg['p2_alarm'] - self.val2) * 0.15 + noise2
            self.val3 += (cfg['p3_alarm'] - self.val3) * 0.1 + noise3
            
        self.update_registers()

    def process_modbus_pdu(self, unit_id, fc, pdu):
        """Обработка стандартных PDU протокола Modbus"""
        # FC 03 (0x03) - Read Holding Registers
        # FC 04 (0x04) - Read Input Registers
        if fc in (3, 4):
            if len(pdu) < 4:
                return bytes([fc | 0x80, 0x03])  # Exception: Illegal Data Value
            start_addr, qty = struct.unpack('>HH', pdu[:4])
            if start_addr + qty > len(self.holding_registers):
                return bytes([fc | 0x80, 0x02])  # Exception: Illegal Data Address
            
            with self.lock:
                regs = self.holding_registers[start_addr:start_addr + qty]
            byte_count = qty * 2
            resp = struct.pack('>BB', fc, byte_count)
            for r in regs:
                resp += struct.pack('>H', r)
            return resp
            
        # FC 01 (0x01) - Read Coils
        elif fc == 1:
            if len(pdu) < 4:
                return bytes([fc | 0x80, 0x03])
            start_addr, qty = struct.unpack('>HH', pdu[:4])
            byte_count = (qty + 7) // 8
            resp = struct.pack('>BB', fc, byte_count)
            with self.lock:
                bit_val = 1 if self.coils[0] else 0
            resp += bytes([bit_val])
            return resp
            
        # FC 05 (0x05) - Write Single Coil
        elif fc == 5:
            if len(pdu) < 4:
                return bytes([fc | 0x80, 0x03])
            addr, val = struct.unpack('>HH', pdu[:4])
            new_state = (val == 0xFF00)
            with self.lock:
                if addr < len(self.coils):
                    self.coils[addr] = new_state
                self.pump_control = 1 if new_state else 0
                self.pump_state = self.pump_control
            print(f"  [SCADA WRITE] Команда Coil {addr} = {new_state} (Насос: {'ПУСК' if new_state else 'СТОП'})")
            return pdu  # Echo request
            
        # FC 06 (0x06) - Write Single Register
        elif fc == 6:
            if len(pdu) < 4:
                return bytes([fc | 0x80, 0x03])
            addr, val = struct.unpack('>HH', pdu[:4])
            with self.lock:
                if addr < len(self.holding_registers):
                    self.holding_registers[addr] = val
                if addr == 4:
                    self.pump_control = val
                    self.pump_state = 1 if val != 0 else 0
            print(f"  [SCADA WRITE] Запись Register {addr} = {val} (Насос: {'ПУСК' if self.pump_state else 'СТОП'})")
            return pdu
            
        # FC 16 (0x10) - Write Multiple Registers
        elif fc == 16:
            if len(pdu) < 5:
                return bytes([fc | 0x80, 0x03])
            start_addr, qty, byte_count = struct.unpack('>HHB', pdu[:5])
            values = struct.unpack(f'>{qty}H', pdu[5:5 + qty * 2])
            with self.lock:
                for i, v in enumerate(values):
                    idx = start_addr + i
                    if idx < len(self.holding_registers):
                        self.holding_registers[idx] = v
                    if idx == 4:
                        self.pump_control = v
                        self.pump_state = 1 if v != 0 else 0
            print(f"  [SCADA WRITE] Запись нескольких регистров start={start_addr}, vals={values}")
            return struct.pack('>BHH', fc, start_addr, qty)
            
        else:
            return bytes([fc | 0x80, 0x01])  # Exception: Illegal Function

    def handle_client(self, client_sock, client_addr):
        # print(f"-> Подключен клиент SCADA: {client_addr}")
        client_sock.settimeout(3.0)
        while self.running:
            try:
                # Modbus TCP MBAP Header = 7 байт:
                # TID (2), PID (2), Length (2), UnitID (1)
                header = client_sock.recv(7)
                if not header or len(header) < 7:
                    break
                tid, pid, length, uid = struct.unpack('>HHHB', header)
                pdu_len = length - 1
                if pdu_len <= 0:
                    continue
                pdu = b''
                while len(pdu) < pdu_len:
                    chunk = client_sock.recv(pdu_len - len(pdu))
                    if not chunk:
                        break
                    pdu += chunk
                if len(pdu) < pdu_len:
                    break
                    
                fc = pdu[0]
                resp_pdu = self.process_modbus_pdu(uid, fc, pdu[1:])
                resp_mbap = struct.pack('>HHHB', tid, pid, len(resp_pdu) + 1, uid)
                client_sock.sendall(resp_mbap + resp_pdu)
            except (socket.timeout, ConnectionResetError, BrokenPipeError):
                break
            except Exception as e:
                # print(f"Ошибка сокета: {e}")
                break
        try:
            client_sock.close()
        except Exception:
            pass

    def server_thread_func(self):
        server_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server_sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            server_sock.bind((self.host, self.port))
            server_sock.listen(5)
            server_sock.settimeout(1.0)
            print(f"[*] Modbus TCP Сервер запущен на {self.host}:{self.port} (ПЛК активен)")
        except Exception as e:
            print(f"[!] Ошибка запуска сервера на порту {self.port}: {e}")
            if self.port == 502:
                print("    Попробуйте запустить консоль с правами Администратора или указать порт 5020.")
            self.running = False
            return

        while self.running:
            try:
                client_sock, client_addr = server_sock.accept()
                t = threading.Thread(target=self.handle_client, args=(client_sock, client_addr), daemon=True)
                t.start()
            except socket.timeout:
                continue
            except Exception:
                break
        server_sock.close()

    def simulation_thread_func(self):
        while self.running:
            self.simulation_step()
            time.sleep(0.5)

    def start(self):
        self.running = True
        self.srv_t = threading.Thread(target=self.server_thread_func, daemon=True)
        self.srv_t.start()
        self.sim_t = threading.Thread(target=self.simulation_thread_func, daemon=True)
        self.sim_t.start()

    def stop(self):
        self.running = False

def print_status(sim):
    cfg = VARIANTS_CONFIG[sim.variant]
    reg0 = sim.holding_registers[0]
    reg1 = sim.holding_registers[1]
    reg2 = sim.holding_registers[2]
    mode_str = "НОРМАЛЬНЫЙ" if sim.mode == 'NORMAL' else "!!! АВАРИЯ (ТРЕВОГА) !!!"
    pump_str = "РАБОТА (1)" if sim.pump_state == 1 else "ОСТАНОВ (0)"
    
    print("\n" + "=" * 65)
    print(f" СТАТУС ВИРТУАЛЬНОГО ПЛК (Modbus TCP 127.0.0.1:{sim.port})")
    print(f" {cfg['name']}")
    print(f" Режим работы: {mode_str} | Состояние насоса: {pump_str}")
    print("-" * 65)
    print(f" [Регистр 0] {cfg['p1_name']}: {sim.val1:6.2f} {cfg['p1_unit']:<5} (в регистре: {reg0:>5}) | Порог: {cfg['p1_thresh']}")
    print(f" [Регистр 1] {cfg['p2_name']}: {sim.val2:6.2f} {cfg['p2_unit']:<5} (в регистре: {reg1:>5}) | Порог: {cfg['p2_thresh']}")
    print(f" [Регистр 2] {cfg['p3_name']}: {sim.val3:6.2f} {cfg['p3_unit']:<5} (в регистре: {reg2:>5}) | Порог: {cfg['p3_thresh']}")
    print(f" [Регистр 3] Статус насоса: {sim.pump_state} | [Регистр 4] Команда SCADA: {sim.pump_control}")
    print("=" * 65)
    print("Команды: [N] Норма | [A] Авария | [P] Переключить насос | [1-5] Сменить вариант | [Q] Выход")

def main():
    port = 502
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            pass
            
    sim = ModbusPLCSimulator(host='0.0.0.0', port=port, variant=1)
    sim.start()
    time.sleep(0.3)
    
    print("\n" + "#" * 65)
    print("  ИМИТАТОР ВИРТУАЛЬНОГО ПЛК ДЛЯ RAPID SCADA (ЛР №3 АСУСН)")
    print("  Студенты: Гайнанов Д.И., Гарифуллин К.Р. (гр. ЭАС-514С)")
    print("#" * 65)
    
    try:
        while sim.running:
            print_status(sim)
            cmd = input("Введите команду > ").strip().lower()
            if cmd == 'q':
                break
            elif cmd == 'n':
                sim.mode = 'NORMAL'
                print(">> Установлен НОРМАЛЬНЫЙ режим работы.")
            elif cmd == 'a':
                sim.mode = 'ALARM'
                print(">> Установлен АВАРИЙНЫЙ режим! Параметры превышают уставки тревог.")
            elif cmd == 'p':
                sim.pump_state = 1 if sim.pump_state == 0 else 0
                sim.pump_control = sim.pump_state
                print(f">> Насос переключен: {'РАБОТА' if sim.pump_state else 'ОСТАНОВ'}")
            elif cmd in ('1', '2', '3', '4', '5'):
                sim.variant = int(cmd)
                cfg = VARIANTS_CONFIG[sim.variant]
                sim.val1 = (cfg['p1_norm'][0] + cfg['p1_norm'][1]) / 2.0
                sim.val2 = (cfg['p2_norm'][0] + cfg['p2_norm'][1]) / 2.0
                sim.val3 = (cfg['p3_norm'][0] + cfg['p3_norm'][1]) / 2.0
                print(f">> Выбран вариант {sim.variant}: {cfg['name']}")
            time.sleep(0.5)
    except (KeyboardInterrupt, EOFError):
        pass
    finally:
        sim.stop()
        print("\nСервер Modbus TCP остановлен.")

if __name__ == '__main__':
    main()

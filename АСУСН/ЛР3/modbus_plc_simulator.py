# -*- coding: utf-8 -*-
"""
Имитатор ПЛК (Modbus TCP Slave / Unit ID 1) для Лабораторной работы №3 (Вариант A).
По умолчанию настроен под Вариант 2 (Расход, Уровень, Вибрация), поддерживает варианты 1–5.

Карта регистров Modbus TCP (адресация с 0, Holding Registers FC 03 / FC 06 / FC 16):
  Reg 0  (uint16)  : Параметр 1 * scale1 (для Вар. 2: Расход Q * 1, м³/ч)
  Reg 1  (uint16)  : Параметр 2 * scale2 (для Вар. 2: Уровень L * 10, 0.1%)
  Reg 2  (uint16)  : Параметр 3 * scale3 (для Вар. 2: Вибрация V * 100, 0.01 мм/с)
  Reg 3  (uint16)  : Состояние агрегата / насоса (0 = СТОП, 1 = ПУСК)
  Reg 4  (uint16)  : Команда управления (R/W через FC 06 / FC 03: 0 = СТОП, 1 = ПУСК)
  Reg 10..11 (f32) : Параметр 1 в формате IEEE 754 Float32 Big-Endian (Q, м³/ч)
  Reg 12..13 (f32) : Параметр 2 в формате IEEE 754 Float32 Big-Endian (L, %)
  Reg 14..15 (f32) : Параметр 3 в формате IEEE 754 Float32 Big-Endian (V, мм/с)
  Coil 0  (bool)   : Дублирует команду ПУСК/СТОП через FC 01 / FC 05

Локальный HTTP API управления (127.0.0.1:5502):
  GET  /status       -> текущее состояние и журнал команд в JSON
  POST /mode/normal  -> штатный режим
  POST /mode/alarm   -> аварийный режим (превышение уставок)
  POST /pump/1       -> команда ПУСК (1)
  POST /pump/0       -> команда СТОП (0)
"""

from __future__ import annotations

import argparse
from http.server import BaseHTTPRequestHandler, HTTPServer
import json
import math
import msvcrt
import random
import socket
import struct
import sys
import threading
import time
from typing import Any


if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


VARIANTS: dict[int, dict[str, Any]] = {
    1: {
        "name": "Вариант 1 (Температура, Уровень, Давление)",
        "p1_name": "Температура", "p1_unit": "°C",   "p1_range": (0, 150), "p1_norm": (45.0, 65.0),   "p1_alarm": 105.0, "p1_thresh": 90.0,  "p1_scale": 10.0,
        "p2_name": "Уровень",     "p2_unit": "%",    "p2_range": (0, 100), "p2_norm": (50.0, 65.0),   "p2_alarm": 88.0,  "p2_thresh": 80.0,  "p2_scale": 10.0,
        "p3_name": "Давление",    "p3_unit": "кПа",  "p3_range": (0, 1000),"p3_norm": (350.0, 480.0), "p3_alarm": 860.0, "p3_thresh": 800.0, "p3_scale": 1.0,
    },
    2: {
        "name": "Вариант 2 (Расход, Уровень, Вибрация)",
        "p1_name": "Расход",      "p1_unit": "м³/ч", "p1_range": (0, 500), "p1_norm": (260.0, 315.0), "p1_alarm": 452.0, "p1_thresh": 400.0, "p1_scale": 1.0,
        "p2_name": "Уровень",     "p2_unit": "%",    "p2_range": (0, 100), "p2_norm": (58.0, 66.0),   "p2_alarm": 72.0,  "p2_thresh": 85.0,  "p2_scale": 10.0,
        "p3_name": "Вибрация",    "p3_unit": "мм/с", "p3_range": (0, 20),  "p3_norm": (4.4, 6.4),     "p3_alarm": 17.6,  "p3_thresh": 15.0,  "p3_scale": 100.0,
    },
    3: {
        "name": "Вариант 3 (Мощность, Ток, Напряжение)",
        "p1_name": "Мощность",    "p1_unit": "кВт",  "p1_range": (0, 100), "p1_norm": (40.0, 58.0),   "p1_alarm": 86.0,  "p1_thresh": 80.0,  "p1_scale": 10.0,
        "p2_name": "Ток",         "p2_unit": "А",    "p2_range": (0, 200), "p2_norm": (85.0, 115.0),  "p2_alarm": 168.0, "p2_thresh": 150.0, "p2_scale": 10.0,
        "p3_name": "Напряжение",  "p3_unit": "В",    "p3_range": (0, 450), "p3_norm": (375.0, 385.0), "p3_alarm": 330.0, "p3_thresh": 350.0, "p3_scale": 10.0,
    },
    4: {
        "name": "Вариант 4 (pH, Температура, Проводимость)",
        "p1_name": "pH",          "p1_unit": "pH",   "p1_range": (0, 14),  "p1_norm": (6.8, 7.3),     "p1_alarm": 9.4,   "p1_thresh": 8.5,   "p1_scale": 100.0,
        "p2_name": "Температура", "p2_unit": "°C",   "p2_range": (0, 100), "p2_norm": (22.0, 28.0),   "p2_alarm": 52.0,  "p2_thresh": 45.0,  "p2_scale": 10.0,
        "p3_name": "Проводимость","p3_unit": "мкСм", "p3_range": (0, 2000),"p3_norm": (420.0, 560.0), "p3_alarm": 1650.0,"p3_thresh": 1500.0,"p3_scale": 1.0,
    },
    5: {
        "name": "Вариант 5 (Скорость, Момент, Темп. обмотки)",
        "p1_name": "Скорость",     "p1_unit": "об/мин","p1_range": (0, 3000),"p1_norm": (1440.0, 1500.0),"p1_alarm": 2650.0,"p1_thresh": 2500.0,"p1_scale": 1.0,
        "p2_name": "Момент",       "p2_unit": "Н·м",   "p2_range": (0, 500), "p2_norm": (190.0, 240.0),   "p2_alarm": 430.0, "p2_thresh": 400.0, "p2_scale": 1.0,
        "p3_name": "Темп. обмотки","p3_unit": "°C",    "p3_range": (0, 180), "p3_norm": (65.0, 82.0),     "p3_alarm": 148.0, "p3_thresh": 130.0, "p3_scale": 10.0,
    },
}


class PlcState:
    """Потокобезопасное состояние регистров и параметров ПЛК."""

    def __init__(self, variant_id: int, initial_mode: str = "normal"):
        self.lock = threading.Lock()
        self.cfg = VARIANTS[variant_id]
        self.variant_id = variant_id
        self.mode = initial_mode  # normal | alarm
        self.pump_state = 1       # 1 = ПУСК (работа), 0 = СТОП
        self.pump_control = 1
        self.coils = [False] * 64
        self.coils[0] = True
        self.holding_regs = [0] * 64
        self.val1 = 0.0
        self.val2 = 0.0
        self.val3 = 0.0
        self.t0 = time.time()
        self.cmd_log: list[dict[str, Any]] = []
        self.running = True
        self.step()

    @staticmethod
    def _float_to_regs(val: float) -> tuple[int, int]:
        packed = struct.pack(">f", float(val))
        return struct.unpack(">HH", packed)

    def step(self) -> None:
        with self.lock:
            t = time.time() - self.t0
            c = self.cfg

            if self.mode == "alarm":
                target1 = c["p1_alarm"] + 8.0 * math.sin(t * 0.45) + random.uniform(-2.0, 2.0)
                target2 = c["p2_alarm"] + 1.8 * math.cos(t * 0.30) + random.uniform(-0.4, 0.4)
                target3 = c["p3_alarm"] + 0.55 * math.sin(t * 0.60) + random.uniform(-0.15, 0.15)
            elif self.pump_state == 0:
                target1 = 0.0
                target2 = sum(c["p2_norm"]) * 0.5 + 1.2 * math.sin(t * 0.18) + random.uniform(-0.2, 0.2)
                target3 = 0.4 + random.uniform(-0.08, 0.08)
            else:
                m1 = sum(c["p1_norm"]) * 0.5
                a1 = (c["p1_norm"][1] - c["p1_norm"][0]) * 0.45
                m2 = sum(c["p2_norm"]) * 0.5
                a2 = (c["p2_norm"][1] - c["p2_norm"][0]) * 0.45
                m3 = sum(c["p3_norm"]) * 0.5
                a3 = (c["p3_norm"][1] - c["p3_norm"][0]) * 0.45
                target1 = m1 + a1 * math.sin(t * 0.25) + random.uniform(-a1 * 0.12, a1 * 0.12)
                target2 = m2 + a2 * math.cos(t * 0.20) + random.uniform(-a2 * 0.10, a2 * 0.10)
                target3 = m3 + a3 * math.sin(t * 0.35 + 1.0) + random.uniform(-a3 * 0.10, a3 * 0.10)

            self.val1 = round(max(float(c["p1_range"][0]), min(float(c["p1_range"][1]), target1)), 2)
            self.val2 = round(max(float(c["p2_range"][0]), min(float(c["p2_range"][1]), target2)), 2)
            self.val3 = round(max(float(c["p3_range"][0]), min(float(c["p3_range"][1]), target3)), 2)

            self.holding_regs[0] = max(0, min(65535, int(round(self.val1 * c["p1_scale"]))))
            self.holding_regs[1] = max(0, min(65535, int(round(self.val2 * c["p2_scale"]))))
            self.holding_regs[2] = max(0, min(65535, int(round(self.val3 * c["p3_scale"]))))
            self.holding_regs[3] = self.pump_state
            self.holding_regs[4] = self.pump_control
            self.coils[0] = bool(self.pump_state)

            r10, r11 = self._float_to_regs(self.val1)
            r12, r13 = self._float_to_regs(self.val2)
            r14, r15 = self._float_to_regs(self.val3)
            self.holding_regs[10], self.holding_regs[11] = r10, r11
            self.holding_regs[12], self.holding_regs[13] = r12, r13
            self.holding_regs[14], self.holding_regs[15] = r14, r15

    def set_mode(self, mode: str) -> None:
        with self.lock:
            self.mode = mode
        self.step()
        print(f"[{time.strftime('%H:%M:%S')}] Режим ПЛК -> {mode.upper()}")

    def apply_pump_cmd(self, source: str, reg_or_coil: int, val: int) -> None:
        ts = time.strftime("%H:%M:%S")
        with self.lock:
            norm_val = 1 if val > 0 else 0
            self.pump_control = norm_val
            self.pump_state = norm_val
            self.holding_regs[3] = norm_val
            self.holding_regs[4] = norm_val
            self.coils[0] = bool(norm_val)
            entry = {
                "time": ts,
                "source": source,
                "addr": reg_or_coil,
                "value": norm_val,
                "text": "ПУСК" if norm_val else "СТОП",
            }
            self.cmd_log.append(entry)
        self.step()
        print(f"[{ts}] КОМАНДА {source} (addr={reg_or_coil}): -> {entry['text']} ({norm_val})")

    def snapshot(self) -> dict[str, Any]:
        with self.lock:
            return {
                "variant": self.variant_id,
                "mode": self.mode,
                "pump_state": self.pump_state,
                "pump_control": self.pump_control,
                "val1": self.val1,
                "val2": self.val2,
                "val3": self.val3,
                "regs_0_4": self.holding_regs[0:5],
                "regs_10_15": self.holding_regs[10:16],
                "cmd_log": list(self.cmd_log[-20:]),
            }


def handle_modbus_pdu(state: PlcState, pdu: bytes) -> bytes:
    """Обработка Modbus PDU: FC 01, 02, 03, 04, 05, 06, 15, 16."""
    if not pdu:
        return b""
    fc = pdu[0]
    try:
        if fc in (0x01, 0x02) and len(pdu) >= 5:
            start, qty = struct.unpack(">HH", pdu[1:5])
            if qty < 1 or qty > 2000 or start + qty > len(state.coils):
                return bytes([fc | 0x80, 0x02])
            with state.lock:
                bits = state.coils[start:start + qty]
            byte_cnt = (qty + 7) // 8
            buf = bytearray(byte_cnt)
            for i, b in enumerate(bits):
                if b:
                    buf[i // 8] |= 1 << (i % 8)
            return bytes([fc, byte_cnt]) + bytes(buf)

        if fc in (0x03, 0x04) and len(pdu) >= 5:
            start, qty = struct.unpack(">HH", pdu[1:5])
            if qty < 1 or qty > 125 or start + qty > len(state.holding_regs):
                return bytes([fc | 0x80, 0x02])
            with state.lock:
                regs = state.holding_regs[start:start + qty]
            data = b"".join(struct.pack(">H", r) for r in regs)
            return bytes([fc, len(data)]) + data

        if fc == 0x05 and len(pdu) >= 5:
            addr, raw_val = struct.unpack(">HH", pdu[1:5])
            if addr >= len(state.coils):
                return bytes([fc | 0x80, 0x02])
            bit_val = 1 if raw_val == 0xFF00 else 0
            if addr == 0:
                state.apply_pump_cmd("FC05 WriteSingleCoil", addr, bit_val)
            else:
                with state.lock:
                    state.coils[addr] = bool(bit_val)
            return pdu[:5]

        if fc == 0x06 and len(pdu) >= 5:
            addr, val = struct.unpack(">HH", pdu[1:5])
            if addr >= len(state.holding_regs):
                return bytes([fc | 0x80, 0x02])
            if addr in (3, 4):
                state.apply_pump_cmd("FC06 WriteSingleReg", addr, val)
            else:
                with state.lock:
                    state.holding_regs[addr] = val
            return pdu[:5]

        if fc == 0x10 and len(pdu) >= 6:
            start, qty, byte_cnt = struct.unpack(">HHB", pdu[1:6])
            if start + qty > len(state.holding_regs) or len(pdu) < 6 + byte_cnt:
                return bytes([fc | 0x80, 0x02])
            for i in range(qty):
                val = struct.unpack(">H", pdu[6 + 2 * i : 8 + 2 * i])[0]
                if start + i in (3, 4):
                    state.apply_pump_cmd("FC16 WriteMultiReg", start + i, val)
                else:
                    with state.lock:
                        state.holding_regs[start + i] = val
            return struct.pack(">BHH", fc, start, qty)

        return bytes([fc | 0x80, 0x01])
    except Exception:
        return bytes([fc | 0x80, 0x04])


def recv_exact(conn: socket.socket, n: int) -> bytes:
    buf = bytearray()
    while len(buf) < n:
        chunk = conn.recv(n - len(buf))
        if not chunk:
            return b""
        buf.extend(chunk)
    return bytes(buf)


def client_worker(conn: socket.socket, addr: tuple[str, int], state: PlcState) -> None:
    try:
        conn.settimeout(10.0)
        while state.running:
            hdr = recv_exact(conn, 7)
            if not hdr:
                break
            tx_id, proto_id, length, unit_id = struct.unpack(">HHHB", hdr)
            if length < 2 or length > 260:
                break
            pdu = recv_exact(conn, length - 1)
            if not pdu:
                break
            resp_pdu = handle_modbus_pdu(state, pdu)
            mbap = struct.pack(">HHHB", tx_id, proto_id, len(resp_pdu) + 1, unit_id)
            conn.sendall(mbap + resp_pdu)
    except Exception:
        pass
    finally:
        try:
            conn.close()
        except Exception:
            pass


def start_control_http(state: PlcState, ctrl_port: int) -> HTTPServer:
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, format: str, *args: Any) -> None:
            return

        def _send_json(self, code: int, obj: dict[str, Any]) -> None:
            raw = json.dumps(obj, ensure_ascii=False).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)

        def do_GET(self) -> None:
            if self.path == "/status":
                self._send_json(200, state.snapshot())
            else:
                self._send_json(404, {"error": "not found"})

        def do_POST(self) -> None:
            if self.path == "/mode/normal":
                state.set_mode("normal")
                self._send_json(200, state.snapshot())
            elif self.path == "/mode/alarm":
                state.set_mode("alarm")
                self._send_json(200, state.snapshot())
            elif self.path == "/pump/1":
                state.apply_pump_cmd("HTTP", 4, 1)
                self._send_json(200, state.snapshot())
            elif self.path == "/pump/0":
                state.apply_pump_cmd("HTTP", 4, 0)
                self._send_json(200, state.snapshot())
            else:
                self._send_json(404, {"error": "not found"})

    srv = HTTPServer(("127.0.0.1", ctrl_port), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


def main() -> None:
    parser = argparse.ArgumentParser(description="Modbus TCP PLC simulator for Lab 3")
    parser.add_argument("--variant", type=int, default=2, choices=list(VARIANTS.keys()))
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=502)
    parser.add_argument("--ctrl-port", type=int, default=5502)
    parser.add_argument("--mode", choices=["normal", "alarm"], default="normal")
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args()

    state = PlcState(args.variant, initial_mode=args.mode)
    start_control_http(state, args.ctrl_port)

    srv_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    srv_sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    srv_sock.bind((args.host, args.port))
    srv_sock.listen(16)
    srv_sock.settimeout(0.5)

    def accept_loop() -> None:
        while state.running:
            try:
                conn, addr = srv_sock.accept()
                threading.Thread(target=client_worker, args=(conn, addr, state), daemon=True).start()
            except socket.timeout:
                continue
            except OSError:
                break

    threading.Thread(target=accept_loop, daemon=True).start()

    c = state.cfg
    print("=" * 72)
    print(f"Modbus TCP ПЛК-имитатор: {c['name']}")
    print(f"Modbus TCP: {args.host}:{args.port} (Unit ID=1) | Control HTTP: 127.0.0.1:{args.ctrl_port}")
    print(f"Регистры 10..15 (Float32) и 0..4 (Uint16/Ctrl): Q>{c['p1_thresh']} {c['p1_unit']}, V>{c['p3_thresh']} {c['p3_unit']}")
    if sys.stdin and sys.stdin.isatty():
        print("Управление с клавиатуры: [n] Норма, [a] Авария, [1] ПУСК, [0] СТОП, [q] Выход")
    print("=" * 72)

    last_print = 0.0
    try:
        while state.running:
            state.step()
            now = time.time()
            if not args.quiet and now - last_print >= 2.0:
                last_print = now
                snap = state.snapshot()
                pump_txt = "РАБОТА (1)" if snap["pump_state"] else "СТОП (0)"
                print(
                    f"[{time.strftime('%H:%M:%S')}] Режим={snap['mode'].upper():6s} | "
                    f"{c['p1_name']}={snap['val1']:6.2f} {c['p1_unit']} | "
                    f"{c['p2_name']}={snap['val2']:5.2f} {c['p2_unit']} | "
                    f"{c['p3_name']}={snap['val3']:5.2f} {c['p3_unit']} | "
                    f"Агрегат={pump_txt}",
                    flush=True,
                )
            if sys.stdin and sys.stdin.isatty() and msvcrt.kbhit():
                ch = msvcrt.getwch().lower()
                if ch == "n":
                    state.set_mode("normal")
                elif ch == "a":
                    state.set_mode("alarm")
                elif ch == "1":
                    state.apply_pump_cmd("Key", 4, 1)
                elif ch == "0":
                    state.apply_pump_cmd("Key", 4, 0)
                elif ch == "q":
                    state.running = False
                    break
            time.sleep(0.5)
    except KeyboardInterrupt:
        pass
    finally:
        state.running = False
        try:
            srv_sock.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()

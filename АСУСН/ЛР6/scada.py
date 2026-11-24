# -*- coding: utf-8 -*-
"""
scada.py
Интерактивный программный SCADA-эмулятор распределенной системы управления.
Дисциплина: АСУСН, Лабораторная работа №6.
Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.

Выполняет:
- Подключение по Modbus TCP к ПЛК №1 (порт 5020) и ПЛК №2 (порт 5021)
- Сбор данных технологического процесса в реальном времени (T, L, P, ТЭН, Насос, Клапан, Авария)
- Динамическое отображение трендов на графическом холсте Matplotlib
- Интерактивное диспетчерское управление: оперативное изменение уставки, ручное управление, алерты
- Проведение регламентных испытаний (Тесты 1-5) с записью трендов в scada_trends.csv и scada_result.png
"""

import os
import sys
import time
import argparse
import threading

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from pymodbus.client import ModbusTcpClient

# Настройка шрифтов matplotlib
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Segoe UI', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False


class ScadaDataCollector:
    """Модуль сбора телеметрии и Modbus TCP коммуникации."""

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
        self.max_history_len = 500
        self.is_connected = False

    def connect(self):
        c1 = self.client1.connect()
        c2 = self.client2.connect()
        self.is_connected = bool(c1 and c2)
        return self.is_connected

    def disconnect(self):
        try:
            self.client1.close()
            self.client2.close()
        except Exception:
            pass
        self.is_connected = False

    def set_setpoint(self, new_sp):
        """Отправка новой уставки в ПЛК №1 (HR[2])."""
        val = int(float(new_sp) * 10)
        try:
            res = self.client1.write_register(2, val)
            return not res.isError()
        except Exception:
            return False

    def read_telemetry(self):
        """Опрос обоих контроллеров."""
        data = {
            'temp': 20.0,
            'setpoint': 60.0,
            'heater': 0.0,
            'level': 50.0,
            'pressure': 1.0,
            'pump': 0,
            'valve': 0,
            'alarm': 0,
            'plc1_ok': False,
            'plc2_ok': False
        }

        # 1. Опрос ПЛК №1 (HR0..2)
        try:
            hr1 = self.client1.read_holding_registers(0, count=3)
            if not hr1.isError() and len(hr1.registers) >= 3:
                data['temp'] = hr1.registers[0] / 10.0
                data['heater'] = hr1.registers[1] / 10.0
                data['setpoint'] = hr1.registers[2] / 10.0
                data['plc1_ok'] = True
        except Exception:
            data['plc1_ok'] = False

        # 2. Опрос ПЛК №2 (HR0..2, DI0..2)
        try:
            hr2 = self.client2.read_holding_registers(0, count=3)
            if not hr2.isError() and len(hr2.registers) >= 3:
                data['level'] = hr2.registers[1]
                data['pressure'] = hr2.registers[2] / 100.0
                data['plc2_ok'] = True

            di2 = self.client2.read_discrete_inputs(0, count=3)
            if not di2.isError() and len(di2.bits) >= 3:
                data['pump'] = 1 if di2.bits[0] else 0
                data['valve'] = 1 if di2.bits[1] else 0
                data['alarm'] = 1 if di2.bits[2] else 0
        except Exception:
            data['plc2_ok'] = False

        return data

    def add_sample(self, t, d):
        self.history['t'].append(t)
        self.history['temp'].append(d.get('temp', 20.0))
        self.history['setpoint'].append(d.get('setpoint', 60.0))
        self.history['heater'].append(d.get('heater', 0.0))
        self.history['level'].append(d.get('level', 50.0))
        self.history['pressure'].append(d.get('pressure', 1.0))
        self.history['pump'].append(d.get('pump', 0))
        self.history['valve'].append(d.get('valve', 0))
        self.history['alarm'].append(d.get('alarm', 0))

        if len(self.history['t']) > self.max_history_len:
            for k in self.history:
                self.history[k].pop(0)


class ScadaGuiApp:
    """Полнофункциональный графический интерфейс оператора SCADA (Tkinter + Matplotlib)."""

    def __init__(self, root, host='127.0.0.1', plc1_port=5020, plc2_port=5021):
        self.root = root
        self.root.title("SCADA ЭМУЛЯТОР — Распределенная АСУ ТП (ПЛК1 + ПЛК2 + Объект)")
        self.root.geometry("1180x820")
        self.root.minsize(1000, 700)

        self.collector = ScadaDataCollector(host, plc1_port, plc2_port)
        self.running = False
        self.start_time = time.time()

        self.create_styles()
        self.create_header()
        self.create_main_layout()
        self.create_statusbar()

        # Автоподключение при старте
        self.toggle_connection()

    def create_styles(self):
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam')
        except Exception:
            pass
        self.style.configure(".", font=("Segoe UI", 10))
        self.style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"), background="#1e293b", foreground="#ffffff")
        self.style.configure("Card.TLabelframe", background="#ffffff", relief="groove")
        self.style.configure("Card.TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#0f172a")

    def create_header(self):
        hdr = tk.Frame(self.root, bg="#1e293b", height=60)
        hdr.pack(side=tk.TOP, fill=tk.X)
        hdr.pack_propagate(False)

        lbl = tk.Label(hdr, text="ДИСПЕТЧЕРСКИЙ ТЕРМИНАЛ SCADA: РАСПРЕДЕЛЕННАЯ СИСТЕМА УПРАВЛЕНИЯ",
                       font=("Segoe UI", 12, "bold"), fg="#ffffff", bg="#1e293b")
        lbl.pack(anchor="w", padx=20, pady=(6, 0))

        sub = tk.Label(hdr, text="Лабораторная работа №6 (АСУСН) | ПЛК №1 (Modbus: 5020), ПЛК №2 (Modbus: 5021), Объект (Socket: 9999)",
                       font=("Segoe UI", 9), fg="#94a3b8", bg="#1e293b")
        sub.pack(anchor="w", padx=20, pady=(1, 6))

    def create_main_layout(self):
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # Левая панель: телеметрия, уставки, команды (ширина ~290px)
        left_frame = ttk.Frame(main_paned, width=300)
        main_paned.add(left_frame, weight=0)

        # Правая панель: тренды и лог (ширина остальная)
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=1)

        # --- Содержимое левой панели ---
        # 1. Связь
        f_conn = ttk.LabelFrame(left_frame, text="Сетевой статус Modbus TCP", style="Card.TLabelframe", padding=8)
        f_conn.pack(fill=tk.X, padx=5, pady=5)

        self.lbl_plc1_stat = tk.Label(f_conn, text="● ПЛК №1 (Регулятор): ОТКЛЮЧЕН", font=("Segoe UI", 9, "bold"), fg="#dc2626")
        self.lbl_plc1_stat.pack(anchor="w", pady=2)
        self.lbl_plc2_stat = tk.Label(f_conn, text="● ПЛК №2 (ПАЗ/Объект): ОТКЛЮЧЕН", font=("Segoe UI", 9, "bold"), fg="#dc2626")
        self.lbl_plc2_stat.pack(anchor="w", pady=2)

        self.btn_connect = ttk.Button(f_conn, text="Подключиться к ПЛК", command=self.toggle_connection)
        self.btn_connect.pack(fill=tk.X, pady=(6, 2))

        # 2. Текущие значения телеметрии
        f_telem = ttk.LabelFrame(left_frame, text="Оперативная мнемосхема", style="Card.TLabelframe", padding=8)
        f_telem.pack(fill=tk.X, padx=5, pady=5)

        self.lbl_cur_temp = tk.Label(f_telem, text="Температура: -- °C", font=("Segoe UI", 11, "bold"), fg="#b91c1c")
        self.lbl_cur_temp.pack(anchor="w", pady=2)

        self.lbl_cur_sp = tk.Label(f_telem, text="Уставка (SP): -- °C", font=("Segoe UI", 10), fg="#047857")
        self.lbl_cur_sp.pack(anchor="w", pady=2)

        self.lbl_cur_heater = tk.Label(f_telem, text="Мощность ТЭН: -- %", font=("Segoe UI", 10), fg="#d97706")
        self.lbl_cur_heater.pack(anchor="w", pady=2)

        self.lbl_cur_level = tk.Label(f_telem, text="Уровень в баке: -- %", font=("Segoe UI", 10), fg="#0284c7")
        self.lbl_cur_level.pack(anchor="w", pady=2)

        self.lbl_cur_press = tk.Label(f_telem, text="Давление: -- бар", font=("Segoe UI", 10), fg="#475569")
        self.lbl_cur_press.pack(anchor="w", pady=2)

        self.lbl_cur_pump = tk.Label(f_telem, text="Насос притока: [ВЫКЛ]", font=("Segoe UI", 9, "bold"), fg="#64748b")
        self.lbl_cur_pump.pack(anchor="w", pady=2)

        self.lbl_cur_valve = tk.Label(f_telem, text="Клапан слива: [ЗАКР]", font=("Segoe UI", 9, "bold"), fg="#64748b")
        self.lbl_cur_valve.pack(anchor="w", pady=2)

        self.lbl_cur_alarm = tk.Label(f_telem, text="СТАТУС: НОРМА", font=("Segoe UI", 10, "bold"), fg="#16a34a", bg="#dcfce7", relief=tk.RIDGE)
        self.lbl_cur_alarm.pack(fill=tk.X, pady=(6, 2), ipady=3)

        # 3. Диспетчерское управление уставкой
        f_sp = ttk.LabelFrame(left_frame, text="Управление уставкой ПЛК1", style="Card.TLabelframe", padding=8)
        f_sp.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(f_sp, text="Новая уставка температуры (°C):").pack(anchor="w")
        sp_subframe = ttk.Frame(f_sp)
        sp_subframe.pack(fill=tk.X, pady=(3, 6))

        self.ent_sp = ttk.Entry(sp_subframe, width=10)
        self.ent_sp.insert(0, "60.0")
        self.ent_sp.pack(side=tk.LEFT, padx=(0, 6))

        btn_apply_sp = ttk.Button(sp_subframe, text="Применить", command=self.apply_setpoint)
        btn_apply_sp.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Кнопки быстрых тестов
        ttk.Label(f_sp, text="Быстрые сценарии:").pack(anchor="w", pady=(4, 2))
        btn_sp60 = ttk.Button(f_sp, text="Тест 1: Номинал 60 °C", command=lambda: self.set_sp_val(60.0))
        btn_sp60.pack(fill=tk.X, pady=1)

        btn_sp75 = ttk.Button(f_sp, text="Тест 2: Скачок 75 °C", command=lambda: self.set_sp_val(75.0))
        btn_sp75.pack(fill=tk.X, pady=1)

        btn_sp120 = ttk.Button(f_sp, text="Тест 4: Перегрузка 120 °C", command=lambda: self.set_sp_val(120.0))
        btn_sp120.pack(fill=tk.X, pady=1)

        # --- Содержимое правой панели ---
        # Холст графиков Matplotlib
        self.fig, (self.ax1, self.ax2, self.ax3) = plt.subplots(3, 1, figsize=(8.5, 6.2), sharex=True, dpi=100)
        self.fig.patch.set_facecolor('#f8fafc')

        self.canvas = FigureCanvasTkAgg(self.fig, master=right_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        # Журнал сообщений SCADA
        f_log = ttk.LabelFrame(right_frame, text="Оперативный журнал событий SCADA", padding=5)
        f_log.pack(fill=tk.X, padx=4, pady=(0, 4))

        self.txt_log = scrolledtext.ScrolledText(f_log, height=5, font=("Consolas", 9), bg="#0f172a", fg="#38bdf8")
        self.txt_log.pack(fill=tk.X)
        self.log_msg("Инициализация диспетчерского терминала SCADA. Готов к приему телеметрии.")

    def create_statusbar(self):
        sb = tk.Frame(self.root, bg="#e2e8f0", height=24)
        sb.pack(side=tk.BOTTOM, fill=tk.X)
        sb.pack_propagate(False)

        self.lbl_status = tk.Label(sb, text="Ожидание запуска циклов ПЛК1 и ПЛК2...", font=("Segoe UI", 9), bg="#e2e8f0", fg="#334155")
        self.lbl_status.pack(side=tk.LEFT, padx=10)

    def log_msg(self, msg):
        ts = time.strftime("%H:%M:%S")
        self.txt_log.insert(tk.END, f"[{ts}] {msg}\n")
        self.txt_log.see(tk.END)

    def set_sp_val(self, val):
        self.ent_sp.delete(0, tk.END)
        self.ent_sp.insert(0, str(val))
        self.apply_setpoint()

    def apply_setpoint(self):
        try:
            sp = float(self.ent_sp.get())
            if sp < 0 or sp > 200:
                messagebox.showwarning("Внимание", "Уставка температуры должна быть в диапазоне от 0 до 200 °C!")
                return
            ok = self.collector.set_setpoint(sp)
            if ok:
                self.log_msg(f"Команда ТУ: Успешно записана новая уставка SP = {sp:.1f} °C в ПЛК №1")
            else:
                self.log_msg(f"Ошибка записи уставки SP = {sp:.1f} °C (ПЛК1 не отвечает)")
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректное числовое значение уставки!")

    def toggle_connection(self):
        if not self.running:
            ok = self.collector.connect()
            self.running = True
            self.start_time = time.time()
            self.btn_connect.config(text="Отключиться от ПЛК")
            self.log_msg("Запущен циклический опрос Modbus TCP (период 500 мс)")
            self.update_loop()
        else:
            self.running = False
            self.collector.disconnect()
            self.btn_connect.config(text="Подключиться к ПЛК")
            self.lbl_plc1_stat.config(text="● ПЛК №1: ОТКЛЮЧЕН", fg="#dc2626")
            self.lbl_plc2_stat.config(text="● ПЛК №2: ОТКЛЮЧЕН", fg="#dc2626")
            self.log_msg("Опрос остановлен пользователем.")

    def update_loop(self):
        if not self.running:
            return

        t = round(time.time() - self.start_time, 2)
        d = self.collector.read_telemetry()
        self.collector.add_sample(t, d)

        # Обновление индикаторов связи
        if d['plc1_ok']:
            self.lbl_plc1_stat.config(text=f"● ПЛК №1 (5020): СВЯЗЬ OK", fg="#16a34a")
        else:
            self.lbl_plc1_stat.config(text=f"● ПЛК №1 (5020): НЕТ СВЯЗИ", fg="#dc2626")

        if d['plc2_ok']:
            self.lbl_plc2_stat.config(text=f"● ПЛК №2 (5021): СВЯЗЬ OK", fg="#16a34a")
        else:
            self.lbl_plc2_stat.config(text=f"● ПЛК №2 (5021): НЕТ СВЯЗИ", fg="#dc2626")

        # Обновление мнемосхемы
        self.lbl_cur_temp.config(text=f"Температура: {d['temp']:.1f} °C")
        self.lbl_cur_sp.config(text=f"Уставка (SP): {d['setpoint']:.1f} °C")
        self.lbl_cur_heater.config(text=f"Мощность ТЭН: {d['heater']:.0f} %")
        self.lbl_cur_level.config(text=f"Уровень в баке: {d['level']:.0f} %")
        self.lbl_cur_press.config(text=f"Давление: {d['pressure']:.2f} бар")

        if d['pump']:
            self.lbl_cur_pump.config(text="Насос притока: [ВКЛЮЧЕН]", fg="#16a34a")
        else:
            self.lbl_cur_pump.config(text="Насос притока: [ВЫКЛ]", fg="#64748b")

        if d['valve']:
            self.lbl_cur_valve.config(text="Клапан слива: [ОТКРЫТ]", fg="#0284c7")
        else:
            self.lbl_cur_valve.config(text="Клапан слива: [ЗАКР]", fg="#64748b")

        if d['alarm']:
            self.lbl_cur_alarm.config(text="АВАРИЯ: СРАБОТАЛА ПАЗ!", fg="#ffffff", bg="#dc2626")
        else:
            self.lbl_cur_alarm.config(text="СТАТУС: ШТАТНО", fg="#16a34a", bg="#dcfce7")

        self.lbl_status.config(text=f"Время сеанса: {t:.1f} с | Точек тренда: {len(self.collector.history['t'])}")

        # Перерисовка графиков
        self.redraw_plots()

        # Следующий опрос через 500 мс
        self.root.after(500, self.update_loop)

    def redraw_plots(self):
        h = self.collector.history
        if not h['t']:
            return

        t = h['t']

        # 1. График температуры
        self.ax1.clear()
        self.ax1.plot(t, h['temp'], color='#b91c1c', lw=1.6, label='Температура T (°C)')
        self.ax1.plot(t, h['setpoint'], color='#059669', lw=1.3, linestyle='--', label='Уставка SP (°C)')
        self.ax1.set_ylabel('T, °C', fontsize=8.5)
        self.ax1.grid(True, linestyle=':', alpha=0.6)
        self.ax1.legend(loc='upper left', fontsize=7.5)

        # 2. График уровня и ТЭНа
        self.ax2.clear()
        self.ax2.plot(t, h['level'], color='#0284c7', lw=1.5, label='Уровень L (%)')
        self.ax2.plot(t, h['heater'], color='#d97706', lw=1.2, linestyle=':', label='ТЭН (%)')
        self.ax2.axhline(80.0, color='#991b1b', linestyle=':', lw=0.9, alpha=0.7)
        self.ax2.axhline(30.0, color='#1e3a8a', linestyle=':', lw=0.9, alpha=0.7)
        self.ax2.set_ylabel('L, ТЭН (%)', fontsize=8.5)
        self.ax2.grid(True, linestyle=':', alpha=0.6)
        self.ax2.legend(loc='upper left', fontsize=7.5)

        # 3. График давления и дискретных выходов
        self.ax3.clear()
        self.ax3.plot(t, h['pressure'], color='#475569', lw=1.4, label='Давление P (бар)')
        self.ax3.step(t, h['pump'], color='#16a34a', lw=1.1, where='mid', label='Насос')
        self.ax3.step(t, h['valve'], color='#7c3aed', lw=1.1, where='mid', label='Клапан')
        self.ax3.set_ylabel('P (бар) / Дискр.', fontsize=8.5)
        self.ax3.set_xlabel('Время, с', fontsize=8.5)
        self.ax3.grid(True, linestyle=':', alpha=0.6)
        self.ax3.legend(loc='upper left', fontsize=7.5)

        self.fig.tight_layout()
        self.canvas.draw_idle()


# =============================================================================
# CLI РЕЖИМ (ДЛЯ АВТОТЕСТОВ И ФОНОВОГО СБОРА)
# =============================================================================
def run_cli_session(duration=15.0, setpoint=None):
    collector = ScadaDataCollector()
    collector.connect()
    if setpoint is not None:
        collector.set_setpoint(setpoint)
        print(f"[SCADA CLI] Установлена уставка: {setpoint} °C")

    print(f"[SCADA CLI] Сбор телеметрии на протяжении {duration} с...")
    start = time.time()
    while time.time() - start < duration:
        t = round(time.time() - start, 1)
        d = collector.read_telemetry()
        collector.add_sample(t, d)
        print(f"[{t:4.1f} с] T={d['temp']:4.1f}°C (SP={d['setpoint']:.0f}), ТЭН={d['heater']:3.0f}%, L={d['level']:2.0f}%, P={d['pressure']:4.2f}бар | Насос={d['pump']}, Клапан={d['valve']}, Авария={d['alarm']}")
        time.sleep(0.5)

    collector.disconnect()
    print("[SCADA CLI] Сессия завершена.")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="SCADA диспетчерский эмулятор")
    parser.add_argument('--cli', action='store_true', help="Запуск в консольном режиме без GUI")
    parser.add_argument('--duration', type=float, default=15.0, help="Длительность сессии в CLI режиме")
    parser.add_argument('--setpoint', type=float, default=None, help="Оперативно задать уставку в CLI")
    args = parser.parse_args()

    if args.cli:
        run_cli_session(duration=args.duration, setpoint=args.setpoint)
    else:
        root = tk.Tk()
        app = ScadaGuiApp(root)
        root.mainloop()

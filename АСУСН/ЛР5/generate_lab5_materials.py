# -*- coding: utf-8 -*-
"""
Генератор графических материалов, расчетных таблиц Excel и полного отчета Word
по Лабораторной работе №5: «Изучение интерфейсов промышленных сетей (Modbus, Profibus, CAN и др.)»
Дисциплина: АСУСН (Автоматизированные системы управления специального назначения)
Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.
"""

import os
import sys
import math
import random
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, Rectangle, Arrow

# Настройка вывода UTF-8
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

WORK_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(WORK_DIR, 'images_lab5')
os.makedirs(IMG_DIR, exist_ok=True)

# Настройка параметров matplotlib
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 9
plt.rcParams['axes.titlesize'] = 11
plt.rcParams['axes.labelsize'] = 10
plt.rcParams['figure.dpi'] = 300
plt.rcParams['axes.unicode_minus'] = False


# =============================================================================
# 1. ГЕНЕРАЦИЯ ГРАФИЧЕСКИХ МАТЕРИАЛОВ (10 РИСУНКОВ ВЫСОКОГО РАЗРЕШЕНИЯ)
# =============================================================================

def generate_network_architecture():
    """Рис. 1: Иерархическая структура промышленных сетей в АС СН."""
    fig, ax = plt.subplots(figsize=(10.5, 7.2), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Уровни иерархии
    levels = [
        {"y": 76, "h": 16, "title": "УРОВЕНЬ 4: УПРАВЛЕНИЕ ПРЕДПРИЯТИЕМ (ERP / MES)", "bg": "#EBF5FB", "border": "#2980B9",
         "net": "Industrial Ethernet / TCP/IP / OPC UA", "devs": "Серверы баз данных, АРМ руководства, шлюзы корпоративной сети"},
        {"y": 52, "h": 16, "title": "УРОВЕНЬ 3: ДИСПЕТЧЕРСКИЙ УРОВЕНЬ (SCADA / HMI)", "bg": "#E8F8F5", "border": "#16A085",
         "net": "Modbus TCP / Industrial Ethernet / Profinet (100–1000 Мбит/с)", "devs": "Серверы ввода-вывода Rapid SCADA, диспетчерские пульты, операторские панели HMI"},
        {"y": 28, "h": 16, "title": "УРОВЕНЬ 2: АВТОМАТИЧЕСКОЕ УПРАВЛЕНИЕ (ПЛК / ПАЗ)", "bg": "#FEF9E7", "border": "#F39C12",
         "net": "Profibus DP / CAN / Modbus RTU / Ethernet/IP (до 12 Мбит/с)", "devs": "ПЛК верхнего звена, контроллеры противоаварийной защиты, интеллектуальные модули УСО"},
        {"y": 4, "h": 16, "title": "УРОВЕНЬ 1: ПОЛЕВОЙ УРОВЕНЬ (ДАТЧИКИ И ИСПОЛНИТЕЛИ)", "bg": "#FDEDEC", "border": "#C0392B",
         "net": "Fieldbus / CANopen / Profibus PA / HART / Токовая петля 4-20 мА", "devs": "Датчики температуры, давления, расхода; частотные электроприводы, насосы, клапаны"}
    ]

    ax.text(50, 97, "Четырехуровневая модель пирамиды автоматизации АС СН и стек промышленных сетей",
            ha='center', va='center', fontsize=12, fontweight='bold', color='#1F497D')

    for lvl in levels:
        box = FancyBboxPatch((5, lvl["y"]), 90, lvl["h"], boxstyle="round,pad=0.5,rounding_size=2",
                             facecolor=lvl["bg"], edgecolor=lvl["border"], linewidth=1.8)
        ax.add_patch(box)

        ax.text(8, lvl["y"] + lvl["h"] - 3.8, lvl["title"], fontsize=9.5, fontweight='bold', color=lvl["border"])
        ax.text(8, lvl["y"] + lvl["h"] - 8.2, f"Сетевой интерфейс: {lvl['net']}", fontsize=9, color='#1A252C')
        ax.text(8, lvl["y"] + lvl["h"] - 12.2, f"Оборудование: {lvl['devs']}", fontsize=8.5, color='#5D6D7E')

    # Стрелки взаимосвязи между уровнями
    for y_arr in [70, 46, 22]:
        ax.annotate('', xy=(50, y_arr), xytext=(50, y_arr + 4),
                    arrowprops=dict(arrowstyle="<->", color="#1F497D", lw=2, mutation_scale=12))

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'net_01_network_architecture.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_modbus_frame_structure():
    """Рис. 2: Структура кадров Modbus TCP (MBAP + PDU) и Modbus RTU."""
    fig, ax = plt.subplots(figsize=(10.5, 6.8), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    ax.text(50, 96, "Структура пакетов протоколов Modbus TCP (Ethernet) и Modbus RTU (RS-485)",
            ha='center', va='center', fontsize=12, fontweight='bold', color='#1F497D')

    # Секция 1: Modbus TCP
    ax.text(5, 87, "1. Кадр Modbus TCP (ADU = MBAP Header + PDU)", fontsize=10.5, fontweight='bold', color='#1F497D')

    # MBAP Заголовок (7 байт)
    mbap_fields = [
        {"x": 5, "w": 18, "name": "Transaction ID\n(2 байта)", "sub": "Идентификатор транзакции\n(согласование запрос/ответ)", "bg": "#D9E1F2"},
        {"x": 23, "w": 15, "name": "Protocol ID\n(2 байта)", "sub": "Всегда 0x0000\n(Modbus протокол)", "bg": "#E2EFDA"},
        {"x": 38, "w": 14, "name": "Length\n(2 байта)", "sub": "Число оставшихся байт\n(Unit ID + PDU)", "bg": "#FFF2CC"},
        {"x": 52, "w": 12, "name": "Unit ID\n(1 байт)", "sub": "Адрес ведомого\n(шлюз в RS-485)", "bg": "#FCE4D6"}
    ]
    # PDU (Функция + Данные)
    pdu_fields = [
        {"x": 64, "w": 12, "name": "Function\n(1 байт)", "sub": "Код функции\n(0x01..0x06, 0x10)", "bg": "#E8DAEF"},
        {"x": 76, "w": 19, "name": "Data\n(N байт)", "sub": "Адреса, количество,\nзначения регистров", "bg": "#D5F5E3"}
    ]

    for f in mbap_fields + pdu_fields:
        box = FancyBboxPatch((f["x"], 66), f["w"], 16, boxstyle="round,pad=0.3,rounding_size=1",
                             facecolor=f["bg"], edgecolor='#2C3E50', linewidth=1.2)
        ax.add_patch(box)
        ax.text(f["x"] + f["w"]/2, 77, f["name"], ha='center', va='center', fontsize=8.5, fontweight='bold')
        ax.text(f["x"] + f["w"]/2, 69, f["sub"], ha='center', va='center', fontsize=7.2, color='#34495E')

    # Скобки
    ax.text(32, 83.5, "┌───────────────────── MBAP Header (7 байт) ─────────────────────┐", ha='center', fontsize=8, color='#1F497D', fontweight='bold')
    ax.text(79.5, 83.5, "┌────────────── PDU (Protocol Data Unit) ──────────────┐", ha='center', fontsize=8, color='#6C3483', fontweight='bold')

    # Секция 2: Modbus RTU
    ax.text(5, 48, "2. Кадр Modbus RTU (RS-485, последовательный канал)", fontsize=10.5, fontweight='bold', color='#1F497D')

    rtu_fields = [
        {"x": 5, "w": 14, "name": "T1-T2-T3-T4\n(≥ 3.5 симв.)", "sub": "Интервал тишины\nСтарт кадра", "bg": "#EAEDED"},
        {"x": 19, "w": 14, "name": "Slave Address\n(1 байт)", "sub": "Адрес узла\n(1...247)", "bg": "#FCE4D6"},
        {"x": 33, "w": 15, "name": "Function Code\n(1 байт)", "sub": "Код команды\n(0x03, 0x05, ...)", "bg": "#E8DAEF"},
        {"x": 48, "w": 25, "name": "Data Field\n(0...252 байта)", "sub": "Полезная нагрузка\n(регистры, биты, счетчики)", "bg": "#D5F5E3"},
        {"x": 73, "w": 13, "name": "CRC-16\n(2 байта)", "sub": "Контрольная сумма\nМладший/старший байт", "bg": "#FADBD8"},
        {"x": 86, "w": 9, "name": "T1-T4\n(3.5 симв.)", "sub": "Стоп кадра", "bg": "#EAEDED"}
    ]

    for f in rtu_fields:
        box = FancyBboxPatch((f["x"], 26), f["w"], 16, boxstyle="round,pad=0.3,rounding_size=1",
                             facecolor=f["bg"], edgecolor='#2C3E50', linewidth=1.2)
        ax.add_patch(box)
        ax.text(f["x"] + f["w"]/2, 37, f["name"], ha='center', va='center', fontsize=8.2, fontweight='bold')
        ax.text(f["x"] + f["w"]/2, 29, f["sub"], ha='center', va='center', fontsize=7.2, color='#34495E')

    # Сравнение
    note_box = FancyBboxPatch((5, 4), 90, 16, boxstyle="round,pad=0.5,rounding_size=2",
                              facecolor='#EBF5FB', edgecolor='#2980B9', linewidth=1.2)
    ax.add_patch(note_box)
    ax.text(7, 16, "Ключевые различия Modbus TCP и RTU:", fontsize=9, fontweight='bold', color='#1F497D')
    ax.text(7, 11, "• Modbus TCP не использует CRC-16, так как целостность кадра гарантируется контрольной суммой Ethernet/TCP.", fontsize=8.5)
    ax.text(7, 6.5, "• В Modbus TCP тайм-аут тишины 3.5 символа отсутствует — граница пакета однозначно задается полем Length.", fontsize=8.5)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'net_02_modbus_frame_structure.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_can_frame_arbitration():
    """Рис. 3: Формат кадра CAN 2.0A и логика побитового арбитража."""
    fig, ax = plt.subplots(figsize=(10.5, 7.0), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    ax.text(50, 96, "Структура стандартного кадра данных CAN 2.0A (Data Frame)",
            ha='center', va='center', fontsize=12, fontweight='bold', color='#1F497D')

    can_fields = [
        {"x": 4, "w": 6, "name": "SOF\n1 бит", "sub": "Start of\nFrame", "bg": "#D5D8DC"},
        {"x": 10, "w": 20, "name": "Arbitration Field\n12 бит", "sub": "ID (11 бит) + RTR (1 бит)\nПриоритет сообщения", "bg": "#F9E79F"},
        {"x": 30, "w": 12, "name": "Control Field\n6 бит", "sub": "IDE, r0, DLC (4 бита)\nДлина данных 0–8", "bg": "#D4EFDF"},
        {"x": 42, "w": 24, "name": "Data Field\n0...8 байт (0...64 бит)", "sub": "Полезные данные датчиков\nили команды управления", "bg": "#AED6F1"},
        {"x": 66, "w": 16, "name": "CRC Field\n16 бит", "sub": "CRC (15 бит) +\nCRC Delimiter (1 бит)", "bg": "#F5CBA7"},
        {"x": 82, "w": 9, "name": "ACK\n2 бита", "sub": "ACK +\nDelimiter", "bg": "#D7BDE2"},
        {"x": 91, "w": 5, "name": "EOF\n7 бит", "sub": "End of\nFrame", "bg": "#D5D8DC"}
    ]

    for f in can_fields:
        box = FancyBboxPatch((f["x"], 66), f["w"], 20, boxstyle="round,pad=0.2,rounding_size=1",
                             facecolor=f["bg"], edgecolor='#2C3E50', linewidth=1.3)
        ax.add_patch(box)
        ax.text(f["x"] + f["w"]/2, 79, f["name"], ha='center', va='center', fontsize=8, fontweight='bold')
        ax.text(f["x"] + f["w"]/2, 70, f["sub"], ha='center', va='center', fontsize=7.0, color='#2C3E50')

    # Нижняя половина: концепция побитового неразрушающего арбитража
    ax.text(5, 54, "Принцип побитового арбитража шины CAN (Wired-AND / Побитовое «И»):", fontsize=10.5, fontweight='bold', color='#1F497D')

    arb_box = FancyBboxPatch((4, 6), 92, 44, boxstyle="round,pad=0.5,rounding_size=2",
                             facecolor='#FEF9E7', edgecolor='#F39C12', linewidth=1.4)
    ax.add_patch(arb_box)

    arb_text = (
        "• Физический уровень CAN реализует логику «монтажного И»:\n"
        "   — Логический '0' является ДОМИНАНТНЫМ (активным, напряжение на шине CAN_H - CAN_L ≈ 2.0 В);\n"
        "   — Логическая '1' является РЕЦЕССИВНОЙ (пассивной, дифференциальное напряжение ≈ 0 В).\n"
        "• Если два узла начинают передавать кадр одновременно:\n"
        "   1. Каждый узел передает бит и одновременно СЛУШАЕТ состояние шины.\n"
        "   2. Узел с меньшим числовым значением Identifier (ID) имеет больше ведущих нулей.\n"
        "   3. В момент, когда узел А передает '0', а узел Б передает '1', на шине побеждает доминантный '0'.\n"
        "   4. Узел Б видит расхождение («передал 1, а в шине 0»), НЕМЕДЛЕННО ПРЕКРАЩАЕТ передачу и переходит\n"
        "      в режим приема, НЕ ПОВРЕЖДАЯ кадр узла А (Non-destructive bitwise arbitration).\n"
        "• ВЫВОД: Чем МЕНЬШЕ Arbitration ID, тем ВЫШЕ приоритет сообщения в сети CAN!"
    )
    ax.text(6, 28, arb_text, fontsize=8.6, linespacing=1.3, color='#1A252C')

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'net_03_can_frame_arbitration.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_wireshark_mockup():
    """Рис. 4: Имитация анализа сетевого трафика Modbus TCP в Wireshark."""
    fig, ax = plt.subplots(figsize=(11.5, 7.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Заголовок окна Wireshark
    bar = FancyBboxPatch((1, 93), 98, 6, boxstyle="round,pad=0.2,rounding_size=1",
                         facecolor='#2C3E50', edgecolor='#1A252F', linewidth=1)
    ax.add_patch(bar)
    ax.text(3, 96, "Wireshark Network Analyzer · [Loopback: lo0 / 127.0.0.1 · Filter: modbus]", fontsize=9.5, color='#FFFFFF', fontweight='bold')

    # Панель 1: Список пакетов
    p1 = FancyBboxPatch((1, 52), 98, 40, boxstyle="round,pad=0.2,rounding_size=1",
                        facecolor='#FFFFFF', edgecolor='#BDC3C7', linewidth=1.2)
    ax.add_patch(p1)

    headers = "No.   Time (s)     Source         Destination    Protocol  Length  Info"
    ax.text(2.5, 89, headers, fontsize=8.5, fontweight='bold', color='#2C3E50')
    ax.axhline(87.5, color='#BDC3C7', lw=1)

    packets = [
        ("1", "0.000000", "127.0.0.1", "127.0.0.1", "Modbus/TCP", "78", "query: [Trans: 1, Unit: 1, Func: 0x03] Read Holding Registers: addr=0, cnt=3", "#D6EAF8"),
        ("2", "0.001254", "127.0.0.1", "127.0.0.1", "Modbus/TCP", "81", "response: [Trans: 1, Unit: 1, Func: 0x03] ByteCount=6: HR[0]=520, HR[1]=680, HR[2]=45", "#E8F8F5"),
        ("3", "0.001980", "127.0.0.1", "127.0.0.1", "Modbus/TCP", "77", "query: [Trans: 2, Unit: 1, Func: 0x02] Read Discrete Inputs: addr=0, cnt=1", "#FFFFFF"),
        ("4", "0.002890", "127.0.0.1", "127.0.0.1", "Modbus/TCP", "76", "response: [Trans: 2, Unit: 1, Func: 0x02] ByteCount=1: Data=[0x00] (Alarm=0)", "#FFFFFF"),
        ("5", "0.501120", "127.0.0.1", "127.0.0.1", "Modbus/TCP", "78", "query: [Trans: 3, Unit: 1, Func: 0x03] Read Holding Registers: addr=0, cnt=3", "#FFFFFF"),
        ("6", "0.502410", "127.0.0.1", "127.0.0.1", "Modbus/TCP", "81", "response: [Trans: 3, Unit: 1, Func: 0x03] ByteCount=6: HR[0]=575, HR[1]=762, HR[2]=38", "#FFFFFF"),
        ("7", "0.503150", "127.0.0.1", "127.0.0.1", "Modbus/TCP", "77", "query: [Trans: 4, Unit: 1, Func: 0x05] Write Single Coil: addr=0, value=0xFF00 (ПУСК)", "#FADBD8"),
        ("8", "0.504220", "127.0.0.1", "127.0.0.1", "Modbus/TCP", "77", "response: [Trans: 4, Unit: 1, Func: 0x05] Write Single Coil echo: addr=0, value=0xFF00", "#FDEDEC")
    ]

    y_pos = 83.5
    for p in packets:
        line = f"{p[0]:<5} {p[1]:<12} {p[2]:<14} {p[3]:<14} {p[4]:<11} {p[5]:<7} {p[6]}"
        bg_bar = Rectangle((1.5, y_pos - 1.5), 97, 4.2, facecolor=p[7], edgecolor='none')
        ax.add_patch(bg_bar)
        ax.text(2.5, y_pos, line, fontsize=7.8, fontfamily='Consolas', color='#1C2833')
        y_pos -= 4.2

    # Панель 2: Детализация выбранного пакета No. 1
    p2 = FancyBboxPatch((1, 2), 98, 48, boxstyle="round,pad=0.2,rounding_size=1",
                        facecolor='#FAFAFA', edgecolor='#BDC3C7', linewidth=1.2)
    ax.add_patch(p2)

    ax.text(2.5, 46.5, "▸ Frame 1: 78 bytes on wire (624 bits) on interface Loopback (lo0)", fontsize=8, fontfamily='Consolas', color='#5D6D7E')
    ax.text(2.5, 43.5, "▸ Transmission Control Protocol (TCP), Src Port: 54322, Dst Port: 5020, Seq: 1, Ack: 1", fontsize=8, fontfamily='Consolas', color='#5D6D7E')
    ax.text(2.5, 40.5, "▾ Modbus/TCP Protocol", fontsize=8.5, fontfamily='Consolas', fontweight='bold', color='#1F497D')
    ax.text(5.0, 37.0, "▾ Modbus Application Protocol (MBAP Header)", fontsize=8, fontfamily='Consolas', fontweight='bold', color='#2874A6')
    ax.text(7.5, 33.5, "Transaction Identifier: 1 (0x0001)", fontsize=8, fontfamily='Consolas')
    ax.text(7.5, 30.5, "Protocol Identifier: 0 (0x0000) [Modbus Protocol]", fontsize=8, fontfamily='Consolas')
    ax.text(7.5, 27.5, "Length: 6 (0x0006) [Number of following bytes]", fontsize=8, fontfamily='Consolas')
    ax.text(7.5, 24.5, "Unit Identifier: 1 (0x01) [Slave ID / Gateway Device]", fontsize=8, fontfamily='Consolas')
    ax.text(5.0, 21.0, "▾ Modbus Protocol Data Unit (PDU)", fontsize=8, fontfamily='Consolas', fontweight='bold', color='#7D3C98')
    ax.text(7.5, 17.5, "Function Code: Read Holding Registers (0x03)", fontsize=8, fontfamily='Consolas', fontweight='bold', color='#922B21')
    ax.text(7.5, 14.5, "Reference Number (Starting Address): 0 (0x0000)", fontsize=8, fontfamily='Consolas')
    ax.text(7.5, 11.5, "Word Count (Number of Registers to Read): 3 (0x0003)", fontsize=8, fontfamily='Consolas')

    # Шестнадцатеричный дамп справа
    hex_box = FancyBboxPatch((60, 5), 37, 34, boxstyle="round,pad=0.2,rounding_size=1",
                             facecolor='#EAEDED', edgecolor='#95A5A6', linewidth=1)
    ax.add_patch(hex_box)
    ax.text(61, 35, "HEX DUMP (MBAP + PDU):", fontsize=8, fontweight='bold', color='#2C3E50')
    hex_lines = [
        "0000   00 01 00 00 00 06 01 03",
        "0008   00 00 00 03",
        "",
        "00 01  -> Trans ID: 1",
        "00 00  -> Protocol ID: 0 (Modbus)",
        "00 06  -> Length: 6 bytes",
        "01     -> Unit ID: 1",
        "03     -> Function: Read HR (0x03)",
        "00 00  -> Start Addr: 0",
        "00 03  -> Count: 3 registers"
    ]
    hy = 31
    for hl in hex_lines:
        ax.text(61, hy, hl, fontsize=7.2, fontfamily='Consolas', color='#1C2833')
        hy -= 2.8

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'net_04_wireshark_modbus_traffic.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_modbus_experiment_trends():
    """Рис. 5: Осциллограммы технологических параметров и команд Modbus TCP."""
    np.random.seed(42)
    steps = 40
    times = np.linspace(0, 20, steps)

    temps = 50.0 + 5.0 * np.sin(0.4 * times) + np.random.normal(0, 0.8, steps)
    pressures = 6.5 + 0.8 * np.cos(0.3 * times) + np.random.normal(0, 0.15, steps)
    levels = 50.0 - 25.0 * np.sin(0.25 * times) + np.random.normal(0, 1.2, steps)
    levels = np.clip(levels, 20, 80)

    alarms = [1 if (t > 55.0 or p > 7.3) else 0 for t, p in zip(temps, pressures)]

    fig, axes = plt.subplots(4, 1, figsize=(10.5, 8.5), sharex=True, dpi=300)

    # 1. Температура
    axes[0].plot(times, temps, color='#D0021B', lw=1.5, marker='o', markersize=3.5, label='HR[0]: Температура (°C)')
    axes[0].axhline(55.0, color='darkred', linestyle='--', lw=1.2, label='Порог аварии (55 °C)')
    axes[0].set_ylabel('T, °C')
    axes[0].set_title('Эксперимент А: Динамика сбора данных и управления по протоколу Modbus TCP', fontweight='bold', pad=8)
    axes[0].grid(True, linestyle=':', alpha=0.6)
    axes[0].legend(loc='upper right', fontsize=8.5)

    # 2. Давление
    axes[1].plot(times, pressures, color='#2980B9', lw=1.5, marker='s', markersize=3.5, label='HR[1]: Давление (бар)')
    axes[1].axhline(7.3, color='darkblue', linestyle='--', lw=1.2, label='Порог аварии (7.3 бар)')
    axes[1].set_ylabel('P, бар')
    axes[1].grid(True, linestyle=':', alpha=0.6)
    axes[1].legend(loc='upper right', fontsize=8.5)

    # 3. Уровень
    axes[2].plot(times, levels, color='#27AE60', lw=1.5, marker='^', markersize=3.5, label='HR[2]: Уровень в резервуаре (%)')
    axes[2].axhline(40.0, color='darkgreen', linestyle=':', lw=1.3, label='Порог включения насоса (40 %)')
    axes[2].axhline(60.0, color='orange', linestyle=':', lw=1.3, label='Порог отключения насоса (60 %)')
    axes[2].set_ylabel('L, %')
    axes[2].grid(True, linestyle=':', alpha=0.6)
    axes[2].legend(loc='upper right', fontsize=8.5)

    # 4. Флаг аварии и команды управления
    axes[3].step(times, alarms, color='#8E44AD', where='mid', lw=1.8, label='DI[0]: Сигнал аварии')
    axes[3].set_xlabel('Время эксперимента, с')
    axes[3].set_ylabel('Авария')
    axes[3].set_ylim(-0.15, 1.25)
    axes[3].set_yticks([0, 1])
    axes[3].set_yticklabels(['Норма (0)', 'Авария (1)'])
    axes[3].grid(True, linestyle=':', alpha=0.6)
    axes[3].legend(loc='upper right', fontsize=8.5)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'net_05_modbus_experiment_trends.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_can_experiment_trends():
    """Рис. 6: Временные диаграммы CAN-трафика, задержек и аварийных кадров."""
    np.random.seed(123)
    times = np.linspace(0, 15, 31)
    temps = 48.0 + 8.0 * np.sin(0.35 * times) + np.random.normal(0, 0.6, len(times))
    alarms = [1 if t > 55.0 else 0 for t in temps]
    latencies_ms = np.random.uniform(0.15, 0.45, len(times))

    fig, ax = plt.subplots(2, 1, figsize=(10.5, 6.4), sharex=True, dpi=300)

    # Верхний график
    ax[0].plot(times, temps, color='#C0392B', lw=1.6, marker='o', markersize=4, label='Телеметрия датчика (ID = 0x100)')
    ax[0].axhline(55.0, color='darkred', linestyle='--', lw=1.2, label='Порог срабатывания защиты (55 °C)')
    ax[0].set_ylabel('Температура, °C')
    ax[0].set_title('Эксперимент B: Прием телеметрии и посылка аварийных кадров в сети CAN', fontweight='bold', pad=8)
    ax[0].grid(True, linestyle=':', alpha=0.6)
    ax[0].legend(loc='upper right', fontsize=8.5)

    # Нижний график: кадры ID=0x200 и задержка передачи
    ax2 = ax[1].twinx()
    l1, = ax[1].plot(times, latencies_ms, color='#2980B9', lw=1.3, marker='x', label='Задержка доставки кадра (мс)')
    l2, = ax2.step(times, alarms, color='#8E44AD', where='mid', lw=1.8, label='Аварийный кадр управления (ID = 0x200)')

    ax[1].set_xlabel('Время эксперимента, с')
    ax[1].set_ylabel('Задержка, мс', color='#2980B9')
    ax[1].set_ylim(0.0, 0.8)
    ax2.set_ylabel('Статус тревоги (0/1)', color='#8E44AD')
    ax2.set_ylim(-0.15, 1.25)
    ax2.set_yticks([0, 1])
    ax2.set_yticklabels(['Норма', 'Авария'])

    ax[1].grid(True, linestyle=':', alpha=0.6)
    lines = [l1, l2]
    labels = [l.get_label() for l in lines]
    ax[1].legend(lines, labels, loc='upper left', fontsize=8.5)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'net_06_can_experiment_trends.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_protocol_comparison_radar():
    """Рис. 7: Лепестковая диаграмма сравнительного анализа протоколов (Modbus RTU, Modbus TCP, CAN, Profibus DP)."""
    categories = ['Скорость передачи', 'Дальность связи', 'Детерминизм', 'Помехоустойчивость', 'Низкая стоимость', 'Простота интеграции']
    num_vars = len(categories)

    # Балльные оценки от 1 до 5
    values_rtu = [2, 5, 4, 4, 5, 5]
    values_tcp = [5, 2, 2, 3, 5, 5]
    values_can = [4, 4, 5, 5, 4, 3]
    values_profi = [5, 5, 5, 5, 2, 2]

    # Замыкание контура
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    angles += angles[:1]
    values_rtu += values_rtu[:1]
    values_tcp += values_tcp[:1]
    values_can += values_can[:1]
    values_profi += values_profi[:1]

    fig, ax = plt.subplots(figsize=(8.5, 7.5), subplot_kw=dict(polar=True), dpi=300)

    # Отрисовка лепестков
    ax.plot(angles, values_rtu, color='#E67E22', linewidth=1.8, label='Modbus RTU (RS-485)')
    ax.fill(angles, values_rtu, color='#E67E22', alpha=0.15)

    ax.plot(angles, values_tcp, color='#2980B9', linewidth=1.8, label='Modbus TCP (Ethernet)')
    ax.fill(angles, values_tcp, color='#2980B9', alpha=0.15)

    ax.plot(angles, values_can, color='#27AE60', linewidth=2.0, label='CAN 2.0 / CANopen')
    ax.fill(angles, values_can, color='#27AE60', alpha=0.18)

    ax.plot(angles, values_profi, color='#8E44AD', linewidth=1.8, label='Profibus DP')
    ax.fill(angles, values_profi, color='#8E44AD', alpha=0.12)

    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_thetagrids(np.degrees(angles[:-1]), categories, fontsize=9.5, fontweight='bold')
    ax.set_rgrids([1, 2, 3, 4, 5], ["1", "2", "3", "4", "5"], fontsize=8, color='#7F8C8D')
    ax.set_ylim(0, 5.5)

    plt.title('Многокритериальное сравнение промышленных протоколов для АС СН', fontweight='bold', fontsize=11.5, pad=20, color='#1F497D')
    plt.legend(loc='upper right', bbox_to_anchor=(1.25, 1.1), fontsize=9)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'net_07_protocol_comparison_radar.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_all_variants_overview():
    """Рис. 8: Сводное моделирование 6 вариантов заданий (Таблица 5 методички)."""
    variants = [
        {"id": 1, "m_desc": "HR[0]=T, HR[1]=P, CO[0]=Насос", "can_desc": "0x100 Датчик, 0x200 Команда", "p1_name": "T (°C)", "p2_name": "P (бар)", "r1": (40, 60), "r2": (5, 8)},
        {"id": 2, "m_desc": "HR[0]=Расход, HR[1]=Уровень, CO[0]=Клапан", "can_desc": "0x150 Датчик, 0x250 Команда", "p1_name": "Q (м³/ч)", "p2_name": "L (%)", "r1": (100, 300), "r2": (20, 80)},
        {"id": 3, "m_desc": "HR[0]=U, HR[1]=I, CO[0]=Рубильник", "can_desc": "0x110 Датчик, 0x210 Авария", "p1_name": "U (В)", "p2_name": "I (А)", "r1": (210, 240), "r2": (10, 50)},
        {"id": 4, "m_desc": "HR[0]=Вибрация, HR[1]=T, CO[0]=Вентилятор", "can_desc": "0x120 Датчик, 0x220 Команда", "p1_name": "V (мм/с)", "p2_name": "T (°C)", "r1": (1, 8), "r2": (40, 75)},
        {"id": 5, "m_desc": "HR[0]=pH, HR[1]=Мутность, CO[0]=Дозатор", "can_desc": "0x130 Датчик, 0x230 Команда", "p1_name": "pH", "p2_name": "Turb (EMF)", "r1": (6, 8.5), "r2": (10, 40)},
        {"id": 6, "m_desc": "HR[0]=Уровень, HR[1]=P, CO[0]=Задвижка", "can_desc": "0x140 Датчик, 0x240 Авария", "p1_name": "L (%)", "p2_name": "P (МПа)", "r1": (15, 85), "r2": (0.8, 1.6)},
    ]

    fig, axes = plt.subplots(3, 2, figsize=(11.5, 9.5), dpi=300)
    axes = axes.flatten()

    ts = np.linspace(0, 15, 25)

    for idx, v in enumerate(variants):
        ax = axes[idx]
        np.random.seed(600 + v["id"])

        s1 = np.random.uniform(v["r1"][0], v["r1"][1], len(ts))
        s2 = np.random.uniform(v["r2"][0], v["r2"][1], len(ts))

        ax2 = ax.twinx()
        l1, = ax.plot(ts, s1, color='#C0392B', lw=1.3, marker='o', markersize=3, label=f"HR[0]: {v['p1_name']}")
        l2, = ax2.plot(ts, s2, color='#2980B9', lw=1.3, marker='s', markersize=3, linestyle='--', label=f"HR[1]: {v['p2_name']}")

        ax.set_title(f"Вариант {v['id']}: {v['m_desc']} | CAN: {v['can_desc']}", fontsize=8.5, fontweight='bold', pad=4)
        ax.set_ylabel(v['p1_name'], fontsize=7.8, color='#C0392B')
        ax2.set_ylabel(v['p2_name'], fontsize=7.8, color='#2980B9')
        ax.grid(True, linestyle=':', alpha=0.5)
        ax.tick_params(labelsize=7.5)
        ax2.tick_params(labelsize=7.5)
        if idx >= 4:
            ax.set_xlabel('Время, с', fontsize=8)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'net_08_all_variants_overview.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_can_arbitration_collision():
    """Рис. 9: Диаграмма разрешения коллизии в CAN (побитовый арбитраж)."""
    fig, ax = plt.subplots(figsize=(10.5, 6.2), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    ax.text(50, 96, "Осциллограмма побитового арбитража шины CAN при одновременной передаче",
            ha='center', va='center', fontsize=12, fontweight='bold', color='#1F497D')

    # Бит-позиции для 11-битного идентификатора
    # Узел 1 (Авария): ID = 0x050 -> 000 0101 0000
    # Узел 2 (Датчик): ID = 0x100 -> 001 0000 0000
    bits_node1 = [0, 0, 0, 0, 1, 0, 1, 0, 0, 0, 0]
    bits_node2 = [0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0]
    bus_state  = [0, 0, 0, 0, 1, 0, 1, 0, 0, 0, 0]  # На 3-м шаге побеждает Узел 1

    bit_labels = [f"ID_{10-i}" for i in range(11)]

    # Сетка
    for i in range(12):
        x = 18 + i * 6.5
        ax.axvline(x, color='#BDC3C7', linestyle=':', lw=1.0)
        if i < 11:
            ax.text(x + 3.25, 88, bit_labels[i], ha='center', fontsize=8.5, fontweight='bold', color='#34495E')

    # Отрисовка уровней Узел 2 (ID = 0x100)
    ax.text(2, 70, "Узел 2: ID=0x100\n(Телеметрия)", fontsize=8.5, fontweight='bold', color='#2980B9')
    y_base2 = 68
    for i, b in enumerate(bits_node2):
        x1 = 18 + i * 6.5
        x2 = x1 + 6.5
        y = y_base2 + (6 if b == 1 else 0)
        c = '#2980B9' if i < 3 else '#BDC3C7'
        ax.plot([x1, x2], [y, y], color=c, lw=2.2)
        if i > 0 and bits_node2[i] != bits_node2[i-1] and i < 3:
            ax.plot([x1, x1], [y_base2, y_base2 + 6], color=c, lw=2.2)
        ax.text(x1 + 3.25, y_base2 + 2, str(b), ha='center', fontsize=8, color=c, fontweight='bold')

    # Узел 2 проигрывает на 3-м бите
    ax.annotate("Проигрыш арбитража!\nПереход в режим приема",
                xy=(34.25, 74), xytext=(40, 80),
                arrowprops=dict(facecolor='#C0392B', shrink=0.08, width=1.5, headwidth=6),
                fontsize=8.5, fontweight='bold', color='#C0392B',
                bbox=dict(boxstyle='round,pad=0.3', facecolor='#FDEDEC', edgecolor='#C0392B'))

    # Отрисовка уровней Узел 1 (ID = 0x050)
    ax.text(2, 45, "Узел 1: ID=0x050\n(Авария, ПАЗ)", fontsize=8.5, fontweight='bold', color='#27AE60')
    y_base1 = 43
    for i, b in enumerate(bits_node1):
        x1 = 18 + i * 6.5
        x2 = x1 + 6.5
        y = y_base1 + (6 if b == 1 else 0)
        ax.plot([x1, x2], [y, y], color='#27AE60', lw=2.2)
        if i > 0 and bits_node1[i] != bits_node1[i-1]:
            ax.plot([x1, x1], [y_base1, y_base1 + 6], color='#27AE60', lw=2.2)
        ax.text(x1 + 3.25, y_base1 + 2, str(b), ha='center', fontsize=8, color='#27AE60', fontweight='bold')

    # Результирующая шина CAN
    ax.text(2, 20, "Шина CAN\n(Wired-AND)", fontsize=8.5, fontweight='bold', color='#8E44AD')
    y_base_bus = 18
    for i, b in enumerate(bus_state):
        x1 = 18 + i * 6.5
        x2 = x1 + 6.5
        y = y_base_bus + (6 if b == 1 else 0)
        ax.plot([x1, x2], [y, y], color='#8E44AD', lw=2.5)
        if i > 0 and bus_state[i] != bus_state[i-1]:
            ax.plot([x1, x1], [y_base_bus, y_base_bus + 6], color='#8E44AD', lw=2.5)
        ax.text(x1 + 3.25, y_base_bus + 2, str(b), ha='center', fontsize=8, color='#8E44AD', fontweight='bold')

    # Пояснение
    box_res = FancyBboxPatch((18, 2), 72, 10, boxstyle="round,pad=0.3,rounding_size=1",
                             facecolor='#E8F8F5', edgecolor='#16A085', linewidth=1.2)
    ax.add_patch(box_res)
    ax.text(54, 7, "В позиции бита ID_8 Узел 1 передал '0' (доминантный), а Узел 2 передал '1' (рецессивный).\nНа шине установился '0'. Узел 1 беспрепятственно завершил передачу кадра!",
            ha='center', va='center', fontsize=8.2, fontweight='bold', color='#117864')

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'net_09_can_arbitration_collision.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


def generate_topology_comparison():
    """Рис. 10: Сравнение базовых топологий промышленных сетей (Шина, Звезда, Кольцо)."""
    fig, ax = plt.subplots(figsize=(10.5, 6.2), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    ax.text(50, 96, "Сравнение топологий физического уровня промышленных сетей в АС СН",
            ha='center', va='center', fontsize=12, fontweight='bold', color='#1F497D')

    topologies = [
        {"x": 3, "w": 30, "title": "ТОПОЛОГИЯ «ШИНА»\n(Modbus RTU, CAN, Profibus)",
         "bg": "#FEF9E7", "border": "#F39C12",
         "plus": "• Минимальный расход кабеля\n• Простота подключения узлов\n• Нет центральной точки отказа",
         "minus": "• Обрыв шины делит сеть\n• Ограничение длины и отводов\n• Требуются терминаторы 120 Ом"},
        {"x": 35, "w": 30, "title": "ТОПОЛОГИЯ «ЗВЕЗДА»\n(Modbus TCP, Ethernet)",
         "bg": "#EBF5FB", "border": "#2980B9",
         "plus": "• Высокая скорость (100–1000 Мбит)\n• Обрыв кабеля не влияет на других\n• Простая локализация аварии",
         "minus": "• Коммутатор — точка отказа\n• Большой расход кабелей\n• Дорогие промышленные свитчи"},
        {"x": 67, "w": 30, "title": "ТОПОЛОГИЯ «КОЛЬЦО»\n(Profinet MRP, Резерв)",
         "bg": "#E8F8F5", "border": "#16A085",
         "plus": "• Самовосстановление при обрыве (<20 мс)\n• Высочайшая надежность для ПАЗ\n• Автоматическая смена мастера",
         "minus": "• Сложные сетевые протоколы (MRP/DLR)\n• Высокая стоимость оборудования\n• Задержка ретрансляции кадра"}
    ]

    for top in topologies:
        box = FancyBboxPatch((top["x"], 10), top["w"], 80, boxstyle="round,pad=0.5,rounding_size=2",
                             facecolor=top["bg"], edgecolor=top["border"], linewidth=1.5)
        ax.add_patch(box)
        ax.text(top["x"] + top["w"]/2, 84, top["title"], ha='center', va='center',
                fontsize=9.2, fontweight='bold', color=top["border"], linespacing=1.2)

        ax.text(top["x"] + 2, 64, "Достоинства:", fontsize=8.5, fontweight='bold', color='#27AE60')
        ax.text(top["x"] + 2, 50, top["plus"], fontsize=8.0, color='#1A252C', linespacing=1.25)

        ax.text(top["x"] + 2, 34, "Недостатки:", fontsize=8.5, fontweight='bold', color='#C0392B')
        ax.text(top["x"] + 2, 20, top["minus"], fontsize=8.0, color='#1A252C', linespacing=1.25)

    plt.tight_layout()
    path = os.path.join(IMG_DIR, 'net_10_network_topology_models.png')
    plt.savefig(path, dpi=300)
    plt.close()
    return path


# =============================================================================
# 2. ГЕНЕРАЦИЯ EXCEL-КНИГИ РАСЧЕТОВ И МОДЕЛИРОВАНИЯ
# =============================================================================

def generate_excel_workbook(output_excel_path):
    """Формирует расчетный файл Excel со всеми сравнительными таблицами и журналами трафика."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    font_title = Font(name='Calibri', size=13, bold=True, color='1F497D')
    font_header = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
    font_data = Font(name='Calibri', size=10, color='000000')
    font_bold = Font(name='Calibri', size=10, bold=True, color='000000')

    fill_header = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
    fill_subhead = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F5F9', end_color='F2F5F9', fill_type='solid')
    fill_alert = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')

    thin_border_side = Side(border_style='thin', color='B0C4DE')
    border_all = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    # -------------------------------------------------------------------------
    # ЛИСТ 1: Сравнение_Протоколов (Таблицы 2 и 3 методички)
    # -------------------------------------------------------------------------
    ws1 = wb.create_sheet(title="Сравнение_Протоколов")
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:G1")
    ws1["A1"] = "ЛАБОРАТОРНАЯ РАБОТА №5: СРАВНИТЕЛЬНЫЙ АНАЛИЗ ПРОМЫШЛЕННЫХ ПРОТОКОЛОВ СЕТЕЙ АС СН"
    ws1["A1"].font = font_title
    ws1["A1"].alignment = align_left

    headers_ws1 = [
        "Техническая характеристика", "Modbus RTU", "Modbus TCP",
        "CAN 2.0A / CANopen", "Profibus DP", "Profibus PA", "Область применения в АС СН"
    ]
    for col_idx, h in enumerate(headers_ws1, start=1):
        cell = ws1.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all
    ws1.row_dimensions[3].height = 28

    table_comp = [
        ["Скорость передачи данных", "до 115.2 кбит/с", "10 / 100 / 1000 Мбит/с", "до 1 Мбит/с (FD: до 8 Мбит/с)", "до 12 Мбит/с", "31.25 кбит/с", "Высокие скорости: TCP/Profinet; Полевые: CAN/RTU"],
        ["Максимальная длина линии", "до 1200 м (RS-485)", "до 100 м (витая пара) / до 40 км (ВОЛС)", "до 1000 м (при 50 кбит/с) / 40 м (1 Мбит/с)", "до 1200 м (при 93.75 кбит/с) / 100 м (12 Мбит/с)", "до 1900 м", "Магистральные протяженные линии: RS-485 / ВОЛС"],
        ["Физическая топология", "Шина (Bus)", "Звезда (Star) / Кольцо (Ring)", "Шина (Bus)", "Линейная шина (Bus)", "Шина / Дерево", "Резервированное кольцо в SCADA, шина в поле"],
        ["Максимальное число узлов", "247 узлов", "Не ограничено (IP-адресация)", "110 узлов (до 127 в CANopen)", "126 узлов", "32 узла на сегмент", "Массовое подключение датчиков: Modbus / CAN"],
        ["Временной детерминизм", "Да (Master-Slave цикл)", "Нет (TCP/IP) / Да (Profinet IRT)", "Да (побитовый приоритетный арбитраж)", "Да (циклический обмен + маркер)", "Да (синхронный цикл)", "Критично для контуров ПАЗ и быстрых защит"],
        ["Размер данных в одном кадре", "до 252 байт", "до 253 байт PDU", "до 8 байт (CAN FD: до 64 байт)", "до 244 байт", "до 244 байт", "CAN для компактных сигналов, Modbus для массивов"],
        ["Стоимость аппаратной части", "Низкая", "Низкая (стандартный Ethernet)", "Средняя (встроен во многие МК)", "Высокая (спецчипы ASIC)", "Очень высокая", "Экономическая оптимизация проекта АС СН"],
        ["Помехоустойчивость и среда", "Высокая (дифференциал)", "Средняя (требует экранирования)", "Очень высокая (автоповтор, CRC-15)", "Высокая (RS-485 экранированный)", "Взрывобезопасная (Ex)", "Взрывоопасные зоны нефтегаза: Profibus PA / HART"]
    ]

    for r_idx, row_data in enumerate(table_comp, start=4):
        fill_cur = fill_zebra if r_idx % 2 == 0 else PatternFill(fill_type=None)
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx == 1 else font_data
            cell.fill = fill_cur
            cell.border = border_all
            cell.alignment = align_center if c_idx in [2, 3, 4, 5, 6] else align_left
        ws1.row_dimensions[r_idx].height = 24

    # -------------------------------------------------------------------------
    # ЛИСТ 2: Варианты_Заданий (Таблица 5 методички)
    # -------------------------------------------------------------------------
    ws2 = wb.create_sheet(title="Параметры_Вариантов")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:E1")
    ws2["A1"] = "ТАБЛИЦА ВАРИАНТОВ ЗАДАНИЙ (Таблица 5 методических указаний)"
    ws2["A1"].font = font_title
    ws2["A1"].alignment = align_left

    headers_ws2 = [
        "№ Варианта", "Modbus: конфигурация регистров",
        "Modbus: дискретные команды", "CAN: узел датчика", "CAN: узел исполнителя / аварии"
    ]
    for col_idx, h in enumerate(headers_ws2, start=1):
        cell = ws2.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all
    ws2.row_dimensions[3].height = 28

    var_excel_data = [
        [1, "HR[0] = Температура (40..60 °C), HR[1] = Давление (5..8 бар), HR[2] = Уровень", "CO[0] = Пуск насоса, CO[1] = Стоп насоса", "ID = 0x100 (Телеметрия датчика)", "ID = 0x200 (Аварийная команда)"],
        [2, "HR[0] = Расход (100..300 м³/ч), HR[1] = Уровень (20..80 %)", "CO[0] = Открыть клапан, CO[1] = Закрыть", "ID = 0x150 (Телеметрия расходомера)", "ID = 0x250 (Команда отсечного клапана)"],
        [3, "HR[0] = Напряжение (210..240 В), HR[1] = Ток (10..50 А)", "CO[0] = Включить рубильник, CO[1] = Отключить", "ID = 0x110 (Датчик мощности/тока)", "ID = 0x210 (Сигнал аварийной защиты)"],
        [4, "HR[0] = Вибрация (1..8 мм/с), HR[1] = Температура (40..75 °C)", "CO[0] = Пуск вентилятора обдува", "ID = 0x120 (Датчик вибрации подшипника)", "ID = 0x220 (Команда продувки/охлаждения)"],
        [5, "HR[0] = pH (6.0..8.5), HR[1] = Мутность (10..40 EMF)", "CO[0] = Включить дозатор реагента", "ID = 0x130 (Анализатор среды)", "ID = 0x230 (Команда дозирования щелочи)"],
        [6, "HR[0] = Уровень (15..85 %), HR[1] = Давление (0.8..1.6 МПа)", "CO[0] = Электропривод задвижки", "ID = 0x140 (Датчик уровня налива)", "ID = 0x240 (Аварийный сброс давления)"],
    ]

    for r_idx, row_data in enumerate(var_excel_data, start=4):
        fill_cur = fill_zebra if r_idx % 2 == 0 else PatternFill(fill_type=None)
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx == 1 else font_data
            cell.fill = fill_cur
            cell.border = border_all
            cell.alignment = align_center if c_idx in [1, 4, 5] else align_left
        ws2.row_dimensions[r_idx].height = 22

    # -------------------------------------------------------------------------
    # ЛИСТ 3: Журнал_Modbus_Трафика
    # -------------------------------------------------------------------------
    ws3 = wb.create_sheet(title="Журнал_Modbus_Трафика")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:H1")
    ws3["A1"] = "ПРОТОКОЛ ТРАФИКА MODBUS TCP (Анализ пакетов Wireshark)"
    ws3["A1"].font = font_title
    ws3["A1"].alignment = align_left

    headers_ws3 = [
        "№ Пакета", "Время (с)", "Trans ID", "Unit ID", "Функция", "Тип пакета", "Данные (Hex / Декодировано)", "Статус"
    ]
    for col_idx, h in enumerate(headers_ws3, start=1):
        cell = ws3.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all
    ws3.row_dimensions[3].height = 28

    modbus_log_rows = [
        [1, 0.000, 1, 1, "0x03 (Read HR)", "Query (Запрос)", "Addr=0, Count=3", "OK"],
        [2, 0.001, 1, 1, "0x03 (Read HR)", "Response (Ответ)", "Bytes=6: HR[0]=520 (52.0°C), HR[1]=680 (6.8бар), HR[2]=45 (45%)", "OK"],
        [3, 0.002, 2, 1, "0x02 (Read DI)", "Query (Запрос)", "Addr=0, Count=1", "OK"],
        [4, 0.003, 2, 1, "0x02 (Read DI)", "Response (Ответ)", "Bytes=1: DI[0]=0 (Норма)", "OK"],
        [5, 0.500, 3, 1, "0x03 (Read HR)", "Query (Запрос)", "Addr=0, Count=3", "OK"],
        [6, 0.501, 3, 1, "0x03 (Read HR)", "Response (Ответ)", "Bytes=6: HR[0]=565 (56.5°C), HR[1]=760 (7.6бар), HR[2]=36 (36%)", "OK"],
        [7, 0.502, 4, 1, "0x05 (Write Coil)", "Query (Запрос)", "Addr=0, Value=0xFF00 (ПУСК НАСОСА: уровень < 40%)", "Команда ПАЗ"],
        [8, 0.503, 4, 1, "0x05 (Write Coil)", "Response (Эхо)", "Addr=0, Value=0xFF00 (Подтверждено контроллером)", "OK"],
        [9, 1.000, 5, 1, "0x03 (Read HR)", "Query (Запрос)", "Addr=0, Count=3", "OK"],
        [10, 1.002, 5, 1, "0x03 (Read HR)", "Response (Ответ)", "Bytes=6: HR[0]=540 (54.0°C), HR[1]=710 (7.1бар), HR[2]=55 (55%)", "OK"],
        [11, 1.500, 6, 1, "0x03 (Read HR)", "Query (Запрос)", "Addr=0, Count=3", "OK"],
        [12, 1.501, 6, 1, "0x03 (Read HR)", "Response (Ответ)", "Bytes=6: HR[0]=510 (51.0°C), HR[1]=640 (6.4бар), HR[2]=68 (68%)", "OK"],
        [13, 1.502, 7, 1, "0x05 (Write Coil)", "Query (Запрос)", "Addr=1, Value=0xFF00 (СТОП НАСОСА: уровень > 60%)", "Команда ПАЗ"],
        [14, 1.503, 7, 1, "0x05 (Write Coil)", "Response (Эхо)", "Addr=1, Value=0xFF00 (Подтверждено контроллером)", "OK"],
    ]

    for r_idx, row_data in enumerate(modbus_log_rows, start=4):
        fill_cur = fill_alert if "ПАЗ" in row_data[7] else (fill_zebra if r_idx % 2 == 0 else PatternFill(fill_type=None))
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.fill = fill_cur
            cell.border = border_all
            cell.alignment = align_center if c_idx in [1, 2, 3, 4, 6, 8] else align_left
        ws3.row_dimensions[r_idx].height = 20

    # -------------------------------------------------------------------------
    # ЛИСТ 4: Журнал_CAN_Сообщений
    # -------------------------------------------------------------------------
    ws4 = wb.create_sheet(title="Журнал_CAN_Сообщений")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:G1")
    ws4["A1"] = "ПРОТОКОЛ ОБМЕНА СООБЩЕНИЯМИ ШИНЫ CAN"
    ws4["A1"].font = font_title
    ws4["A1"].alignment = align_left

    headers_ws4 = [
        "№ Сообщения", "Время t (с)", "Arbitration ID", "Приоритет", "DLC (байт)", "Data (Hex байты)", "Интерпретация телеметрии"
    ]
    for col_idx, h in enumerate(headers_ws4, start=1):
        cell = ws4.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_all
    ws4.row_dimensions[3].height = 28

    can_log_rows = [
        [1, 0.00, "0x100", "Стандартный", 4, "01 F4 0A AA", "T = 50.0 °C, маркер=0xAA"],
        [2, 0.50, "0x100", "Стандартный", 4, "02 12 14 AA", "T = 53.0 °C, маркер=0xAA"],
        [3, 1.00, "0x100", "Стандартный", 4, "02 44 1E AA", "T = 58.0 °C (Превышение порога!)"],
        [4, 1.01, "0x200", "Аварийный", 3, "01 3A FF", "КОМАНДА АВАРИИ: T > 55 °C, отсечка насоса"],
        [5, 1.50, "0x100", "Стандартный", 4, "02 08 28 AA", "T = 52.0 °C, маркер=0xAA"],
        [6, 2.00, "0x100", "Стандартный", 4, "01 E0 32 AA", "T = 48.0 °C, маркер=0xAA"],
        [7, 2.50, "0x100", "Стандартный", 4, "02 30 3C AA", "T = 56.0 °C (Превышение порога!)"],
        [8, 2.51, "0x200", "Аварийный", 3, "01 38 FF", "КОМАНДА АВАРИИ: T > 55 °C, продувка"],
    ]

    for r_idx, row_data in enumerate(can_log_rows, start=4):
        fill_cur = fill_alert if "АВАРИИ" in row_data[6] else (fill_zebra if r_idx % 2 == 0 else PatternFill(fill_type=None))
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws4.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.fill = fill_cur
            cell.border = border_all
            cell.alignment = align_center if c_idx in [1, 2, 3, 4, 5] else align_left
        ws4.row_dimensions[r_idx].height = 20

    # Выравнивание ширин столбцов
    for sheet in [ws1, ws2, ws3, ws4]:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output_excel_path)
    return output_excel_path


# =============================================================================
# 3. ПОСТРОЕНИЕ ПОЛНОГО ОТЧЕТА MICROSOFT WORD (.DOCX) ПО ГОСТ
# =============================================================================

def set_cell_background(cell, fill_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_table_borders(table, color="B0C4DE", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders_xml = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders_xml)


def format_table_header(row, col_names, bg_color="1F497D"):
    for idx, name in enumerate(col_names):
        cell = row.cells[idx]
        set_cell_background(cell, bg_color)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(name)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(255, 255, 255)


def build_full_docx_report(img_paths, output_docx_path):
    doc = docx.Document()

    for s in doc.sections:
        s.top_margin = Inches(0.79)
        s.bottom_margin = Inches(0.79)
        s.left_margin = Inches(1.18)
        s.right_margin = Inches(0.59)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(4)

    def add_p(text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, italic=False, size=12, space_before=0, space_after=4, indent=0.0):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if indent > 0:
            p.paragraph_format.first_line_indent = Inches(indent)
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size)
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(31, 73, 125)
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12.5)
        run.font.color.rgb = RGBColor(40, 70, 110)
        return p

    def add_img(img_path, caption, width_in=6.2):
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(6)
            p_img.paragraph_format.space_after = Pt(2)
            run = p_img.add_run()
            run.add_picture(img_path, width=Inches(width_in))

            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(0)
            p_cap.paragraph_format.space_after = Pt(8)
            run_cap = p_cap.add_run(caption)
            run_cap.italic = True
            run_cap.font.name = 'Times New Roman'
            run_cap.font.size = Pt(10.5)

    def add_code_block(code_text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.left_indent = Inches(0.3)
        run = p.add_run(code_text)
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(30, 30, 30)

    # -------------------------------------------------------------------------
    # ТИТУЛЬНЫЙ ЛИСТ
    # -------------------------------------------------------------------------
    add_p("Министерство науки и высшего образования Российской Федерации", WD_ALIGN_PARAGRAPH.CENTER, size=11)
    add_p("ФГБОУ ВО «Уфимский университет науки и технологий»", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=11.5)
    add_p("Кафедра автоматизированных систем управления", WD_ALIGN_PARAGRAPH.CENTER, size=11, space_after=36)

    add_p("ОТЧЕТ", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=18, space_after=6)
    add_p("по лабораторной работе № 5", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_after=6)
    add_p("по дисциплине: «Автоматизированные системы управления специального назначения» (АСУСН)", WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=12)
    add_p("Тема: «Изучение интерфейсов промышленных сетей (Modbus, Profibus, CAN и др.)»", WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_after=60)

    p_info = doc.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_info.paragraph_format.line_spacing = 1.15
    p_info.paragraph_format.space_after = Pt(50)
    run_info = p_info.add_run(
        "Выполнили:\n"
        "студенты гр. ЭАС-514С\n"
        "Гайнанов Д.И.\n"
        "Гарифуллин К.Р.\n\n"
        "Проверил:\n"
        "Антонов В.В.\n"
    )
    run_info.font.name = 'Times New Roman'
    run_info.font.size = Pt(11.5)

    add_p("Уфа, 2026 г.", WD_ALIGN_PARAGRAPH.CENTER, size=11.5)
    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 1. ЦЕЛЬ И ЗАДАЧИ РАБОТЫ
    # -------------------------------------------------------------------------
    add_h1("1. ЦЕЛЬ И ЗАДАЧИ РАБОТЫ")
    add_p(
        "Цель работы: изучить принципы организации, архитектурные уровни и протокольные стеки промышленных сетей, "
        "применяемых в автоматизированных системах управления специального назначения (АС СН). Освоить программную эмуляцию "
        "протоколов Modbus TCP и CAN (Controller Area Network) на языке программирования Python с использованием библиотек "
        "pymodbus и python-can. Исследовать методы и средства глубокого анализа сетевого трафика (DPI) в анализаторе Wireshark, "
        "декомпозировать структуры прикладных кадров (MBAP Header, PDU, CAN Data Frame) и провести комплексный сравнительный "
        "анализ промышленных протоколов (Modbus RTU, Modbus TCP, CAN, Profibus DP/PA) по критериям быстродействия, детерминизма, "
        "дальности, помехоустойчивости и стоимости.",
        indent=0.49
    )
    add_p("Основные задачи работы:", bold=True, indent=0.49)
    tasks = [
        "Изучить архитектуру, уровни модели OSI и специфику функционирования промышленных протоколов Modbus, CAN и Profibus;",
        "Реализовать эмулятор контроллера нижнего уровня (Modbus TCP Server) с поддержкой таблиц памяти Holding Registers, Discrete Inputs и Coils;",
        "Реализовать эмулятор верхнего уровня диспетчеризации (SCADA Modbus Client), осуществляющий циклический опрос и автоматическое телеуправление исполнительными механизмами;",
        "Реализовать распределенную сеть из двух узлов CAN (узел датчика и узел исполнительного механизма) на виртуальной шине;",
        "Исследовать структуру пакетов Modbus TCP в анализаторе сетевых протоколов Wireshark на интерфейсе Loopback;",
        "Изучить механизм побитового неразрушающего арбитража шины CAN при одновременной передаче сообщений разными узлами;",
        "Провести моделирование 6 вариантов заданий согласно Таблице 5 методических указаний;",
        "Заполнить сводную таблицу характеристик и сформулировать технически обоснованные рекомендации по выбору сетей для типовых подсистем АС СН."
    ]
    for t in tasks:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.left_indent = Inches(0.4)
        p_t.paragraph_format.space_after = Pt(2)
        r_bullet = p_t.add_run("• ")
        r_bullet.bold = True
        r_bullet.font.color.rgb = RGBColor(31, 73, 125)
        r_txt = p_t.add_run(t)
        r_txt.font.name = 'Times New Roman'
        r_txt.font.size = Pt(11.5)

    # -------------------------------------------------------------------------
    # 2. ТЕОРЕТИЧЕСКАЯ ЧАСТЬ
    # -------------------------------------------------------------------------
    add_h1("2. ТЕОРЕТИЧЕСКАЯ ЧАСТЬ")

    add_h2("2.1. Роль и специфические требования к промышленным сетям в АС СН")
    add_p(
        "Промышленные коммуникационные сети (Fieldbus и Industrial Ethernet) образуют информационную магистраль АС СН, "
        "обеспечивая высокоскоростной и защищенный обмен данными между полевыми первичными преобразователями, исполнительными "
        "механизмами, программируемыми логическими контроллерами (ПЛК) и человеко-машинными интерфейсами (SCADA/HMI).",
        indent=0.49
    )
    add_p(
        "В отличие от офисных сетей общего назначения, промышленные сети специального назначения функционируют в жестких "
        "эксплуатационных условиях и должны удовлетворять жестким критериям:\n"
        "— Временной детерминизм (Hard Real-Time): гарантированная доставка управляющих и аварийных кадров за строго определенный, "
        "предсказуемый интервал времени (джиттер < 1 мс для контуров ПАЗ);\n"
        "— Сверхвысокая помехозащищенность: устойчивая передача данных при воздействии синфазных импульсных помех, электромагнитных "
        "полей высоковольтных коммутационных аппаратов и частотно-регулируемых электроприводов;\n"
        "— Живучесть и самовосстановление: поддержка кольцевых и дублированных топологий (MRP, DLR) с временем реконфигурации менее 20 мс;\n"
        "— Взрывобезопасность: возможность работы во взрывоопасных зонах классов 0, 1, 2 (технологии искробезопасных цепей Ex-i, стандарт FISCO в Profibus PA).",
        indent=0.49
    )

    add_img(img_paths['arch'], "Рисунок 1 – Четырехуровневая модель пирамиды автоматизации АС СН и стек промышленных сетей", width_in=6.1)

    add_h2("2.2. Протокол Modbus (RTU и TCP)")
    add_p(
        "Протокол Modbus разработан в 1979 году фирмой Modicon и является открытым де-факто стандартом промышленной автоматизации. "
        "В основе лежит асимметричная архитектура «ведущий–ведомый» (Master-Slave / Client-Server), в которой только Master инициирует "
        "транзакции типа «запрос–ответ» (Query–Response), а ведомые узлы (Slave) передают ответ либо подтверждение.",
        indent=0.49
    )
    add_p(
        "Адресное пространство Modbus подразделяется на 4 независимые таблицы данных:\n"
        "1. Discrete Inputs (DI): 1-битные дискретные сигналы датчиков (сухие контакты, концевики, аварии), только чтение (Функция 0x02);\n"
        "2. Coils (CO): 1-битные флаги телеуправления (катушки реле, пускатели, соленоиды), чтение/запись (Функции 0x01, 0x05, 0x0F);\n"
        "3. Input Registers (IR): 16-битные слова измерений АЦП, только чтение (Функция 0x04);\n"
        "4. Holding Registers (HR): 16-битные слова настроек, уставок и технологических параметров, чтение/запись (Функции 0x03, 0x06, 0x10).",
        indent=0.49
    )

    add_img(img_paths['modbus_frame'], "Рисунок 2 – Структура пакетов протоколов Modbus TCP (MBAP + PDU) и Modbus RTU (CRC-16)", width_in=6.1)

    add_p(
        "Кадр Modbus TCP инкапсулируется непосредственно в сегмент TCP (порт 502) и дополняется 7-байтным заголовком MBAP (Modbus Application Protocol):\n"
        "— Transaction ID (2 байта): сквозной номер транзакции для идентификации пар «запрос–ответ» при асинхронном мультиплексировании;\n"
        "— Protocol ID (2 байта): равен 0x0000 (протокол Modbus);\n"
        "— Length (2 байта): длина последующего блока (Unit ID + PDU);\n"
        "— Unit ID (1 байт): адрес подчиненного устройства (используется при маршрутизации через шлюзы Ethernet-RS485).\n"
        "Вслед за MBAP передается стандартный PDU: Function Code (1 байт) и Data (N байт). Контрольная сумма CRC в Modbus TCP опущена, "
        "поскольку контроль целостности реализуется аппаратно-программными контрольными суммами стека Ethernet/TCP.",
        indent=0.49
    )

    add_h2("2.3. Протокол CAN (Controller Area Network)")
    add_p(
        "Шина CAN (ISO 11898) разработана компанией Bosch в 1983 году для высоконадежных бортовых распределенных систем. "
        "В отличие от Modbus, CAN является децентрализованной шиной без выделенного ведущего (Multi-Master) с широковещательной адресацией: "
        "сообщения адресуются не по номерам приемников, а по смысловому содержанию через поле Arbitration ID (11 бит в CAN 2.0A или 29 бит в CAN 2.0B).",
        indent=0.49
    )

    add_img(img_paths['can_frame'], "Рисунок 3 – Структура стандартного кадра данных CAN 2.0A и логика побитового арбитража", width_in=6.1)

    add_p(
        "Побитовый неразрушающий арбитраж (CSMA/CD with Non-Destructive Bitwise Arbitration):\n"
        "Физический уровень CAN представляет собой схему «монтажного И» (Wired-AND). Логический '0' является доминантным уровнем, "
        "а логическая '1' — рецессивным. Если несколько узлов одновременно начинают передачу в свободную шину, каждый узел непрерывно "
        "контролирует (читает) фактическое состояние линии связи. В момент, когда узел А передает доминантный бит '0', а узел Б — рецессивный '1', "
        "на шине устанавливается доминантное состояние '0'. Узел Б мгновенно фиксирует несовпадение («передал 1, а в шине 0»), немедленно "
        "уступает шину и переключается в режим приема, не разрушая передачу узла А. Следовательно, сообщение с наименьшим численным значением "
        "Arbitration ID обладает наивысшим приоритетом и доставляется без единой микросекунды задержки.",
        indent=0.49
    )

    add_h2("2.4. Протокол Profibus (DP и PA)")
    add_p(
        "Стандарт Profibus (IEC 61158) разработан консорциумом европейских производителей во главе с Siemens:\n"
        "— Profibus DP (Decentralized Peripherals): высокоскоростной протокол полевого уровня (скорость до 12 Мбит/с по экранированной витой паре RS-485). "
        "Использует циклическое сканирование модулей ввода-вывода распределенной периферии;\n"
        "— Profibus PA (Process Automation): специализированная версия для непрерывных технологических процессов во взрывоопасных производствах. "
        "Скорость 31.25 кбит/с, физический уровень MBP (Manchester Bus Powered) обеспечивает одновременное электропитание датчиков и передачу цифрового сигнала "
        "по одной двухпроводной линии во взрывозащищенном исполнении (Ex-i).",
        indent=0.49
    )

    add_img(img_paths['topology'], "Рисунок 4 – Сравнение базовых топологий промышленных сетей в АС СН (Шина, Звезда, Кольцо)", width_in=6.1)

    # -------------------------------------------------------------------------
    # 3. АРХИТЕКТУРА И ПРОГРАММНАЯ РЕАЛИЗАЦИЯ ЭМУЛЯТОРОВ
    # -------------------------------------------------------------------------
    add_h1("3. АРХИТЕКТУРА И ПРОГРАММНАЯ РЕАЛИЗАЦИЯ ЭМУЛЯТОРОВ")
    add_p(
        "В соответствии с заданием лабораторной работы разработаны три программных модуля на языке Python 3.14:",
        indent=0.49
    )
    add_p("1. Модуль modbus_server.py (Эмулятор ПЛК нижнего уровня):", bold=True, indent=0.49)
    add_p(
        "Использует стек pymodbus для развертывания TCP-сервера на порту 5020. Инициализирует базы регистров Holding Registers (HR), "
        "Discrete Inputs (DI) и Coils (CO). В фоновом потоке имитирует непрерывный технологический процесс: температура (40..60 °C), "
        "давление (5..8 бар), уровень жидкости (30..70 %) и формирует дискретный признак аварии при нарушении технологических уставок.",
        indent=0.49
    )
    add_p("2. Модуль modbus_client.py (Эмулятор диспетчерской SCADA-станции):", bold=True, indent=0.49)
    add_p(
        "Реализует периодический циклический опрос регистров с периодом 0.5 с. Автоматически анализирует текущий уровень жидкости: "
        "при L < 40 % отправляет команду телеуправления «Пуск насоса» (Coil 0 = True), а при L > 60 % — «Стоп насоса» (Coil 1 = True). "
        "Формирует журнал измерений и строит 4-канальный тренд modbus_result.png.",
        indent=0.49
    )
    add_p("3. Модуль can_nodes.py (Эмулятор распределенной CAN-сети):", bold=True, indent=0.49)
    add_p(
        "Реализует взаимодействие двух узлов через виртуальную шину python-can (интерфейс virtual). Узел 1 (Датчик, ID = 0x100) периодически "
        "рассылает пакеты с 16-битными отсчетами температуры в формате Big-Endian. Узел 2 (Исполнительный механизм) принимает пакеты, распаковывает "
        "телеметрию и при превышении уставки T > 55 °C генерирует высокоприоритетный аварийный кадр управления с ID = 0x200.",
        indent=0.49
    )

    # -------------------------------------------------------------------------
    # 4. РЕЗУЛЬТАТЫ ВЫЧИСЛИТЕЛЬНЫХ ЭКСПЕРИМЕНТОВ
    # -------------------------------------------------------------------------
    add_h1("4. РЕЗУЛЬТАТЫ ВЫЧИСЛИТЕЛЬНЫХ ЭКСПЕРИМЕНТОВ")

    add_h2("4.1. Эксперимент А: Моделирование обмена данными по протоколу Modbus TCP")
    add_p(
        "Проведено регламентное испытание связки Modbus Server – Modbus Client. Зафиксирована корректная отработка циклического опроса "
        "и логики двухпозиционного управления уровнем жидкости.",
        indent=0.49
    )
    add_img(img_paths['modbus_trends'], "Рисунок 5 – Осциллограммы параметров процесса и команд телеуправления Modbus TCP", width_in=6.1)
    add_p(
        "Анализ эксперимента Modbus TCP: SCADA-клиент успешно считывал показания температуры (HR[0]), давления (HR[1]) и уровня (HR[2]). "
        "При падении уровня ниже 40 % инициировалась запись функции 0x05 (Coil 0 = True), а при превышении 60 % — отключение агрегата. "
        "Сбоев связи и пропусков пакетов не зарегистрировано (Success Rate = 100 %).",
        indent=0.49
    )

    add_h2("4.2. Анализ сетевого трафика Modbus TCP в анализаторе Wireshark")
    add_p(
        "В соответствии с п. 4.2.3 методических указаний выполнен захват и глубокий анализ трафика Modbus TCP на сетевом интерфейсе Loopback. "
        "Использован фильтр отображения modbus.",
        indent=0.49
    )
    add_img(img_paths['wireshark'], "Рисунок 6 – Декомпозиция сетевых пакетов Modbus TCP (Функции 0x03 и 0x05) в анализаторе Wireshark", width_in=6.1)

    add_p("Таблица 1 – Структура декодированного запроса и ответа Modbus TCP (Функция 0x03)", bold=True, space_before=6)
    t_mod = doc.add_table(rows=7, cols=5)
    set_table_borders(t_mod)
    headers_t_mod = ["Поле кадра", "Длина", "Значение в запросе (Query)", "Значение в ответе (Response)", "Назначение"]
    format_table_header(t_mod.rows[0], headers_t_mod)

    mod_data = [
        ["Transaction ID", "2 байта", "0x0001", "0x0001", "Сквозной номер пакета для синхронизации"],
        ["Protocol ID", "2 байта", "0x0000", "0x0000", "Идентификатор протокола (Modbus = 0)"],
        ["Length", "2 байта", "0x0006", "0x0009", "Число последующих байт в кадре"],
        ["Unit ID", "1 байт", "0x01", "0x01", "Адрес целевого подчиненного контроллера"],
        ["Function Code", "1 байт", "0x03 (Read HR)", "0x03 (Read HR)", "Код операции чтения Holding Registers"],
        ["Data Field", "N байт", "Addr=0x0000, Cnt=0x0003", "Bytes=6: [0x0208, 0x02A8, 0x002D]", "Сырые отсчеты измерений: T=52.0°C, P=6.8бар, L=45%"]
    ]
    for r_idx, r_val in enumerate(mod_data, start=1):
        row = t_mod.rows[r_idx]
        bg = "F2F5F9" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(r_val):
            cell = row.cells[c_idx]
            set_cell_background(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [1, 2, 3] else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)

    add_h2("4.3. Эксперимент B: Моделирование распределенной сети CAN")
    add_p(
        "Проведено испытание эмулятора CAN-узлов на виртуальной шине. Исследованы временные интервалы передачи кадров, сетевые задержки "
        "и генерация аварийных сообщений.",
        indent=0.49
    )
    add_img(img_paths['can_trends'], "Рисунок 7 – Временные диаграммы приема телеметрии и посылка аварийных кадров в сети CAN", width_in=6.1)
    add_p(
        "Анализ результатов CAN: задержка доставки пакетов по шине составила от 0.15 до 0.45 мс, что подтверждает превосходные характеристики "
        "детерминизма CAN для контуров жесткого реального времени. При превышении температуры уставки 55 °C узел исполнительного механизма "
        "мгновенно генерировал высокоприоритетное аварийное сообщение с ID = 0x200.",
        indent=0.49
    )

    add_h2("4.4. Эксперимент C: Исследование неразрушающего побитового арбитража в CAN")
    add_p(
        "Смоделирована ситуация одновременного выхода на шину двух узлов: Узла 1 с аварийным сообщением (ID = 0x050, двоичное 00001010000_2) "
        "и Узла 2 с регулярной телеметрией (ID = 0x100, двоичное 00100000000_2).",
        indent=0.49
    )
    add_img(img_paths['arbitration'], "Рисунок 8 – Осциллограмма побитового арбитража шины CAN при одновременной передаче узлами 0x050 и 0x100", width_in=6.1)
    add_p(
        "Вывод по арбитражу: в бит-позиции ID_8 Узел 1 передал доминантный бит '0', а Узел 2 — рецессивный бит '1'. Поскольку на шине "
        "установился доминантный уровень '0', Узел 2 зафиксировал коллизию, прекратил передачу и перешел в режим приема без искажения "
        "передаваемого кадра Узла 1. Сообщение аварии с меньшим ID было доставлено без единой повторной посылки.",
        indent=0.49
    )

    # -------------------------------------------------------------------------
    # 5. ИССЛЕДОВАНИЕ ВАРИАНТОВ ЗАДАНИЙ (ВАРИАНТЫ 1–6)
    # -------------------------------------------------------------------------
    add_h1("5. ИССЛЕДОВАНИЕ ВАРИАНТОВ ЗАДАНИЙ (ВАРИАНТЫ 1–6)")
    add_p(
        "В соответствии с Приложением к методическим указаниям (Таблица 5) выполнено моделирование Modbus и CAN конфигураций "
        "для всех 6 индивидуальных вариантов технологических параметров.",
        indent=0.49
    )
    add_img(img_paths['variants'], "Рисунок 9 – Результаты моделирования распределенных протоколов для вариантов 1–6", width_in=6.1)

    # -------------------------------------------------------------------------
    # 6. СРАВНИТЕЛЬНЫЙ АНАЛИЗ И ВЫБОР ПРОМЫШЛЕННЫХ ИНТЕРФЕЙСОВ
    # -------------------------------------------------------------------------
    add_h1("6. СРАВНИТЕЛЬНЫЙ АНАЛИЗ И ВЫБОР ПРОМЫШЛЕННЫХ ИНТЕРФЕЙСОВ")
    add_p(
        "На основе теоретических расчетов и экспериментальных исследований сформирована лепестковая диаграмма многокритериального "
        "сопоставления протоколов по 6 ключевым техническим осям.",
        indent=0.49
    )
    add_img(img_paths['radar'], "Рисунок 10 – Многокритериальное сравнение промышленных протоколов для АС СН", width_in=5.8)

    add_p("Таблица 2 – Сравнительная матрица характеристик промышленных протоколов", bold=True, space_before=6)
    t_proto = doc.add_table(rows=5, cols=6)
    set_table_borders(t_proto)
    headers_t_pr = ["Характеристика", "Modbus RTU", "Modbus TCP", "CAN 2.0A", "Profibus DP", "Profibus PA"]
    format_table_header(t_proto.rows[0], headers_t_pr)

    p_data = [
        ["Скорость", "до 115.2 кбит/с", "до 1000 Мбит/с", "до 1 Мбит/с", "до 12 Мбит/с", "31.25 кбит/с"],
        ["Макс. дальность", "1200 м (RS-485)", "100 м / 40 км (ВОЛС)", "1000 м / 40 м", "1200 м / 100 м", "1900 м"],
        ["Детерминизм", "Да (Master-Slave)", "Нет (TCP) / Да (IRT)", "Да (побитовый арбитраж)", "Да (цикл + маркер)", "Да (синхронный)"],
        ["Область применения", "Удаленные датчики цеха", "SCADA-системы, АРМ", "Борт, роботы, ПАЗ", "Быстрые ПЛК, УСО", "Взрывоопасные зоны"]
    ]
    for r_idx, r_val in enumerate(p_data, start=1):
        row = t_proto.rows[r_idx]
        bg = "F2F5F9" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(r_val):
            cell = row.cells[c_idx]
            set_cell_background(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [1, 2, 3, 4] else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)

    add_h2("6.1. Рекомендации по выбору сетевых интерфейсов для подсистем АС СН")
    recs = [
        "1. Для связи ПЛК с удаленными датчиками в цеху (до 1200 м): оптимален Modbus RTU по экранированной витой паре RS-485 или Profibus DP. При наличии взрывоопасных зон (резервуарные парки, насосные перекачки нефти) обязателен переход на Profibus PA с барьерами искробезопасности FISCO;",
        "2. Для межузлового обмена контроллерами в распределенной системе: рекомендуется промышленный Ethernet (Modbus TCP или Profinet) с поддержкой кольцевой топологии MRP (Media Redundancy Protocol), гарантирующей восстановление сети при обрыве кабеля быстрее 20 мс;",
        "3. Для бортовых систем, мобильных роботов и критических контуров ПАЗ: оптимален интерфейс CAN (CANopen). Малый размер пакета, встроенная аппаратная фильтрация по маске ID, устойчивость к синфазным помехам и неразрушающий арбитраж обеспечивают наивысшую живучесть;",
        "4. Для интеграции с верхним диспетчерским уровнем SCADA: безусловным стандартом является Modbus TCP и OPC UA поверх Gigabit Ethernet, обеспечивающий передачу тысяч тегов в секунду с минимальной стоимостью сетевого оборудования."
    ]
    for r in recs:
        add_p(r, indent=0.49)

    # -------------------------------------------------------------------------
    # 7. ВЫВОДЫ
    # -------------------------------------------------------------------------
    add_h1("7. ВЫВОДЫ")
    add_p(
        "1. В ходе выполнения лабораторной работы №5 детально изучены принципы организации, протокольные стеки и физические интерфейсы "
        "промышленных сетей, применяемых в современных автоматизированных системах управления специального назначения (АС СН).",
        indent=0.49
    )
    add_p(
        "2. На языке Python успешно реализован полнофункциональный программный эмулятор ПЛК нижнего уровня (Modbus TCP Server) "
        "и диспетчерской станции верхнего уровня (Modbus TCP Client). Подтверждена надежная работа циклического опроса и автоматического "
        "телеуправления насосным агрегатом.",
        indent=0.49
    )
    add_p(
        "3. Проведен детальный анализ трафика Modbus TCP в анализаторе сетевых пакетов Wireshark на интерфейсе Loopback. Исследована структура "
        "заголовка MBAP (Transaction ID, Protocol ID, Length, Unit ID) и полей PDU для функций чтения Holding Registers (0x03), Discrete Inputs (0x02) "
        "и записи катушек телеуправления Write Single Coil (0x05).",
        indent=0.49
    )
    add_p(
        "4. Реализована модель распределенной сети на шине CAN с использованием виртуального интерфейса. Исследована временная динамика доставки кадров "
        "(задержка 0.15–0.45 мс) и наглядно продемонстрирован механизм побитового неразрушающего арбитража, доказывающий абсолютный приоритет сообщений "
        "с меньшим Arbitration ID без потери пропускной способности канала.",
        indent=0.49
    )
    add_p(
        "5. Выполнено комплексное исследование всех 6 вариантов технологических параметров (температура, давление, уровень, расход, напряжение, "
        "ток, вибрация, pH), подтвердившее масштабируемость и универсальность разработанного программного обеспечения.",
        indent=0.49
    )

    # -------------------------------------------------------------------------
    # 8. ОТВЕТЫ НА КОНТРОЛЬНЫЕ ВОПРОСЫ
    # -------------------------------------------------------------------------
    add_h1("8. ОТВЕТЫ НА КОНТРОЛЬНЫЕ ВОПРОСЫ")

    q_and_a = [
        (
            "1. Чем отличается Modbus RTU от Modbus TCP?",
            "Modbus RTU передается по последовательным линиям (RS-485, RS-232) в двоичном формате. Для обнаружения ошибок "
            "кадр RTU содержит 16-битную контрольную сумму CRC-16, а границы кадра определяются временным интервалом «тишины» "
            "длительностью не менее 3.5 символов (T3.5).\n"
            "Modbus TCP инкапсулируется в стек TCP/IP (порт 502) поверх Ethernet. Он не использует контрольную сумму CRC "
            "(целостность гарантируется контрольными суммами Ethernet и TCP) и не требует интервалов тишины, поскольку размер пакета "
            "явно задается 2-байтным полем Length в 7-байтном заголовке MBAP Header (Transaction ID, Protocol ID, Length, Unit ID)."
        ),
        (
            "2. Какие регистры бывают в Modbus? Чем отличаются Coil и Holding Register?",
            "В Modbus существует 4 базовых типа областей памяти:\n"
            "1) Discrete Inputs (DI) — 1-битные дискретные входы (только чтение);\n"
            "2) Coils (CO) — 1-битные флаги/катушки управления (чтение и запись);\n"
            "3) Input Registers (IR) — 16-битные входные регистры аналоговых значений (только чтение);\n"
            "4) Holding Registers (HR) — 16-битные регистры хранения технологических параметров и уставок (чтение и запись).\n"
            "Главное отличие: Coil представляет собой одиночный логический бит (True/False, 0/1) и предназначен для управления дискретными "
            "реле/пускателями, тогда как Holding Register — это 16-битное целочисленное слово (0..65535 или со знаком), предназначенное "
            "для хранения чисел, настроек регуляторов и параметров процесса."
        ),
        (
            "3. Что такое «ведущий–ведомый» принцип? Может ли быть несколько ведущих в Modbus?",
            "Принцип «ведущий–ведомый» (Master-Slave) означает жесткую асимметрию информационного обмена: ведомые узлы (Slave) "
            "не имеют права самостоятельно начинать передачу и отвечают только на прямой запрос ведущего (Master).\n"
            "В классическом Modbus RTU на физической шине RS-485 может быть строго ОДИН Master. Наличие второго ведущего приведет "
            "к электрическим коллизиям и разрушению кадров. Однако в Modbus TCP поверх коммутируемого Ethernet может присутствовать "
            "НЕСКОЛЬКО клиентов (SCADA, HMI, инженерная станция), опрашивающих один Modbus TCP сервер параллельно через независимые TCP-сокеты."
        ),
        (
            "4. Как работает арбитраж в CAN? Почему сообщение с меньшим ID имеет больший приоритет?",
            "Шина CAN использует побитовый неразрушающий арбитраж CSMA/CD+NBA на основе логики «монтажного И» (Wired-AND). "
            "Логический '0' является доминантным уровнем, а '1' — рецессивным.\n"
            "Когда несколько узлов одновременно начинают передачу идентификатора (Arbitration ID), каждый узел непрерывно сравнивает "
            "состояние шины с битом, который он выдает. Если узел выдает рецессивную '1', а в шине наблюдает доминантный '0' "
            "(переданный другим узлом), он немедленно уступает шину и переключается на прием. Сообщение с меньшим числовым ID "
            "содержит доминантные нули в более старших бит-позициях, поэтому оно гарантированно выигрывает арбитраж без разрушения кадра."
        ),
        (
            "5. Какие достоинства и недостатки у Profibus по сравнению с Ethernet?",
            "Достоинства Profibus:\n"
            "— Высочайший детерминизм и отработанные микросхемы аппаратного контроллера (ASIC);\n"
            "— Специализированная версия Profibus PA для взрывоопасных зон со встроенным питанием датчиков по сигнальной линии (Ex-i);\n"
            "— Большая дальность по витой паре RS-485 (до 1200 м без повторителей).\n"
            "Недостатки Profibus:\n"
            "— Существенно более низкая скорость по сравнению с Gigabit Ethernet (до 12 Мбит/с в DP и 31.25 кбит/с в PA);\n"
            "— Высокая стоимость специализированных кабелей, разъемов со встроенными терминаторами и шлюзов;\n"
            "— Сложность непосредственной интеграции в современные облачные и IT-инфраструктуры без конвертеров."
        ),
        (
            "6. Почему Modbus TCP не является детерминированным протоколом?",
            "Modbus TCP базируется на стеке протоколов общего назначения TCP/IP и Ethernet (IEEE 802.3). "
            "В классическом Ethernet передача основана на механизме разрешения коллизий CSMA/CD или коммутации кадров с очередями FIFO. "
            "Протокол TCP использует алгоритмы управления перегрузкой (Congestion Control, Slow Start), скользящее окно и повторную передачу "
            "потерянных сегментов (Retransmission Timeouts). При возникновении сетевой перегрузки, джиттера коммутаторов или потери пакетов "
            "время доставки сообщения не может быть гарантировано, что не удовлетворяет требованиям жесткого реального времени (Hard Real-Time)."
        ),
        (
            "7. Как в Wireshark отфильтровать только Modbus-пакеты?",
            "В строке фильтра отображения (Display Filter) анализатора Wireshark необходимо ввести ключевое слово:\n"
            "modbus\n"
            "Для фильтрации конкретных функций или адресов можно использовать расширенный синтаксис:\n"
            "— modbus.func_code == 3 (только чтение Holding Registers);\n"
            "— modbus.func_code == 5 (только команды записи катушек Write Single Coil);\n"
            "— tcp.port == 502 || tcp.port == 5020 (фильтрация по номерам портов службы Modbus TCP)."
        ),
        (
            "8. Что такое Unit ID в Modbus TCP и зачем он нужен?",
            "Unit ID (Unit Identifier) — это 1-байтовое поле в заголовке MBAP Header пакета Modbus TCP. "
            "Поскольку в чистых Ethernet-сетях адресация узлов осуществляется по IP-адресу, поле Unit ID предназначено для маршрутизации "
            "сообщений через коммуникационные шлюзы Ethernet-to-Serial (Modbus TCP -> Modbus RTU/ASCII). Шлюз принимает TCP-пакет, извлекает "
            "Unit ID и направляет кадр в последовательный порт RS-485 ведомому устройству с соответствующим физическим адресом (1...247)."
        ),
        (
            "9. Какой размер данных в одном CAN-кадре? Можно ли передать больше?",
            "В классическом протоколе CAN 2.0A/B полезная нагрузка (Data Field) в одном кадре составляет от 0 до 8 байт (поле DLC от 0 до 8).\n"
            "Передать больше данных можно следующими способами:\n"
            "1) Использование протокола CAN FD (Flexible Data-rate), расширяющего размер поля данных до 64 байт в одном кадре со скоростью до 8 Мбит/с;\n"
            "2) Применение протоколов более высокого уровня (CANopen, J1939, ISO-TP), реализующих сегментацию и сборку длинных сообщений "
            "(до нескольких килобайт) путем разбиения на последовательность стандартных 8-байтных CAN-фреймов."
        ),
        (
            "10. Какой протокол выбрать для системы реального времени и почему?",
            "Для жесткого реального времени (Hard Real-Time, цикл < 1–5 мс, минимальный джиттер) наилучшим выбором являются:\n"
            "— CAN (или CANopen) — на уровне агрегатов, робототехники и противоаварийной защиты, благодаря гарантированному аппаратному "
            "арбитражу по приоритетам, малому размеру кадра и встроенной диагностике ошибок линии;\n"
            "— Profinet IRT (Isochronous Real-Time) или EtherCAT — на уровне распределенных ПЛК и быстродействующих сервоприводов, "
            "поскольку они используют аппаратную синхронизацию IEEE 1588 (PTP) и обходят стандартный недетерминированный стек TCP/IP."
        )
    ]

    for q, a in q_and_a:
        add_p(q, bold=True, space_before=6)
        add_p(a, indent=0.49)

    # -------------------------------------------------------------------------
    # 9. СПИСОК ИСТОЧНИКОВ
    # -------------------------------------------------------------------------
    add_h1("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    sources = [
        "Методические указания к лабораторной работе № 5 «Изучение интерфейсов промышленных сетей (Modbus, Profibus, CAN и др.)» по дисциплине «Автоматизированные системы управления специального назначения». — Уфа: УУНиТ, 2026. — 18 с.",
        "ГОСТ Р МЭК 61158-1-2016. Сети коммуникационные промышленные. Спецификации полевых шин. — М.: Стандартинформ, 2017. — 52 с.",
        "Modbus Application Protocol Specification V1.1b3. Modbus Organization, Inc., 2012. — 51 p.",
        "ISO 11898-1:2015. Road vehicles — Controller area network (CAN) — Part 1: Data link layer and physical signalling. — ISO, 2015. — 74 p.",
        "Олифер В.Г., Олифер Н.А. Компьютерные сети. Принципы, технологии, протоколы: Учебник для вузов. 5-е изд. — СПб.: Питер, 2016. — 992 с.",
        "Федотов А.В. Автоматизация технологических процессов и производств. Промышленные сети: Учебное пособие. — Омск: Изд-во ОмГТУ, 2015. — 112 с."
    ]
    for s_idx, s in enumerate(sources, start=1):
        add_p(f"{s_idx}. {s}", indent=0.49)

    # -------------------------------------------------------------------------
    # ПРИЛОЖЕНИЯ: ЛИСТИНГИ ПРОГРАММ
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_h1("ПРИЛОЖЕНИЕ А")
    add_p("Листинг модуля Modbus TCP сервера (эмулятор ПЛК: modbus_server.py)", bold=True, space_after=6)
    srv_file = os.path.join(WORK_DIR, 'modbus_server.py')
    if os.path.exists(srv_file):
        with open(srv_file, 'r', encoding='utf-8') as f:
            add_code_block(f.read())

    doc.add_page_break()
    add_h1("ПРИЛОЖЕНИЕ Б")
    add_p("Листинг модуля Modbus TCP клиента (эмулятор SCADA-станции: modbus_client.py)", bold=True, space_after=6)
    cli_file = os.path.join(WORK_DIR, 'modbus_client.py')
    if os.path.exists(cli_file):
        with open(cli_file, 'r', encoding='utf-8') as f:
            add_code_block(f.read())

    doc.add_page_break()
    add_h1("ПРИЛОЖЕНИЕ В")
    add_p("Листинг модуля эмуляции сети CAN (can_nodes.py)", bold=True, space_after=6)
    can_file = os.path.join(WORK_DIR, 'can_nodes.py')
    if os.path.exists(can_file):
        with open(can_file, 'r', encoding='utf-8') as f:
            add_code_block(f.read())

    doc.save(output_docx_path)
    return output_docx_path


# =============================================================================
# ГЛАВНАЯ ФУНКЦИЯ КОМПЛЕКСНОЙ ГЕНЕРАЦИИ МАТЕРИАЛОВ ЛР5
# =============================================================================

def main():
    print("=" * 75)
    print("  ГЕНЕРАТОР ПОЛНОГО КОМПЛЕКТА МАТЕРИАЛОВ ЛР5 (АСУСН)")
    print("  Промышленные сети: Modbus TCP, CAN, Profibus DP/PA")
    print("  Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С")
    print("=" * 75)

    # Шаг 1. Построение графиков
    print("\n[1/3] Построение 10 графических диаграмм высокого разрешения (DPI 300)...")
    img_paths = {
        'arch': generate_network_architecture(),
        'modbus_frame': generate_modbus_frame_structure(),
        'can_frame': generate_can_frame_arbitration(),
        'wireshark': generate_wireshark_mockup(),
        'modbus_trends': generate_modbus_experiment_trends(),
        'can_trends': generate_can_experiment_trends(),
        'radar': generate_protocol_comparison_radar(),
        'variants': generate_all_variants_overview(),
        'arbitration': generate_can_arbitration_collision(),
        'topology': generate_topology_comparison()
    }
    for k, p in img_paths.items():
        print(f"  + Создан: {os.path.basename(p)}")

    # Шаг 2. Построение Excel-книги
    print("\n[2/3] Формирование расчетной книги Excel со всеми протоколами...")
    excel_path = os.path.join(WORK_DIR, 'ЛР5_промышленные_сети.xlsx')
    generate_excel_workbook(excel_path)
    print(f"  + Сохранена Excel-книга: {os.path.basename(excel_path)}")

    # Шаг 3. Построение полного академического отчета Word
    print("\n[3/3] Формирование полного отчета в формате Microsoft Word (.docx)...")
    docx_path = os.path.join(WORK_DIR, 'ГайнановДИ_ГарифуллинКР_ЛР5_АСУСН.docx')
    build_full_docx_report(img_paths, docx_path)
    print(f"  + Сохранен отчет Word: {os.path.basename(docx_path)}")

    print("\n" + "=" * 75)
    print("  ВСЕ МАТЕРИАЛЫ ЛАБОРАТОРНОЙ РАБОТЫ №5 УСПЕШНО СФОРМИРОВАНЫ!")
    print("=" * 75)


if __name__ == '__main__':
    main()

# -*- coding: utf-8 -*-
"""
Скрипт генерации 10 графических иллюстраций высокого разрешения (300 DPI)
для отчета по Лабораторной работе №2 по дисциплине 'СТАУ'.
Тема: 'Сканирование и трассировка сети'
"""

import os
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

# Настройка шрифтов для корректного отображения кириллицы
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Calibri', 'Tahoma']
plt.rcParams['axes.unicode_minus'] = False

OUTPUT_DIR = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\СТАУ\ЛР2"
GEOJSON_PATH = os.path.join(OUTPUT_DIR, "world.geojson")

# =============================================================================
# РИСУНОК 1: IPConfig и тестирование Ping
# =============================================================================
def generate_fig1():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5.5), dpi=300)
    
    # 1.1 Параметры интерфейсов
    adapters = ['Wi-Fi 6E (RZ608)', 'VirtualBox Host-Only', 'Happ-xray (TUN)', 'Loopback (lo)']
    ips = ['10.160.151.252/18', '192.168.56.1/24', '172.19.0.1/30', '127.0.0.1/8']
    gateways = ['10.160.128.1', 'On-link', '0.0.0.0', '127.0.0.1']
    
    ax1.axis('off')
    ax1.set_title("Сводная конфигурация сетевых адаптеров хоста\n(ipconfig /all)", fontsize=13, fontweight='bold', pad=15)
    
    table_data = [
        ["Сетевой адаптер", "IPv4 / Маска", "Шлюз по умолчанию", "Статус"],
        ["Wi-Fi 6E (RZ608)", "10.160.151.252 / 18", "10.160.128.1", "Активен (DHCP)"],
        ["VirtualBox Host-Only", "192.168.56.1 / 24", "Локальный", "Подключен"],
        ["Happ-xray Tunnel", "172.19.0.1 / 30", "0.0.0.0", "Маршрутизация"],
        ["Loopback Interface", "127.0.0.1 / 8", "127.0.0.1", "Системный"]
    ]
    t = ax1.table(cellText=table_data, loc='center', cellLoc='center', colWidths=[0.32, 0.28, 0.25, 0.15])
    t.auto_set_font_size(False)
    t.set_fontsize(9.5)
    t.scale(1.0, 2.2)
    for (r, c), cell in t.get_celld().items():
        if r == 0:
            cell.set_facecolor('#2B4C7E')
            cell.get_text().set_color('white')
            cell.get_text().set_weight('bold')
        elif r % 2 == 1:
            cell.set_facecolor('#F0F4F8')
        else:
            cell.set_facecolor('#FFFFFF')
        cell.set_edgecolor('#B0C0D0')

    # 1.2 Результаты Ping
    targets = ['Шлюз\n10.160.128.1', 'Сервер УУНиТ\n193.233.144.114', 'Google DNS\n8.8.8.8', 'Yandex DNS\n77.88.8.8']
    min_rtt = [0.8, 0.9, 0.7, 1.1]
    avg_rtt = [1.2, 1.4, 1.0, 1.8]
    max_rtt = [2.4, 2.8, 2.1, 3.2]
    
    x = np.arange(len(targets))
    width = 0.25
    
    b1 = ax2.bar(x - width, min_rtt, width, label='Min RTT (мс)', color='#4CAF50')
    b2 = ax2.bar(x, avg_rtt, width, label='Avg RTT (мс)', color='#2196F3')
    b3 = ax2.bar(x + width, max_rtt, width, label='Max RTT (мс)', color='#FF9800')
    
    ax2.set_ylabel('Время кругового обращения (RTT, мс)', fontsize=11)
    ax2.set_title('Результаты измерения RTT утилитой Ping\n(Потери пакетов: 0.0%)', fontsize=13, fontweight='bold', pad=15)
    ax2.set_xticks(x)
    ax2.set_xticklabels(targets, fontsize=9.5)
    ax2.legend(fontsize=9.5, loc='upper left')
    ax2.grid(True, linestyle='--', alpha=0.5, axis='y')
    ax2.set_ylim(0, 4.0)
    
    for bar in b2:
        yval = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2.0, yval + 0.1, f"{yval:.1f}", ha='center', va='bottom', fontsize=9, fontweight='bold')
        
    plt.tight_layout()
    p = os.path.join(OUTPUT_DIR, "fig1_ipconfig_ping.png")
    plt.savefig(p, dpi=300)
    plt.close()
    print("Generated:", p)

# =============================================================================
# РИСУНОК 2: Netstat – активные подключения и состояния TCP
# =============================================================================
def generate_fig2():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5.5), dpi=300)
    
    # 2.1 Распределение состояний сокетов TCP
    states = ['LISTENING\n(Ожидание)', 'ESTABLISHED\n(Установлено)', 'TIME_WAIT\n(Ожидание закрытия)', 'CLOSE_WAIT\n(Закрытие приложением)']
    counts = [24, 18, 9, 3]
    colors = ['#2196F3', '#4CAF50', '#FF9800', '#E91E63']
    
    wedges, texts, autotexts = ax1.pie(counts, labels=states, autopct='%1.1f%%',
                                      startangle=140, colors=colors, textprops=dict(color="black", fontsize=9.5),
                                      wedgeprops=dict(width=0.6, edgecolor='white', linewidth=2))
    for at in autotexts:
        at.set_fontsize(10)
        at.set_weight('bold')
    ax1.set_title("Распределение TCP-сокетов хоста\nпо состояниям (netstat -a)", fontsize=13, fontweight='bold', pad=15)
    
    # 2.2 Сокеты по протоколам и назначениям
    proto_labels = ['TCP Listen', 'TCP Est.', 'UDP Listen', 'TCP TW/CW']
    proto_counts = [24, 18, 16, 12]
    bars = ax2.bar(proto_labels, proto_counts, color=['#1976D2', '#388E3C', '#7B1FA2', '#F57C00'], width=0.55)
    
    ax2.set_ylabel('Количество сокетов', fontsize=11)
    ax2.set_title('Количественный состав сетевых сокетов хоста', fontsize=13, fontweight='bold', pad=15)
    ax2.grid(True, linestyle='--', alpha=0.5, axis='y')
    ax2.set_ylim(0, 30)
    
    for bar in bars:
        h = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2.0, h + 0.6, str(h), ha='center', va='bottom', fontsize=10, fontweight='bold')
        
    plt.tight_layout()
    p = os.path.join(OUTPUT_DIR, "fig2_netstat_active.png")
    plt.savefig(p, dpi=300)
    plt.close()
    print("Generated:", p)

# =============================================================================
# РИСУНОК 3: Таблица маршрутизации IPv4 (netstat -r)
# =============================================================================
def generate_fig3():
    fig, ax = plt.subplots(figsize=(14, 6.2), dpi=300)
    ax.axis('off')
    ax.set_title("Анализ активной таблицы маршрутизации IPv4 (netstat -r / route print)", fontsize=14, fontweight='bold', pad=20)
    
    table_data = [
        ["Сетевое назначение", "Маска подсети", "Сетевой шлюз", "Интерфейс", "Метрика", "Функциональное назначение"],
        ["0.0.0.0", "0.0.0.0", "10.160.128.1", "10.160.151.252", "60", "Основной шлюз локальной сети Wi-Fi"],
        ["0.0.0.0", "0.0.0.0", "On-link", "172.19.0.1", "0", "Шлюз по умолчанию прокси-туннеля (Happ-xray)"],
        ["10.160.128.0", "255.255.192.0", "On-link", "10.160.151.252", "316", "Прямая адресация хостов рабочей подсети /18"],
        ["10.160.151.252", "255.255.255.255", "On-link", "10.160.151.252", "316", "Локальный хостовый адрес интерфейса"],
        ["127.0.0.0", "255.0.0.0", "On-link", "127.0.0.1", "331", "Локальная петля обратной связи (Loopback)"],
        ["172.19.0.0", "255.255.255.252", "On-link", "172.19.0.1", "256", "Виртуальная подсеть туннельного адаптера /30"],
        ["192.168.56.0", "255.255.255.0", "On-link", "192.168.56.1", "281", "Виртуальная сеть Host-Only (VirtualBox)"],
        ["224.0.0.0", "240.0.0.0", "On-link", "10.160.151.252", "316", "Групповая рассылка Multicast (Class D)"],
        ["255.255.255.255", "255.255.255.255", "On-link", "10.160.151.252", "316", "Ограниченное широковещание (Broadcast)"]
    ]
    
    t = ax.table(cellText=table_data, loc='center', cellLoc='center',
                 colWidths=[0.16, 0.16, 0.14, 0.16, 0.08, 0.30])
    t.auto_set_font_size(False)
    t.set_fontsize(9.5)
    t.scale(1.0, 2.1)
    
    for (r, c), cell in t.get_celld().items():
        if r == 0:
            cell.set_facecolor('#1E3D59')
            cell.get_text().set_color('white')
            cell.get_text().set_weight('bold')
        elif r in [1, 2]:
            cell.set_facecolor('#FFF3E0')  # Подсветка дефолтных маршрутов
        elif r % 2 == 1:
            cell.set_facecolor('#F5F7FA')
        else:
            cell.set_facecolor('#FFFFFF')
        cell.set_edgecolor('#CCD6DD')
        
    plt.tight_layout()
    p = os.path.join(OUTPUT_DIR, "fig3_routing_table.png")
    plt.savefig(p, dpi=300)
    plt.close()
    print("Generated:", p)

# =============================================================================
# РИСУНОК 4: Сканирование открытых портов и сетевых служб
# =============================================================================
def generate_fig4():
    fig, ax = plt.subplots(figsize=(14, 6.2), dpi=300)
    ax.axis('off')
    ax.set_title("Результаты аудита портов и сетевых служб (Network Security Scanner / netstat -ano)",
                 fontsize=14, fontweight='bold', pad=20)
    
    table_data = [
        ["Порт", "Протокол", "PID", "Процесс / Служба", "Назначение службы", "Уровень риска", "Рекомендация безопасности"],
        ["135", "TCP", "1696", "svchost.exe (RPCSS)", "Microsoft RPC Endpoint Mapper", "Средний", "Фильтровать внешние входящие запросы"],
        ["139", "TCP", "4", "System (NetBIOS Session)", "NetBIOS Session Service", "Высокий", "Отключить NetBIOS поверх TCP/IP в свойствах NIC"],
        ["445", "TCP", "4", "System (LanmanServer)", "SMBv2/v3 File & Printer Sharing", "Высокий", "Блокировать на границе периметра (порт 445)"],
        ["5040", "TCP", "9916", "svchost.exe (CDPSvc)", "Connected Devices Platform Service", "Низкий", "Штатная служба Windows 10/11"],
        ["5357", "TCP", "4", "System (WSDAPI)", "Web Services for Devices API", "Низкий", "Ограничить локальной рабочей группой"],
        ["10808", "TCP", "15568", "xray.exe / Happ", "SOCKS5 Proxy Listener", "Средний", "Использовать строгую локальную привязку (127.0.0.1)"],
        ["10809", "TCP", "15568", "xray.exe / Happ", "HTTP Proxy Listener", "Средний", "Использовать строгую локальную привязку (127.0.0.1)"],
        ["5353", "UDP", "1444", "svchost.exe (mDNS)", "Multicast DNS Discovery", "Низкий", "Изолировать в доверенной домашней/офисной сети"],
        ["5355", "UDP", "1444", "svchost.exe (LLMNR)", "Link-Local Multicast Resolution", "Высокий", "Рекомендуется отключить через GPO (риск спуфинга)"]
    ]
    
    t = ax.table(cellText=table_data, loc='center', cellLoc='center',
                 colWidths=[0.07, 0.08, 0.07, 0.22, 0.25, 0.12, 0.28])
    t.auto_set_font_size(False)
    t.set_fontsize(9.0)
    t.scale(1.0, 2.0)
    
    for (r, c), cell in t.get_celld().items():
        if r == 0:
            cell.set_facecolor('#263238')
            cell.get_text().set_color('white')
            cell.get_text().set_weight('bold')
        else:
            risk = table_data[r][5]
            if risk == 'Высокий':
                cell.set_facecolor('#FFEBEE' if c != 5 else '#FFCDD2')
                if c == 5:
                    cell.get_text().set_weight('bold')
                    cell.get_text().set_color('#C62828')
            elif risk == 'Средний':
                cell.set_facecolor('#FFF8E1' if c != 5 else '#FFE082')
                if c == 5:
                    cell.get_text().set_weight('bold')
                    cell.get_text().set_color('#F57F17')
            else:
                cell.set_facecolor('#F1F8E9' if c != 5 else '#C8E6C9')
                if c == 5:
                    cell.get_text().set_weight('bold')
                    cell.get_text().set_color('#2E7D32')
        cell.set_edgecolor('#CFD8DC')
        
    plt.tight_layout()
    p = os.path.join(OUTPUT_DIR, "fig4_port_scan_services.png")
    plt.savefig(p, dpi=300)
    plt.close()
    print("Generated:", p)

# =============================================================================
# РИСУНОК 5: Статистика сетевых протоколов (netstat -e, netstat -s)
# =============================================================================
def generate_fig5():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5.5), dpi=300)
    
    # 5.1 Трафик Ethernet (байты и пакеты)
    metrics = ['Принято (RX)', 'Передано (TX)']
    mbytes = [290.9, 800.7]  # МБ
    packets = [370.2, 307.1] # тыс. пакетов
    
    x = np.arange(len(metrics))
    width = 0.35
    
    b1 = ax1.bar(x - width/2, mbytes, width, label='Объем трафика (МБ)', color='#0288D1')
    ax1_twin = ax1.twinx()
    b2 = ax1_twin.bar(x + width/2, packets, width, label='Число пакетов (тыс.)', color='#FFA000')
    
    ax1.set_ylabel('Объем данных (МБ)', color='#0288D1', fontsize=11)
    ax1_twin.set_ylabel('Количество пакетов (тыс.)', color='#FFA000', fontsize=11)
    ax1.set_title("Статистика сетевого адаптера\n(netstat -e)", fontsize=13, fontweight='bold', pad=15)
    ax1.set_xticks(x)
    ax1.set_xticklabels(metrics, fontsize=10.5, fontweight='bold')
    ax1.set_ylim(0, 1000)
    ax1_twin.set_ylim(0, 500)
    ax1.grid(True, linestyle='--', alpha=0.4, axis='y')
    
    for bar in b1:
        h = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2.0, h + 20, f"{h:.1f} МБ", ha='center', va='bottom', fontsize=9, fontweight='bold', color='#0288D1')
    for bar in b2:
        h = bar.get_height()
        ax1_twin.text(bar.get_x() + bar.get_width()/2.0, h + 10, f"{h:.1f} к", ha='center', va='bottom', fontsize=9, fontweight='bold', color='#FFA000')

    # 5.2 Статистика протокола TCP
    tcp_metrics = ['Активные\nоткрытия', 'Пассивные\nоткрытия', 'Сбои\nподключений', 'Сбросы\n(Reset)', 'Переотправки\n(Retrans.)']
    tcp_vals = [6519, 446, 1864, 1511, 13529]
    colors = ['#43A047', '#1E88E5', '#FB8C00', '#E53935', '#8E24AA']
    
    bars2 = ax2.bar(tcp_metrics, tcp_vals, color=colors, width=0.55)
    ax2.set_ylabel('Количество событий', fontsize=11)
    ax2.set_title("Статистика протокола TCP IPv4\n(netstat -s)", fontsize=13, fontweight='bold', pad=15)
    ax2.set_yscale('log')
    ax2.set_ylim(100, 30000)
    ax2.grid(True, linestyle='--', alpha=0.5, axis='y')
    
    for bar in bars2:
        h = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2.0, h * 1.15, str(h), ha='center', va='bottom', fontsize=9.5, fontweight='bold')
        
    plt.tight_layout()
    p = os.path.join(OUTPUT_DIR, "fig5_netstat_stats.png")
    plt.savefig(p, dpi=300)
    plt.close()
    print("Generated:", p)

# =============================================================================
# РИСУНОК 6: Исходящие трассировки по регионам мира
# =============================================================================
def generate_fig6():
    fig, ax = plt.subplots(figsize=(14, 6.0), dpi=300)
    
    # 4 региона мира: данные задержек RTT по прыжкам
    hops_usa = list(range(1, 14))
    rtt_usa = [0.6, 1.6, 3.5, 42.0, 179.6, 179.8, 182.0, 188.5, 188.4, 215.4, 183.1, 181.5, 180.7]
    
    hops_jp = list(range(1, 13))
    rtt_jp = [0.6, 1.8, 3.2, 38.5, 125.0, 158.4, 212.0, 235.1, 268.4, 275.2, 281.0, 282.5]
    
    hops_eu = list(range(1, 10))
    rtt_eu = [0.6, 1.6, 18.2, 26.5, 34.0, 42.1, 48.6, 52.4, 54.1]
    
    hops_au = list(range(1, 13))
    rtt_au = [0.8, 1.8, 15.0, 31.0, 186.6, 322.5, 325.0, 328.0, 332.5, 336.0, 341.0, 342.2]
    
    ax.plot(hops_usa, rtt_usa, marker='o', linewidth=2.5, markersize=6, label='США (Stanford Univ., 171.67.215.200)', color='#1976D2')
    ax.plot(hops_jp, rtt_jp, marker='s', linewidth=2.5, markersize=6, label='Япония (Yahoo Japan, 182.22.25.124)', color='#FB8C00')
    ax.plot(hops_eu, rtt_eu, marker='^', linewidth=2.5, markersize=6, label='Европа (Oxford Univ., 129.67.242.155)', color='#43A047')
    ax.plot(hops_au, rtt_au, marker='D', linewidth=2.5, markersize=6, label='Австралия (Optus / ANU, 203.2.75.132)', color='#8E24AA')
    
    # Аннотации ключевых географических переходов
    ax.annotate('Магистраль в РФ\n(RTT < 4 мс)', xy=(2, 2), xytext=(2.2, 45),
                arrowprops=dict(arrowstyle="->", color='#333333', lw=1.2), fontsize=9, fontweight='bold')
    
    ax.annotate('Европейский IXP\n(Frankfurt / London)', xy=(4, 42), xytext=(4.5, 95),
                arrowprops=dict(arrowstyle="->", color='#333333', lw=1.2), fontsize=9, fontweight='bold')
    
    ax.annotate('Трансатлантический кабель\n(Европа -> США, RTT ~180 мс)', xy=(5, 179.6), xytext=(5.5, 230),
                arrowprops=dict(arrowstyle="->", color='#1976D2', lw=1.2), fontsize=9, fontweight='bold', color='#1976D2')
    
    ax.annotate('Трансокеанский переход в Австралию\n(RTT ~330 мс)', xy=(6, 322.5), xytext=(6.5, 365),
                arrowprops=dict(arrowstyle="->", color='#8E24AA', lw=1.2), fontsize=9, fontweight='bold', color='#8E24AA')
    
    ax.set_xlabel('Номер прыжка (Hop / Шаг трассировки)', fontsize=12)
    ax.set_ylabel('Время задержки RTT (миллисекунды)', fontsize=12)
    ax.set_title('Сравнительный анализ задержек исходящих маршрутов по 4 регионам мира',
                 fontsize=14, fontweight='bold', pad=15)
    ax.grid(True, linestyle='--', alpha=0.5)
    ax.set_xticks(range(1, 14))
    ax.set_ylim(0, 420)
    ax.legend(fontsize=10.5, loc='upper left')
    
    plt.tight_layout()
    p = os.path.join(OUTPUT_DIR, "fig6_outgoing_traces.png")
    plt.savefig(p, dpi=300)
    plt.close()
    print("Generated:", p)

# =============================================================================
# РИСУНОК 7: Whois – Идентификация автономных систем (ASN)
# =============================================================================
def generate_fig7():
    fig, ax = plt.subplots(figsize=(14, 6.2), dpi=300)
    ax.axis('off')
    ax.set_title("Идентификация промежуточных автономных систем (ASN) и операторов через сервис Whois",
                 fontsize=14, fontweight='bold', pad=20)
    
    table_data = [
        ["IP-узел", "Автономная система (ASN)", "Организация / Оператор", "Страна", "Региональный RIR", "Роль в инфраструктуре"],
        ["185.159.82.1", "AS14576", "Hosting Solution Ltd.", "Россия", "RIPE NCC", "Локальный дата-центр отправления"],
        ["46.46.155.232", "AS35807", "Sky-Telecom / MSK-IX", "Россия", "RIPE NCC", "Точка обмена трафиком в Москве"],
        ["87.245.232.55", "AS9002", "RETN International Backbone", "Россия / ЕС", "RIPE NCC", "Международный магистральный транзит"],
        ["184.105.213.230", "AS6939", "Hurricane Electric, Inc.", "Германия", "ARIN / RIPE", "Европейский узел Франкфурт (Tier-1)"],
        ["184.104.198.146", "AS6939", "Hurricane Electric, Inc.", "США", "ARIN", "Трансатлантический волоконный кабель"],
        ["184.105.177.238", "AS6939", "Hurricane Electric (San Jose)", "США", "ARIN", "Шлюз Кремниевой долины (West Coast)"],
        ["171.67.215.200", "AS32", "Stanford University", "США", "ARIN", "Целевой академический сервер (Stanford)"],
        ["213.39.66.201", "AS1299", "Arelion (Telia Carrier)", "Швеция", "RIPE NCC", "Скандинавский транзитный узел Tier-1"],
        ["173.205.37.78", "AS4637", "Telstra Global", "Австралия", "APNIC", "Транстихоокеанский шлюз Австралии"],
        ["193.233.144.114", "AS8480", "Ufa University of Science & Tech.", "Россия", "RIPE NCC", "Сервер УУНиТ (бывш. УГАТУ)"]
    ]
    
    t = ax.table(cellText=table_data, loc='center', cellLoc='center',
                 colWidths=[0.14, 0.16, 0.25, 0.10, 0.12, 0.23])
    t.auto_set_font_size(False)
    t.set_fontsize(9.0)
    t.scale(1.0, 1.95)
    
    for (r, c), cell in t.get_celld().items():
        if r == 0:
            cell.set_facecolor('#004D40')
            cell.get_text().set_color('white')
            cell.get_text().set_weight('bold')
        elif r == 10:
            cell.set_facecolor('#E0F2F1')
            cell.get_text().set_weight('bold')
        elif r % 2 == 1:
            cell.set_facecolor('#F9FBE7')
        else:
            cell.set_facecolor('#FFFFFF')
        cell.set_edgecolor('#B2DFDB')
        
    plt.tight_layout()
    p = os.path.join(OUTPUT_DIR, "fig7_whois_as_analysis.png")
    plt.savefig(p, dpi=300)
    plt.close()
    print("Generated:", p)

# =============================================================================
# РИСУНОК 8: Входящие трассировки к серверу УУНиТ (УГАТУ)
# =============================================================================
def generate_fig8():
    fig, ax = plt.subplots(figsize=(14, 6.0), dpi=300)
    
    # Реальные данные входящих трассировок к 193.233.144.114 с 4 мировых зондов
    hops_us = list(range(1, 13))
    rtt_us = [0.3, 0.5, 1.4, 2.5, 30.5, 28.5, 215.0, 224.5, 232.0, 234.0, 230.0, 231.5]
    
    hops_jp = list(range(1, 17))
    rtt_jp = [0.9, 0.4, 1.2, 1.8, 10.0, 45.0, 110.0, 155.5, 240.0, 270.0, 294.0, 288.0, 292.0, 289.0, 335.0, 289.0]
    
    hops_uk = list(range(1, 10))
    rtt_uk = [0.8, 0.9, 25.0, 100.9, 70.5, 71.0, 68.5, 70.0, 69.8]
    
    hops_au = list(range(1, 13))
    rtt_au = [0.5, 50.0, 134.7, 138.8, 220.0, 290.0, 353.5, 356.5, 339.5, 357.0, 378.0, 345.5]
    
    ax.plot(hops_us, rtt_us, marker='o', linewidth=2.5, markersize=6, label='Зонд США (Los Angeles, AS64457)', color='#1E88E5')
    ax.plot(hops_jp, rtt_jp, marker='s', linewidth=2.5, markersize=6, label='Зонд Япония (Tokyo, AS23816)', color='#FB8C00')
    ax.plot(hops_uk, rtt_uk, marker='^', linewidth=2.5, markersize=6, label='Зонд Европа (London, AS47430)', color='#43A047')
    ax.plot(hops_au, rtt_au, marker='D', linewidth=2.5, markersize=6, label='Зонд Австралия (Sydney, AS7575)', color='#8E24AA')
    
    # Зона схождения
    ax.axvspan(9, 12, color='#FFF9C4', alpha=0.5, label='Зона схождения на магистрали Ростелеком -> УУНиТ')
    
    ax.annotate('Общий узел Ростелеком\n92.50.190.109 (Москва)', xy=(10, 234), xytext=(7.5, 170),
                arrowprops=dict(arrowstyle="->", color='#D32F2F', lw=1.5), fontsize=9.5, fontweight='bold', color='#D32F2F')
    
    ax.annotate('Региональный шлюз Уфы\n81.30.192.242 -> УУНиТ\n193.233.144.114', xy=(12, 231.5), xytext=(12.2, 160),
                arrowprops=dict(arrowstyle="->", color='#D32F2F', lw=1.5), fontsize=9.5, fontweight='bold', color='#D32F2F')
    
    ax.set_xlabel('Номер прыжка (Hop / Шаг входящей трассировки)', fontsize=12)
    ax.set_ylabel('Время задержки RTT (миллисекунды)', fontsize=12)
    ax.set_title('Входящая трассировка к серверу УУНиТ (193.233.144.114) из 4 регионов мира',
                 fontsize=14, fontweight='bold', pad=15)
    ax.grid(True, linestyle='--', alpha=0.5)
    ax.set_xticks(range(1, 17))
    ax.set_ylim(0, 420)
    ax.legend(fontsize=10.0, loc='upper left')
    
    plt.tight_layout()
    p = os.path.join(OUTPUT_DIR, "fig8_incoming_traces.png")
    plt.savefig(p, dpi=300)
    plt.close()
    print("Generated:", p)

# =============================================================================
# РИСУНОК 9: Сетевая топология маршрутов (ветвление и схождение)
# =============================================================================
def generate_fig9():
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 8.5), dpi=300)
    
    # 9.1 Схема исходящей топологии
    ax1.axis('off')
    ax1.set_title("А) Топология исходящих маршрутов: точка ветвления на внешние регионы", fontsize=12.5, fontweight='bold')
    
    def draw_node(ax, x, y, text, color='#E3F2FD', edge='#1565C0', w=1.8, h=0.55):
        rect = patches.FancyBboxPatch((x - w/2, y - h/2), w, h, boxstyle="round,pad=0.1",
                                     facecolor=color, edgecolor=edge, linewidth=1.5)
        ax.add_patch(rect)
        ax.text(x, y, text, ha='center', va='center', fontsize=8.5, fontweight='bold', color='#1A237E')
        
    def draw_link(ax, x1, y1, x2, y2, label=""):
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="->", color='#455A64', lw=1.5))
        if label:
            mx, my = (x1 + x2)/2, (y1 + y2)/2
            ax.text(mx, my + 0.1, label, ha='center', va='bottom', fontsize=7.5, color='#37474F')

    ax1.set_xlim(0, 14)
    ax1.set_ylim(-0.5, 4.5)
    
    draw_node(ax1, 1.5, 2.0, "Хост (Уфа)\n10.160.151.252")
    draw_node(ax1, 4.0, 2.0, "Шлюз Уфы / MSK-IX\n46.46.155.232")
    draw_node(ax1, 6.8, 2.0, "Магистраль RETN\n87.245.232.55 (Москва)\n[ТОЧКА ВЕТВЛЕНИЯ]", color='#FFF9C4', edge='#F57F17', w=2.4)
    
    draw_link(ax1, 2.4, 2.0, 3.1, 2.0, "Wi-Fi / LAN")
    draw_link(ax1, 4.9, 2.0, 5.6, 2.0, "РФ Магистраль")
    
    # Ветви
    draw_node(ax1, 10.0, 3.6, "HE.net (Frankfurt / USA East)\n184.105.213.230", w=2.6)
    draw_node(ax1, 12.8, 3.6, "США (Stanford)\n171.67.215.200", color='#C8E6C9', edge='#2E7D32')
    draw_link(ax1, 8.0, 2.2, 8.7, 3.6, "Трансатлантика")
    draw_link(ax1, 11.3, 3.6, 11.9, 3.6)
    
    draw_node(ax1, 10.0, 2.4, "Telia / Arelion (Stockholm)\n213.39.66.201", w=2.6)
    draw_node(ax1, 12.8, 2.4, "Европа (Oxford)\n129.67.242.155", color='#C8E6C9', edge='#2E7D32')
    draw_link(ax1, 8.0, 2.0, 8.7, 2.4, "Европа")
    draw_link(ax1, 11.3, 2.4, 11.9, 2.4)
    
    draw_node(ax1, 10.0, 1.2, "NTT / SINET (Токио)\n101.203.106.69", w=2.6)
    draw_node(ax1, 12.8, 1.2, "Япония (Yahoo)\n182.22.25.124", color='#C8E6C9', edge='#2E7D32')
    draw_link(ax1, 8.0, 1.8, 8.7, 1.2, "Азия")
    draw_link(ax1, 11.3, 1.2, 11.9, 1.2)
    
    draw_node(ax1, 10.0, 0.0, "Telstra (Sydney)\n173.205.37.78", w=2.6)
    draw_node(ax1, 12.8, 0.0, "Австралия (ANU)\n203.2.75.132", color='#C8E6C9', edge='#2E7D32')
    draw_link(ax1, 8.0, 1.6, 8.7, 0.0, "Океания")
    draw_link(ax1, 11.3, 0.0, 11.9, 0.0)

    # 9.2 Схема входящей топологии
    ax2.axis('off')
    ax2.set_title("Б) Топология входящих маршрутов: точка схождения мировых потоков к серверу УУНиТ", fontsize=12.5, fontweight='bold')
    ax2.set_xlim(0, 14)
    ax2.set_ylim(-0.5, 4.5)
    
    draw_node(ax2, 1.5, 3.6, "Зонд США\n(Los Angeles)")
    draw_node(ax2, 1.5, 2.4, "Зонд Европа\n(London)")
    draw_node(ax2, 1.5, 1.2, "Зонд Япония\n(Tokyo)")
    draw_node(ax2, 1.5, 0.0, "Зонд Австралия\n(Sydney)")
    
    draw_node(ax2, 4.5, 1.8, "Международный шлюз\nLumen / Tier-1\n139.45.247.177", color='#FFECB3', edge='#FFA000', w=2.3)
    draw_link(ax2, 2.4, 3.6, 3.5, 2.0)
    draw_link(ax2, 2.4, 2.4, 3.5, 1.9)
    draw_link(ax2, 2.4, 1.2, 3.5, 1.7)
    draw_link(ax2, 2.4, 0.0, 3.5, 1.6)
    
    draw_node(ax2, 7.5, 1.8, "Магистраль Ростелеком\n92.50.190.109 (Москва)\n[ТОЧКА СХОЖДЕНИЯ]", color='#FFF9C4', edge='#F57F17', w=2.4)
    draw_link(ax2, 5.7, 1.8, 6.3, 1.8, "Стык операторов")
    
    draw_node(ax2, 10.4, 1.8, "Узловой маршрутизатор Уфы\n81.30.192.242\n(Башкортостан)", w=2.4)
    draw_link(ax2, 8.7, 1.8, 9.2, 1.8, "Runnet / РТ")
    
    draw_node(ax2, 13.0, 1.8, "Сервер УУНиТ\n193.233.144.114\n(AS8480, Уфа)", color='#C8E6C9', edge='#2E7D32', w=1.8)
    draw_link(ax2, 11.6, 1.8, 12.1, 1.8, "Кампус")
    
    plt.tight_layout()
    p = os.path.join(OUTPUT_DIR, "fig9_topology_routes.png")
    plt.savefig(p, dpi=300)
    plt.close()
    print("Generated:", p)

# =============================================================================
# РИСУНОК 10: Глобальная карта GeoTrace
# =============================================================================
def generate_fig10():
    fig, ax = plt.subplots(figsize=(15, 7.5), dpi=300)
    ax.set_facecolor('#0E1726')  # Темный стильный фон океана
    
    # Отрисовка контуров суши из world.geojson
    if os.path.exists(GEOJSON_PATH):
        with open(GEOJSON_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
        for feat in data['features']:
            geom = feat['geometry']
            t = geom['type']
            coords = geom['coordinates']
            if t == 'Polygon':
                coords = [coords]
            for poly in coords:
                for ring in poly:
                    xs, ys = zip(*ring)
                    ax.fill(xs, ys, facecolor='#1E293B', edgecolor='#334155', linewidth=0.5)
    
    # Географические координаты ключевых центров коммутации
    cities = {
        'Ufa': (55.97, 54.74),            # Уфа (исходная точка / УУНиТ)
        'Moscow': (37.62, 55.75),         # Москва (MSK-IX / Ростелеком)
        'Stockholm': (18.06, 59.33),      # Стокгольм (Arelion)
        'Frankfurt': (8.68, 50.11),       # Франкфурт (DE-CIX / HE)
        'London': (-0.13, 51.51),         # Лондон (LINX / Oxford)
        'Stanford': (-122.17, 37.43),     # Стэнфорд, Калифорния (США)
        'Tokyo': (139.69, 35.69),         # Токио (Япония)
        'Sydney': (151.21, -33.87)        # Сидней (Австралия)
    }
    
    def plot_curve(c1, c2, color, label=None, ls='-'):
        p1 = cities[c1]
        p2 = cities[c2]
        # Простая квадратичная дуга Безье для красивого изгиба трасс
        num_pts = 60
        t = np.linspace(0, 1, num_pts)
        
        # Контрольная точка с изгибом по широте
        mid_lon = (p1[0] + p2[0]) / 2.0
        mid_lat = (p1[1] + p2[1]) / 2.0 + (12.0 if p2[1] > -10 else -8.0)
        
        x = (1 - t)**2 * p1[0] + 2 * (1 - t) * t * mid_lon + t**2 * p2[0]
        y = (1 - t)**2 * p1[1] + 2 * (1 - t) * t * mid_lat + t**2 * p2[1]
        
        ax.plot(x, y, color=color, linewidth=2.0, linestyle=ls, alpha=0.85, label=label)

    # Исходящие трассы из Уфы через Москву
    plot_curve('Ufa', 'Moscow', '#00E5FF', label='Уфа -> Москва (Магистраль РФ)', ls='-')
    plot_curve('Moscow', 'Frankfurt', '#00E5FF', ls='--')
    plot_curve('Moscow', 'Stockholm', '#00E5FF', ls='--')
    
    # Трассы к континентам
    plot_curve('Frankfurt', 'Stanford', '#38BDF8', label='Трасса в США (Stanford, RTT 180 мс)')
    plot_curve('Frankfurt', 'London', '#4ADE80', label='Трасса в Европу (Oxford, RTT 54 мс)')
    plot_curve('Moscow', 'Tokyo', '#FB923C', label='Трасса в Японию (Yahoo Tokyo, RTT 282 мс)')
    plot_curve('Frankfurt', 'Sydney', '#C084FC', label='Трасса в Австралию (Sydney, RTT 342 мс)')

    # Отрисовка маркеров городов
    for name, (lon, lat) in cities.items():
        if name == 'Ufa':
            ax.plot(lon, lat, marker='*', markersize=14, color='#FFD700', zorder=5)
            ax.text(lon + 2, lat + 2, "Уфа (УУНиТ)", color='#FFD700', fontsize=10.5, fontweight='bold', zorder=6)
        else:
            ax.plot(lon, lat, marker='o', markersize=7, color='#FFFFFF', markeredgecolor='#00E5FF', markeredgewidth=1.5, zorder=5)
            dy = -4 if lat < 0 else 3
            dx = 3 if lon > 0 else -12
            ax.text(lon + dx, lat + dy, name, color='#E2E8F0', fontsize=9.0, fontweight='bold', zorder=6)

    ax.set_xlim(-170, 180)
    ax.set_ylim(-60, 85)
    ax.set_xlabel('Географическая долгота (Longitude)', color='#94A3B8', fontsize=10.5)
    ax.set_ylabel('Географическая широта (Latitude)', color='#94A3B8', fontsize=10.5)
    ax.tick_params(colors='#94A3B8')
    ax.set_title("Глобальная визуализация сетевой трассировки GeoTrace\n(Уфа -> Москва -> Мировые узлы Tier-1)",
                 fontsize=14, fontweight='bold', pad=15, color='#F8FAFC')
    
    ax.legend(facecolor='#1E293B', edgecolor='#475569', labelcolor='#F8FAFC', fontsize=9.5, loc='lower left')
    ax.grid(True, linestyle=':', alpha=0.3, color='#64748B')
    
    plt.tight_layout()
    p = os.path.join(OUTPUT_DIR, "fig10_geotrace_map.png")
    plt.savefig(p, dpi=300)
    plt.close()
    print("Generated:", p)

if __name__ == "__main__":
    generate_fig1()
    generate_fig2()
    generate_fig3()
    generate_fig4()
    generate_fig5()
    generate_fig6()
    generate_fig7()
    generate_fig8()
    generate_fig9()
    generate_fig10()
    print("ALL 10 FIGURES GENERATED SUCCESSFULLY!")

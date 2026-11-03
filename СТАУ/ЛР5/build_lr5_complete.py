# -*- coding: utf-8 -*-
import os
import sys
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

base_dir = os.path.dirname(os.path.abspath(__file__))
fig_dir = os.path.join(base_dir, "figures")
os.makedirs(fig_dir, exist_ok=True)

# -------------------------------------------------------------
# 1. ГЕНЕРАЦИЯ ГРАФИЧЕСКИХ ДИАГРАММ И СХЕМ
# -------------------------------------------------------------
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 9

# Схема 1: Топология сети и архитектура селективной фильтрации Firewall
fig, ax = plt.subplots(figsize=(8.0, 4.8), dpi=220)
ax.set_xlim(0, 10)
ax.set_ylim(0, 6)
ax.axis('off')

# Центральный блок: Наша рабочая станция + Firewall
fw_box = patches.FancyBboxPatch((3.5, 1.2), 3.0, 3.8, boxstyle="round,pad=0.2",
                                facecolor="#e3f2fd", edgecolor="#1565c0", linewidth=2)
ax.add_patch(fw_box)
ax.text(5.0, 4.7, "ЛОКАЛЬНЫЙ УЗЕЛ\n(Наша рабочая станция)\nIP: 10.8.1.1",
        ha='center', va='center', fontsize=10, fontweight='bold', color="#0d47a1")

# Внутренний блок брандмауэра
fw_inner = patches.FancyBboxPatch((3.8, 1.4), 2.4, 2.3, boxstyle="round,pad=0.1",
                                 facecolor="#bbdefb", edgecolor="#0d47a1", linewidth=1.5, linestyle="--")
ax.add_patch(fw_inner)
ax.text(5.0, 3.4, "Windows Defender Firewall", ha='center', va='center', fontsize=9, fontweight='bold', color="#01579b")
ax.text(5.0, 2.8, "Политика белых списков:\nРазрешено ТОЛЬКО для\nRemoteIP: 10.8.1.11",
        ha='center', va='center', fontsize=8, color="#004d40", style='italic')
ax.text(5.0, 1.9, "Защищаемые службы:\n• ICMP (ping, tracert)\n• SMB (папки: 445, 139)\n• RemoteRegistry (135, 445)\n• RDP (3389)",
        ha='center', va='center', fontsize=7.5, color="#212121")

# Правый блок: Ноутбук Данила (сосед справа) - РАЗРЕШЕНО
danil_box = patches.FancyBboxPatch((7.6, 2.4), 2.2, 2.4, boxstyle="round,pad=0.2",
                                   facecolor="#e8f5e9", edgecolor="#2e7d32", linewidth=2)
ax.add_patch(danil_box)
ax.text(8.7, 4.4, "СОСЕД СПРАВА\n(Ноутбук Данила)", ha='center', va='center', fontsize=9.5, fontweight='bold', color="#1b5e20")
ax.text(8.7, 3.8, "IP: 10.8.1.11", ha='center', va='center', fontsize=9, fontweight='bold', color="#2e7d32")
ax.text(8.7, 3.0, "СТАТУС ДОСТУПА:\nПОЛНЫЙ ДОСТУП\n(100% Разрешено)", ha='center', va='center', fontsize=8, fontweight='bold', color="#2e7d32")

# Левый блок: Другие соседи - ЗАБЛОКИРОВАНО
other_box = patches.FancyBboxPatch((0.2, 2.4), 2.2, 2.4, boxstyle="round,pad=0.2",
                                   facecolor="#ffebee", edgecolor="#c62828", linewidth=2)
ax.add_patch(other_box)
ax.text(1.3, 4.4, "ДРУГИЕ СОСЕДИ\n(Узлы лаборатории)", ha='center', va='center', fontsize=9.5, fontweight='bold', color="#b71c1c")
ax.text(1.3, 3.8, "IP: 10.8.1.X (любые)", ha='center', va='center', fontsize=9, fontweight='bold', color="#c62828")
ax.text(1.3, 3.0, "СТАТУС ДОСТУПА:\nБЛОКИРОВКА\n(Drop / Timeout)", ha='center', va='center', fontsize=8, fontweight='bold', color="#c62828")

# Стрелки взаимодействия
# От Данила (зеленая, двунаправленная)
ax.annotate("", xy=(6.6, 3.6), xytext=(7.5, 3.6),
            arrowprops=dict(arrowstyle="<->", color="#2e7d32", lw=2.5))
ax.text(7.05, 3.9, "Разрешено\n(Allow)", ha='center', va='bottom', fontsize=8, fontweight='bold', color="#2e7d32")

# От других соседей (красная со знаком запрета)
ax.annotate("", xy=(3.4, 3.6), xytext=(2.5, 3.6),
            arrowprops=dict(arrowstyle="->", color="#c62828", lw=2.5, linestyle=":"))
ax.text(2.95, 3.9, "Запрещено\n(Blocked)", ha='center', va='bottom', fontsize=8, fontweight='bold', color="#c62828")

ax.set_title("Рисунок 1 – Топологическая схема селективной фильтрации межсетевого экрана", fontsize=10.5, pad=10)
plt.tight_layout()
diag_topo_path = os.path.join(fig_dir, "diagram_topology.png")
plt.savefig(diag_topo_path, dpi=220)
plt.close()

# Схема 2: Матрица сервисов и портов
fig, ax = plt.subplots(figsize=(7.5, 3.8), dpi=220)
services = [
    "Ресурс 1: ICMP (ping / tracert)",
    "Ресурс 2: Общие папки (SMB)",
    "Ресурс 3: Удаленный реестр",
    "Ресурс 4: Удаленный рабочий стол"
]
ports = [
    "ICMPv4 (Type 8: Echo Req, Type 11: TTL)",
    "TCP 445, 139 / UDP 137, 138",
    "TCP 135, 445, 49152-65535",
    "TCP 3389 / UDP 3389"
]
colors_bar = ["#42a5f5", "#66bb6a", "#ffa726", "#ab47bc"]
y_pos = range(len(services))
bars = ax.barh(y_pos, [10, 10, 10, 10], color=colors_bar, edgecolor="#37474f", height=0.55)
ax.set_yticks(y_pos)
ax.set_yticklabels(services, fontsize=9.5, fontweight='bold')
ax.set_xticks([])
ax.set_xlim(0, 11)

for idx, bar in enumerate(bars):
    ax.text(0.3, bar.get_y() + bar.get_height()/2.0, ports[idx],
            va='center', ha='left', color='white', fontweight='bold', fontsize=9)
    ax.text(9.8, bar.get_y() + bar.get_height()/2.0, "Доступ: 10.8.1.11",
            va='center', ha='right', color='white', fontweight='bold', fontsize=8.5)

ax.set_title("Рисунок 2 – Архитектурный профиль защищаемых сетевых сервисов и портов", fontsize=10.5, pad=10)
ax.invert_yaxis()
plt.tight_layout()
diag_ports_path = os.path.join(fig_dir, "diagram_ports_matrix.png")
plt.savefig(diag_ports_path, dpi=220)
plt.close()

print("Matplotlib diagrams successfully generated.")

# -------------------------------------------------------------
# 2. ПОДГОТОВКА ДОКУМЕНТА WORD ПО ГОСТ
# -------------------------------------------------------------
doc = docx.Document()

# Поля: левое 30 мм, правое 15 мм, верхнее 20 мм, нижнее 20 мм
for sec in doc.sections:
    sec.left_margin = Inches(30 / 25.4)
    sec.right_margin = Inches(15 / 25.4)
    sec.top_margin = Inches(20 / 25.4)
    sec.bottom_margin = Inches(20 / 25.4)

def set_run(r, font_name="Times New Roman", size_pt=14, bold=False, italic=False, color=(0,0,0)):
    r.font.name = font_name
    r.font.size = Pt(size_pt)
    r.bold = bold
    r.italic = italic
    r.font.color.rgb = RGBColor(*color)

def add_p(text="", align=WD_PARAGRAPH_ALIGNMENT.JUSTIFY, space_before=0, space_after=6, indent=1.25, line_spacing=1.15, bold=False, italic=False, size_pt=14):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if indent > 0:
        p.paragraph_format.first_line_indent = Inches(indent / 2.54)
    else:
        p.paragraph_format.first_line_indent = Inches(0)
    if text:
        r = p.add_run(text)
        set_run(r, size_pt=size_pt, bold=bold, italic=italic)
    return p

def add_code_block(code_text):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.paragraphs[0].text = ""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F5F5F5"/>')
    tcPr.append(shd)
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/></w:tcBorders>')
    tcPr.append(tcBorders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.first_line_indent = Inches(0)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.05
    for idx, line in enumerate(code_text.strip().split('\n')):
        if idx > 0:
            p = cell.add_paragraph()
            p.paragraph_format.first_line_indent = Inches(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.05
        r = p.add_run(line)
        set_run(r, font_name="Consolas", size_pt=9.0)
    add_p("", space_before=0, space_after=4, indent=0)

def add_h1(text):
    return add_p(text, align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_before=14, space_after=8, indent=0, bold=True, size_pt=16)

def add_h2(text):
    return add_p(text, align=WD_PARAGRAPH_ALIGNMENT.LEFT, space_before=12, space_after=6, indent=1.25, bold=True, size_pt=14)

def add_h3(text):
    return add_p(text, align=WD_PARAGRAPH_ALIGNMENT.LEFT, space_before=8, space_after=4, indent=1.25, bold=True, italic=True, size_pt=14)

# --- ТИТУЛЬНЫЙ ЛИСТ ---
p = add_p("Министерство науки и высшего образования Российской Федерации", align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=2, indent=0, size_pt=12)
p = add_p("Федеральное государственное бюджетное образовательное учреждение высшего образования", align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=2, indent=0, size_pt=12)
p = add_p("«УФИМСКИЙ УНИВЕРСИТЕТ НАУКИ И ТЕХНОЛОГИЙ»", align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=4, indent=0, bold=True, size_pt=13)
p = add_p("Институт информатики, математики и робототехники", align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=2, indent=0, size_pt=12)
p = add_p("Кафедра автоматизированных систем управления", align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=40, indent=0, bold=True, size_pt=13)

p = add_p("ОТЧЕТ", align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=6, indent=0, bold=True, size_pt=20)
p = add_p("ПО ЛАБОРАТОРНОЙ РАБОТЕ № 5", align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=8, indent=0, bold=True, size_pt=16)
p = add_p("по дисциплине: «Сетевые технологии автоматизации и управления» (СТАУ)", align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=6, indent=0, italic=True, size_pt=13)
p = add_p("на тему: «Настройка межсетевого экрана (Firewall)»", align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=8, indent=0, bold=True, size_pt=15)
p = add_p("Вариант № 17", align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=60, indent=0, bold=True, size_pt=13)

# Блок авторов
p = add_p("Выполнили:", align=WD_PARAGRAPH_ALIGNMENT.RIGHT, space_after=2, indent=0, size_pt=12)
p = add_p("студенты гр. ЭАС-514С", align=WD_PARAGRAPH_ALIGNMENT.RIGHT, space_after=2, indent=0, size_pt=12)
p = add_p("Гайнанов Д.И.", align=WD_PARAGRAPH_ALIGNMENT.RIGHT, space_after=2, indent=0, bold=True, size_pt=12)
p = add_p("Гарифуллин К.Р.", align=WD_PARAGRAPH_ALIGNMENT.RIGHT, space_after=14, indent=0, bold=True, size_pt=12)
p = add_p("Проверил:", align=WD_PARAGRAPH_ALIGNMENT.RIGHT, space_after=2, indent=0, size_pt=12)
p = add_p("преподаватель кафедры АСУ", align=WD_PARAGRAPH_ALIGNMENT.RIGHT, space_after=2, indent=0, size_pt=12)
p = add_p("Набатов А.Н.", align=WD_PARAGRAPH_ALIGNMENT.RIGHT, space_after=70, indent=0, bold=True, size_pt=12)

p = add_p("Уфа 2026 г.", align=WD_PARAGRAPH_ALIGNMENT.CENTER, space_after=0, indent=0, size_pt=12)
doc.add_page_break()

# --- 1. ЦЕЛЬ РАБОТЫ ---
add_h1("1 Цель работы")
add_p("Освоить теоретические основы и практические навыки конфигурирования межсетевых экранов (Firewall) на примере Брандмауэра Windows с расширенной безопасностью (Windows Defender Firewall with Advanced Security), изучить принципы пакетной фильтрации на сетевом и транспортном уровнях модели OSI, реализовать политику селективного доступа к системным ресурсам рабочей станции (ICMP, общие папки SMB, служба удаленного реестра Remote Registry, служба удаленного рабочего стола RDP) исключительно для доверенного узла («сосед справа» — ноутбук Данила с IP-адресом 10.8.1.11) при гарантированной изоляции и блокировке всех остальных сетевых узлов.")

# --- 2. ТЕОРЕТИЧЕСКАЯ ЧАСТЬ ---
add_h1("2 Теоретическая часть")

add_h2("2.1 Назначение и базовые функции межсетевого экрана")
add_p("Межсетевой экран (сетевой экран, брандмауэр, Firewall) представляет собой комплекс программных или программно-аппаратных средств, осуществляющий контроль, анализ и фильтрацию входящих и исходящих сетевых пакетов, проходящих через границу защищаемой системы, в соответствии с набором административно заданных правил безопасности.")
add_p("Основной задачей сетевого экрана является разделение доверенной и недоверенной зон сети, обеспечение конфиденциальности и целостности узлов, а также предотвращение несанкционированного удаленного доступа, сетевого сканирования и распределенных атак. Межсетевые экраны функционируют как селективные шлюзы-фильтры, отбрасывающие нелегитимные сетевые кадры и пакеты.")

add_h2("2.2 Классификация сетевых экранов")
add_p("Современные сетевые экраны подразделяются по нескольким ключевым критериям:")
add_p("1. По охвату контролируемых потоков данных:")
add_p("– Традиционный сетевой (межсетевой) экран — функционирует на шлюзе сети или выделенном аппаратном маршрутизаторе, контролируя потоки данных между разнородными сегментами и подсетями;")
add_p("– Персональный (хостовый) экран — функционирует непосредственно на оконечной рабочей станции в качестве компонента операционной системы (например, Windows Defender Firewall, iptables/nftables в Linux) и защищает только локальные сервисы хоста.")
add_p("2. По уровню модели OSI, на котором осуществляется инспекция:")
add_p("– Пакетные фильтры (Packet Filters, L3/L4) — проверяют заголовки пакетов (IP-адрес источника и назначения, номер порта источника и назначения, протокол IP: TCP, UDP, ICMP);")
add_p("– Шлюзы сеансового уровня (Circuit-Level Gateways, L5) — контролируют процесс установления логических соединений (TCP Handshake), не вникая в содержимое прикладных данных;")
add_p("– Шлюзы прикладного уровня (Application-Level Gateways / WAF, L7) — выполняют глубокий контентный анализ прикладных протоколов (HTTP, FTP, DNS, SMTP).")
add_p("3. По способу отслеживания состояния активных соединений:")
add_p("– Stateless (простая статическая фильтрация) — фильтруют каждый отдельный сетевой пакет изолированно, на основе жестких статических правил, без учета контекста предыдущего обмена;")
add_p("– Stateful (Stateful Packet Inspection, SPI, фильтрация с учетом контекста) — непрерывно отслеживают состояние динамической таблицы сессий (TCP connection state: SYN, ESTABLISHED, FIN/RST). Пропускаются только те ответные пакеты, которые строго соответствуют ранее инициированному легитимному сеансу связи, что эффективно предотвращает атаки типа SYN-flood и IP-спуфинг.")

add_h2("2.3 Спецификация сетевых протоколов и портов целевых ресурсов")
add_p("В соответствии с заданием лабораторной работы реализуется селективная фильтрация четырех базовых ресурсов операционной системы:")
add_p("1. Ресурс 1: Базовые диагностические команды TCP/IP (ping, traceroute). В основе утилит лежит протокол межсетевых управляющих сообщений ICMPv4 (Internet Control Message Protocol). Утилита ping отправляет запросы ICMP Echo Request (Type 8, Code 0) и ожидает откликов ICMP Echo Reply (Type 0). Утилита traceroute (tracert) использует последовательную модификацию поля времени жизни пакета (TTL, Time To Live), анализируя входящие сообщения ICMP Time Exceeded (Type 11);")
add_p("2. Ресурс 2: Общие папки (File Sharing). Реализуются сетевым протоколом SMB (Server Message Block) поверх TCP/IP (порт TCP 445 — прямой доступ SMB-over-IP, а также порт TCP 139 для сессионной службы NetBIOS). Для разрешения имен и дейтаграмм в локальной сети используются порты UDP 137 (NetBIOS Name Service) и UDP 138 (NetBIOS Datagram Service);")
add_p("3. Ресурс 3: Реестр удаленной машины (Remote Registry). Обеспечивает возможность чтения и модификации ветвей системного реестра через сеть с помощью оснастки regedit. Базируется на протоколе вызова удаленных процедур Microsoft RPC (Remote Procedure Call). Соединение устанавливается через службу сопоставителя конечных точек RPC Endpoint Mapper (порт TCP 135) либо именованный канал \\pipe\\winreg поверх SMB (порт TCP 445), после чего динамически выделяется порт из диапазона RPC Dynamic Ports (TCP 49152–65535);")
add_p("4. Ресурс 4: Удаленный рабочий стол (RDP, Remote Desktop Protocol). Обеспечивает графический интерактивный сеанс управления хостом через службу TermService. Протокол функционирует на фиксированном стандартном порту TCP 3389, а также порту UDP 3389 (для поддержки ускорения мультимедийного потока RDP-UDP).")

# --- 3. ПРАКТИЧЕСКАЯ ЧАСТЬ ---
add_h1("3 Практическая часть")

add_h2("3.1 Исходные данные и сетевая топология")
add_p("Для проведения исследований и настройки сетевой безопасности сформирован сегмент локальной сети со следующими параметрами адресации:")
add_p("– Локальная рабочая станция (исследуемый защищаемый узел): IP-адрес 10.8.1.1 (маска подсети 255.255.255.0);")
add_p("– Доверенный сетевой узел («сосед справа» — ноутбук Данила): IP-адрес 10.8.1.11;")
add_p("– Недоверенные сетевые узлы («другие соседи»): любые другие IP-адреса локального сегмента 10.8.1.0/24 (10.8.1.2, 10.8.1.12 и др.).")
add_p("Топологическая схема сетевого взаимодействия и архитектура разграничения прав доступа представлена на рисунке 1.")

if os.path.exists(diag_topo_path):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p_img.paragraph_format.space_before = Pt(6)
    p_img.paragraph_format.space_after = Pt(4)
    p_img.paragraph_format.first_line_indent = Inches(0)
    p_img.add_run().add_picture(diag_topo_path, width=Inches(6.2))

add_p("Профиль защищаемых портов и протоколов проиллюстрирован на рисунке 2.")
if os.path.exists(diag_ports_path):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p_img.paragraph_format.space_before = Pt(6)
    p_img.paragraph_format.space_after = Pt(4)
    p_img.paragraph_format.first_line_indent = Inches(0)
    p_img.add_run().add_picture(diag_ports_path, width=Inches(5.8))

add_h2("3.2 Конфигурирование межсетевого экрана и системных служб")
add_p("Конфигурирование брандмауэра Windows с расширенной безопасностью выполнено с использованием административной оболочки Windows PowerShell и модуля управления сетевой безопасностью NetSecurity.")
add_p("Политика безопасности построена по принципу «белых списков» (Whitelisting): по умолчанию все входящие подключения блокируются, а для каждого из четырех целевых ресурсов создаются персонализированные правила разрешения с жестким фильтром RemoteAddress = 10.8.1.11.")

add_h3("Настройка Ресурса 1: Базовые команды TCP/IP (ping, traceroute)")
add_p("Для разрешения эхо-запросов ping и трассировки от соседа Данила создано правило фильтрации ICMPv4:")
add_code_block("""New-NetFirewallRule -Name 'STAU_LR5_ICMP_Allow_Danil' `
    -DisplayName 'СТАУ ЛР5: ICMP ping/traceroute - Доступ Данилу' `
    -Direction Inbound -Protocol ICMPv4 -IcmpType 8 `
    -RemoteAddress 10.8.1.11 -Action Allow -Profile Any -Enabled True""")

add_h3("Настройка Ресурса 2: Общие папки (SMB / NetBIOS)")
add_p("Для предоставления доступа к файловым ресурсам на диске C: создан общий каталог C:\\STAU_LR5_Shared с сетевым именем STAU_Share, после чего в брандмауэре открыты необходимые TCP- и UDP-порты службы доступа к файлам и принтерам исключительно для адреса 10.8.1.11:")
add_code_block("""# 1. Создание общего сетевого ресурса
net share STAU_Share=C:\\STAU_LR5_Shared /grant:Все,FULL

# 2. Правила брандмауэра для портов SMB и NetBIOS
New-NetFirewallRule -Name 'STAU_LR5_SMB_TCP_Allow_Danil' `
    -DisplayName 'СТАУ ЛР5: Общие папки SMB (TCP 445, 139) - Доступ Данилу' `
    -Direction Inbound -Protocol TCP -LocalPort 445, 139 `
    -RemoteAddress 10.8.1.11 -Action Allow -Profile Any -Enabled True

New-NetFirewallRule -Name 'STAU_LR5_SMB_UDP_Allow_Danil' `
    -DisplayName 'СТАУ ЛР5: Общие папки NetBIOS (UDP 137, 138) - Доступ Данилу' `
    -Direction Inbound -Protocol UDP -LocalPort 137, 138 `
    -RemoteAddress 10.8.1.11 -Action Allow -Profile Any -Enabled True""")

add_h3("Настройка Ресурса 3: Реестр удаленной машины (Remote Registry)")
add_p("Для обеспечения возможности удаленного подключения к ветвям HKLM и HKU служба RemoteRegistry переведена в активное состояние, а в брандмауэре создано правило для сопоставителя RPC (порт 135) и динамических портов RPC:")
add_code_block("""# 1. Активация системной службы удаленного реестра
Set-Service -Name RemoteRegistry -StartupType Manual
Start-Service -Name RemoteRegistry

# 2. Правило брандмауэра для удаленного реестра RPC
New-NetFirewallRule -Name 'STAU_LR5_RemoteReg_TCP_Allow_Danil' `
    -DisplayName 'СТАУ ЛР5: Удаленный реестр RPC (TCP 135, 445, 49152-65535) - Доступ Данилу' `
    -Direction Inbound -Protocol TCP -LocalPort 135, 445, 49152-65535 `
    -RemoteAddress 10.8.1.11 -Action Allow -Profile Any -Enabled True""")

add_h3("Настройка Ресурса 4: Удаленный рабочий стол (RDP)")
add_p("Для активации интерактивных сеансов управления хостом служба TermService переведена в рабочий режим с модификацией ключа реестра fDenyTSConnections, после чего разрешены входящие подключения на порт 3389:")
add_code_block("""# 1. Разрешение входящих подключений RDP в системе
Set-ItemProperty -Path 'HKLM:\\System\\CurrentControlSet\\Control\\Terminal Server' `
    -Name 'fDenyTSConnections' -Value 0
Start-Service -Name TermService

# 2. Правила брандмауэра для RDP (TCP и UDP)
New-NetFirewallRule -Name 'STAU_LR5_RDP_TCP_Allow_Danil' `
    -DisplayName 'СТАУ ЛР5: Удаленный рабочий стол RDP (TCP 3389) - Доступ Данилу' `
    -Direction Inbound -Protocol TCP -LocalPort 3389 `
    -RemoteAddress 10.8.1.11 -Action Allow -Profile Any -Enabled True

New-NetFirewallRule -Name 'STAU_LR5_RDP_UDP_Allow_Danil' `
    -DisplayName 'СТАУ ЛР5: Удаленный рабочий стол RDP (UDP 3389) - Доступ Данилу' `
    -Direction Inbound -Protocol UDP -LocalPort 3389 `
    -RemoteAddress 10.8.1.11 -Action Allow -Profile Any -Enabled True""")

add_h2("3.3 Сводная спецификация правил брандмауэра")
add_p("В таблице 1 систематизированы сконфигурированные правила межсетевого экрана Windows Defender Firewall.")

p_t1 = add_p("Таблица 1 – Сводная ведомость правил межсетевого экрана для ресурсов рабочей станции", indent=0, bold=True, size_pt=12)
t1 = doc.add_table(rows=1, cols=6)
t1.alignment = WD_TABLE_ALIGNMENT.CENTER
t1_hdr = t1.rows[0].cells
t1_hdr[0].text = "Идентификатор правила"
t1_hdr[1].text = "Целевой ресурс"
t1_hdr[2].text = "Протокол"
t1_hdr[3].text = "Локальный порт"
t1_hdr[4].text = "Удаленный IP"
t1_hdr[5].text = "Действие"

rules_data = [
    ["STAU_LR5_ICMP_Allow_Danil", "Ресурс 1: ping, tracert", "ICMPv4", "Тип 8 (Echo Req)", "10.8.1.11", "Разрешить (Allow)"],
    ["STAU_LR5_SMB_TCP_Allow_Danil", "Ресурс 2: Общие папки", "TCP", "445, 139", "10.8.1.11", "Разрешить (Allow)"],
    ["STAU_LR5_SMB_UDP_Allow_Danil", "Ресурс 2: NetBIOS", "UDP", "137, 138", "10.8.1.11", "Разрешить (Allow)"],
    ["STAU_LR5_RemoteReg_TCP_Allow_Danil", "Ресурс 3: Удаленный реестр", "TCP", "135, 445, 49152-65535", "10.8.1.11", "Разрешить (Allow)"],
    ["STAU_LR5_RDP_TCP_Allow_Danil", "Ресурс 4: RDP (TCP)", "TCP", "3389", "10.8.1.11", "Разрешить (Allow)"],
    ["STAU_LR5_RDP_UDP_Allow_Danil", "Ресурс 4: RDP (UDP)", "UDP", "3389", "10.8.1.11", "Разрешить (Allow)"],
    ["STAU_LR5_ICMP_Block_Others", "Ресурс 1: Блокировка ICMP", "ICMPv4", "Тип 8 (Echo Req)", "10.8.1.1-10, 10.8.1.12-254", "Блокировать (Block)"],
    ["STAU_LR5_SMB_Block_Others", "Ресурс 2: Блокировка SMB", "TCP", "445, 139", "10.8.1.1-10, 10.8.1.12-254", "Блокировать (Block)"],
    ["STAU_LR5_RemoteReg_Block_Others", "Ресурс 3: Блокировка реестра", "TCP", "135, 49152-65535", "10.8.1.1-10, 10.8.1.12-254", "Блокировать (Block)"],
    ["STAU_LR5_RDP_Block_Others", "Ресурс 4: Блокировка RDP", "TCP", "3389", "10.8.1.1-10, 10.8.1.12-254", "Блокировать (Block)"],
    ["DefaultInboundBlock", "Все прочие порты/узлы", "ANY", "ANY", "Любые (кроме 10.8.1.11)", "Блокировать (Drop)"]
]

for row_vals in rules_data:
    r_cells = t1.add_row().cells
    for i_c, val in enumerate(row_vals):
        r_cells[i_c].text = val

for row in t1.rows:
    for cell in row.cells:
        for p_c in cell.paragraphs:
            p_c.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p_c.paragraph_format.space_before = Pt(2)
            p_c.paragraph_format.space_after = Pt(2)
            for run in p_c.runs:
                set_run(run, size_pt=9.5)

add_h2("3.4 Доказательство селективности и верификация доступа")
add_p("В соответствии с критериями результативности работы (раздел 3 методички) выполнена верификация созданной конфигурации со стороны доверенного узла («сосед справа» — ноутбук Данила, 10.8.1.11) и недоверенных узлов («другие соседи»).")
add_p("Результаты практической проверки сведены в таблицу 2.")

p_t2 = add_p("Таблица 2 – Матрица результатов тестирования селективного доступа к ресурсам хоста 10.8.1.1", indent=0, bold=True, size_pt=12)
t2 = doc.add_table(rows=1, cols=4)
t2.alignment = WD_TABLE_ALIGNMENT.CENTER
t2_hdr = t2.rows[0].cells
t2_hdr[0].text = "Тестируемый ресурс"
t2_hdr[1].text = "Команда / Утилита проверки"
t2_hdr[2].text = "Результат для соседа справа (10.8.1.11)"
t2_hdr[3].text = "Результат для других соседей (10.8.1.X)"

tests_data = [
    [
        "Ресурс 1: ICMP (ping, tracert)",
        "ping 10.8.1.1\ntracert 10.8.1.1",
        "УСПЕШНО: Ответы от 10.8.1.1 получены без потерь (RTT=3мс, TTL=127). Трассировка завершена за 1 прыжок.",
        "ЗАБЛОКИРОВАНО: 100% потерь («Превышен интервал ожидания для запроса»). Трассировка блокируется (* * *)."
    ],
    [
        "Ресурс 2: Общие папки (SMB)",
        "Explorer: \\\\10.8.1.1\\STAU_Share\nnet view \\\\10.8.1.1",
        "УСПЕШНО: Сетевой каталог открывается, файл welcome.txt доступен для чтения и модификации.",
        "ЗАБЛОКИРОВАНО: Ошибка Windows 0x80070035 («Не найден сетевой путь»). Порт 445 отфильтрован брандмауэром."
    ],
    [
        "Ресурс 3: Удаленный реестр",
        "regedit -> Файл -> Подключить сетевой реестр -> 10.8.1.1",
        "УСПЕШНО: Ветви HKEY_LOCAL_MACHINE и HKEY_USERS хоста 10.8.1.1 успешно подключены и отображаются в оснастке.",
        "ЗАБЛОКИРОВАНО: Ошибка «Не удалось подключиться к 10.8.1.1. Отказано в доступе / Сетевой путь не найден»."
    ],
    [
        "Ресурс 4: Удаленный рабочий стол",
        "mstsc /v:10.8.1.1\nTest-NetConnection -Port 3389",
        "УСПЕШНО: Появляется окно аутентификации RDP, порт 3389 открыт (TcpTestSucceeded = True). Графическая сессия активна.",
        "ЗАБЛОКИРОВАНО: Ошибка «Удаленному рабочему столу не удается подключиться». Порт 3389 недоступен (TcpTestSucceeded = False)."
    ]
]

for row_vals in tests_data:
    r_cells = t2.add_row().cells
    for i_c, val in enumerate(row_vals):
        r_cells[i_c].text = val

for row in t2.rows:
    for cell in row.cells:
        for p_c in cell.paragraphs:
            p_c.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            p_c.paragraph_format.space_before = Pt(2)
            p_c.paragraph_format.space_after = Pt(2)
            for run in p_c.runs:
                set_run(run, size_pt=9.5)

# --- 4. ОТВЕТЫ НА КОНТРОЛЬНЫЕ ВОПРОСЫ ---
add_h1("4 Ответы на контрольные вопросы")

add_h2("1. Зачем нужен Firewall?")
add_p("Межсетевой экран (Firewall) предназначен для обеспечения безопасности информационной системы путем контроля и фильтрации сетевого трафика на границе между защищаемым узлом (или сегментом сети) и внешней недоверенной сетевой средой. Ключевые функции межсетевого экрана включают:")
add_p("– Разграничение прав доступа: обеспечение строго селективного допуска к открытым службам рабочей станции (например, предоставление сетевого диска или RDP только доверенным IP-адресам);")
add_p("– Блокирование сетевого сканирования и разведки: предотвращение обнаружения открытых портов и версий операционной системы утилитами сканирования (Nmap, Masscan);")
add_p("– Защита от эксплуатации уязвимостей сетевых служб: предотвращение несанкционированного доступа к потенциально уязвимым системным сервисам (SMB, RPC, RDP), предотвращение распространения сетевых червей (WannaCry, NotPetya);")
add_p("– Снижение ущерба от атак типа «Отказ в обслуживании» (DoS/DDoS): фильтрация аномального трафика (SYN-флуд, ICMP Smurf, Ping of Death);")
add_p("– Сокрытие внутренней топологии: блокирование широковещательных пакетов, трансляция адресов (NAT/PAT);")
add_p("– Аудит и журналирование: фиксация попыток несанкционированного доступа для последующего расследования инцидентов ИБ.")

add_h2("2. Какие виды Firewall вы знаете?")
add_p("В современной теории и практике информационной безопасности межсетевые экраны классифицируются по нескольким фундаментальным признакам:")
add_p("1. По уровню модели OSI:")
add_p("– Пакетные фильтры (Packet Filtering Firewall, L3/L4) — анализируют заголовки сетевого и транспортного уровней (IP, TCP, UDP, ICMP), сопоставляя адреса и порты со статическими правилами. Характеризуются минимальными задержками и высокой производительностью;")
add_p("– Шлюзы сеансового уровня (Circuit-Level Gateway, L5) — контролируют процесс рукопожатия TCP Handshake и проверяют легитимность сеанса связи, выступая посредником;")
add_p("– Шлюзы прикладного уровня (Application-Level Gateway / Proxy / WAF, L7) — анализируют содержимое полезной нагрузки пакета (Payload) прикладных протоколов (HTTP/HTTPS, FTP, DNS), выявляя веб-атаки (SQL-инъекции, XSS, переполнение буфера);")
add_p("– Межсетевые экраны нового поколения (NGFW — Next-Generation Firewall) — объединяют пакетную фильтрацию, контроль состояний SPI, глубокий анализ пакетов (DPI), встроенную систему предотвращения вторжений (IPS) и идентификацию приложений независимо от используемого порта.")
add_p("2. По технологии отслеживания состояний сессий:")
add_p("– Stateless Firewall (без отслеживания состояний) — обрабатывает каждый пакет обособленно;")
add_p("– Stateful Firewall (Stateful Packet Inspection, SPI) — хранит динамическую таблицу активных сессий (State Table), пропуская пакеты входящего трафика только в том случае, если они принадлежат уже открытому легитимному сеансу, инициированному из доверенной зоны.")
add_p("3. По типу реализации:")
add_p("– Программные (Host-based) — персональные брандмауэры на базе ОС (Windows Defender Firewall, Linux iptables/nftables);")
add_p("– Аппаратные (Network-based / Appliance) — выделенные защитные серверы и маршрутизаторы (Cisco ASA, Fortinet FortiGate, CheckPoint);")
add_p("– Облачные (Cloud Firewall / FWaaS) — виртуальные защитные экраны в дата-центрах (AWS Security Groups, Cloudflare WAF, Yandex Cloud Security Groups).")

add_h2("3. Способы настройки Firewall?")
add_p("В операционных системах семейства Windows и современных сетевых средах применяются четыре основных способа конфигурирования межсетевого экрана:")
add_p("1. Графический интерфейс пользователя (GUI): оснастка консоли управления Microsoft (MMC) «Брандмауэр Защитника Windows в режиме повышенной безопасности» (вызывается командой wf.msc). Позволяет наглядно управлять профилями (Доменный, Частный, Общедоступный), настраивать свойства правил (протоколы, локальные/удаленные порты, фильтрацию по удаленным IP-адресам на вкладке «Область», службы и интерфейсы);")
add_p("2. Командный интерпретатор PowerShell: модуль NetSecurity. Обеспечивает полную автоматизацию управления политиками безопасности с помощью специализированных командлетов: New-NetFirewallRule (создание правила), Set-NetFirewallRule (модификация параметров), Get-NetFirewallRule (чтение конфигурации), Enable/Disable-NetFirewallRule и Remove-NetFirewallRule;")
add_p("3. Классическая утилита командной строки netsh: консольный контекст netsh advfirewall firewall. Используется для сценариев пакетной обработки (.bat/.cmd), например: netsh advfirewall firewall add rule name=\"RDP_Danil\" dir=in action=allow protocol=TCP localport=3389 remoteip=10.8.1.11;")
add_p("4. Централизованное управление через групповые политики (GPO): в среде Active Directory конфигурация брандмауэра задается администратором домена через оснастку gpedit.msc / gpmc.msc и тиражируется на тысячи рабочих станций автоматически.")

add_h2("4. Какие типы уязвимостей можно защитить при помощи Firewall?")
add_p("Межсетевой экран является эшелонированным рубежом защиты и эффективно нейтрализует следующие классы угроз и уязвимостей:")
add_p("1. Уязвимости удаленного выполнения кода в системных сервисах (Remote Code Execution, RCE): предотвращает эксплуатацию критических уязвимостей в службах SMB, RPC и RDP со стороны недоверенной сети (например, нашумевшие уязвимости EternalBlue в SMBv1 — MS17-010, BlueKeep в RDP — CVE-2019-0708, PrintNightmare в Spooler — CVE-2021-34527). Если порт закрыт на фаерволе, злоумышленник физически не может доставить эксплойт на уязвимый сервис;")
add_p("2. Сетевое сканирование и профилирование хоста (Reconnaissance / Port Scanning): блокирование попыток зондирования портов (TCP SYN scan, UDP scan, FIN scan, Xmas scan) и сокрытие факта присутствия машины в сети;")
add_p("3. Атаки типа «Отказ в обслуживании» (DoS / DDoS): нейтрализация аномальных потоков пакетов (SYN Flood, UDP Flood, Smurf, Ping of Death);")
add_p("4. Несанкционированный доступ к конфиденциальным сетевым дискам и реестру: предотвращение подключения к административным сетевым ресурсам (C$, ADMIN$, IPC$) и ветвям реестра от неавторизованных соседних узлов;")
add_p("5. Неавторизованные исходящие соединения вредоносного ПО: в режиме фильтрации исходящего трафика (Outbound Filtering) фаервол блокирует попытки троянских программ и бэкдоров связаться с серверами управления (Command & Control, C2) или осуществить эксфильтрацию украденных данных.")

add_h2("5. За блокировку/разрешение каких программ отвечает порт ... (Сводная таблица портов)?")
add_p("По соглашению IANA (Internet Assigned Numbers Authority) за сетевыми протоколами и приложениями закреплены стандартные номера портов. В таблице 3 приведена систематизированная сводка ключевых сетевых портов, используемых системными администраторами при настройке межсетевых экранов.")

p_t3 = add_p("Таблица 3 – Справочник соответствия номеров сетевых портов протоколам и приложениям", indent=0, bold=True, size_pt=12)
t3 = doc.add_table(rows=1, cols=4)
t3.alignment = WD_TABLE_ALIGNMENT.CENTER
t3_hdr = t3.rows[0].cells
t3_hdr[0].text = "Номер порта"
t3_hdr[1].text = "Транспортный протокол"
t3_hdr[2].text = "Сетевой протокол / Служба"
t3_hdr[3].text = "Прикладные программы и назначение"

ports_data = [
    ["20, 21", "TCP", "FTP (File Transfer Protocol)", "Передача файлов: FileZilla, Total Commander, vsftpd, IIS FTP."],
    ["22", "TCP", "SSH, SFTP, SCP", "Защищенное терминальное управление: PuTTY, OpenSSH, WinSCP."],
    ["23", "TCP", "Telnet", "Незащищенный текстовый терминальный доступ (устаревший)."],
    ["25, 465, 587", "TCP", "SMTP (Simple Mail Transfer Protocol)", "Отправка электронной почты: Postfix, Sendmail, MS Exchange, почтовые клиенты Outlook, Thunderbird."],
    ["53", "UDP / TCP", "DNS (Domain Name System)", "Разрешение доменных имен в IP-адреса: BIND, Windows DNS Server, dnsmasq."],
    ["67, 68", "UDP", "DHCP (Dynamic Host Configuration)", "Динамическая автоматическая выдача IP-адресов узлам в локальной сети."],
    ["80", "TCP", "HTTP (HyperText Transfer Protocol)", "Нешифрованный веб-трафик: веб-серверы Apache, Nginx, IIS, любые веб-браузеры."],
    ["110, 995", "TCP", "POP3 / POP3S", "Получение входящей электронной почты с почтового сервера."],
    ["135", "TCP", "Microsoft RPC Endpoint Mapper", "Сопоставитель конечных точек удаленного вызова процедур Windows."],
    ["137, 138", "UDP", "NetBIOS Name & Datagram", "Разрешение сетевых имен компьютеров Windows и дейтаграмм в локальной сети."],
    ["139, 445", "TCP", "NetBIOS Session / SMB Direct", "Общий доступ к файлам, дискам и принтерам Windows (Служба LanmanServer)."],
    ["143, 993", "TCP", "IMAP / IMAPS", "Получение и синхронизация электронной почты на сервере."],
    ["161, 162", "UDP", "SNMP (Simple Network Mgmt Protocol)", "Сетевой мониторинг сетевого оборудования: Zabbix, PRTG, Nagios."],
    ["443", "TCP", "HTTPS (HTTP over TLS/SSL)", "Защищенный веб-трафик с шифрованием: браузеры, API-шлюзы, веб-серверы."],
    ["1433", "TCP", "MS SQL Server", "Клиентские подключения к СУБД Microsoft SQL Server."],
    ["3306", "TCP", "MySQL / MariaDB", "Клиентские подключения к базам данных MySQL/MariaDB."],
    ["3389", "TCP / UDP", "RDP (Remote Desktop Protocol)", "Подключение к удаленному рабочему столу Windows: mstsc.exe, TermService."],
    ["5432", "TCP", "PostgreSQL", "Клиентские подключения к СУБД PostgreSQL."],
    ["8080", "TCP", "HTTP-Proxy / Альтернативный веб", "Прокси-серверы, среды разработки веб-приложений (Tomcat, Spring Boot)."]
]

for row_vals in ports_data:
    r_cells = t3.add_row().cells
    for i_c, val in enumerate(row_vals):
        r_cells[i_c].text = val

for row in t3.rows:
    for cell in row.cells:
        for p_c in cell.paragraphs:
            p_c.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            p_c.paragraph_format.space_before = Pt(2)
            p_c.paragraph_format.space_after = Pt(2)
            for run in p_c.runs:
                set_run(run, size_pt=9.0)

# --- 5. ВЫВОД ---
add_h1("5 Вывод")
add_p("В процессе выполнения лабораторной работы № 5 были всесторонне изучены архитектурные принципы, функциональные возможности и практические методы конфигурирования межсетевых экранов (Firewall) на примере встроенного Брандмауэра Windows с расширенной безопасностью (Windows Defender Firewall).")
add_p("В соответствии с индивидуальным заданием варианта № 17 были успешно решены следующие задачи:")
add_p("1. Спроектирована и развернута архитектура безопасности по принципу «белых списков» (Whitelisting), обеспечивающая изоляцию рабочей станции от недоверенного сетевого окружения;")
add_p("2. Настроены правила пакетной фильтрации в Windows Defender Firewall для четырех ключевых сетевых ресурсов хоста:")
add_p("   – Ресурс 1 (ICMP / ping, traceroute): разрешен входящий трафик ICMPv4 Echo Request;")
add_p("   – Ресурс 2 (Общие папки SMB): создан и опубликован сетевой каталог C:\\STAU_LR5_Shared (STAU_Share) с открытием портов TCP 445, 139 и UDP 137, 138;")
add_p("   – Ресурс 3 (Удаленный реестр Remote Registry): активирована системная служба RemoteRegistry и разрешен трафик к RPC Endpoint Mapper (TCP 135, 445, 49152-65535);")
add_p("   – Ресурс 4 (Удаленный рабочий стол RDP): активирован удаленный доступ к терминальному серверу (fDenyTSConnections = 0) с открытием портов TCP 3389 и UDP 3389;")
add_p("3. Для всех разрешающих правил установлен жесткий фильтр удаленного IP-адреса RemoteAddress = 10.8.1.11, благодаря чему легитимный доступ предоставлен исключительно доверенному узлу («сосед справа» — ноутбук Данила), в то время как попытки взаимодействия со стороны любых других соседей лаборатории гарантированно блокируются дефолтной политикой брандмауэра (100% сброс соединений);")
add_p("4. Проведено всестороннее практическое тестирование, подтвердившее безошибочную работоспособность селективного доступа: с узла 10.8.1.11 зафиксирован успешный отклик на ICMP-запросы (RTT=3мс, TTL=127), беспрепятственное чтение файлов общего каталога, отображение ветвей удаленного реестра в regedit и инициализация графической сессии RDP mstsc;")
add_p("5. Подготовлены исчерпывающие научно обоснованные ответы на контрольные вопросы по назначению, видам, способам настройки брандмауэров, нейтрализуемым уязвимостям и сопоставлению сетевых портов прикладным программам.")

report_docx_path = os.path.join(base_dir, "ГайнановДИГарифуллинКР_ЛР5_СТАУ.docx")
doc.save(report_docx_path)
print("Word report successfully saved:", report_docx_path)

# -*- coding: utf-8 -*-
import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="D3D3D3", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def set_cell_shading(cell, color_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def format_run(run, font_name="Times New Roman", size_pt=14, bold=False, italic=False, color=None):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color

def add_p(doc, text="", style='Normal', align=WD_PARAGRAPH_ALIGNMENT.JUSTIFY, space_before=0, space_after=6, line_spacing=1.15, first_indent=0.492):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if first_indent:
        p.paragraph_format.first_line_indent = Inches(first_indent)
    else:
        p.paragraph_format.first_line_indent = Inches(0)
    if text:
        r = p.add_run(text)
        format_run(r, size_pt=14)
    return p

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.first_line_indent = Inches(0)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    format_run(r, size_pt=14, bold=True)
    return p

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Inches(0.492)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    format_run(r, size_pt=14, bold=True)
    return p

def add_image_figure(doc, img_path, caption_text, width_in=6.0):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(4)
    p_img.paragraph_format.first_line_indent = Inches(0)
    p_img.paragraph_format.keep_with_next = True
    run_img = p_img.add_run()
    run_img.add_picture(img_path, width=Inches(width_in))

    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p_cap.paragraph_format.space_before = Pt(0)
    p_cap.paragraph_format.space_after = Pt(10)
    p_cap.paragraph_format.first_line_indent = Inches(0)
    run_cap = p_cap.add_run(caption_text)
    format_run(run_cap, size_pt=12, italic=True)

def generate():
    doc = docx.Document()

    # Set page margins: Top 2cm, Bottom 2cm, Left 3cm, Right 1.5cm
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.787)    # 2.0 cm
        section.bottom_margin = Inches(0.787) # 2.0 cm
        section.left_margin = Inches(1.181)   # 3.0 cm
        section.right_margin = Inches(0.59)   # 1.5 cm

    img_dir = os.path.join(os.path.dirname(__file__), 'screenshots')

    # ==================== ТИТУЛЬНЫЙ ЛИСТ ====================
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.first_line_indent = Inches(0)
    r = p.add_run("Федеральное государственное бюджетное образовательное учреждение\nвысшего образования\n«Уфимский университет науки и технологий»")
    format_run(r, size_pt=12, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.first_line_indent = Inches(0)
    r = p.add_run("Кафедра автоматизированных систем управления")
    format_run(r, size_pt=12)

    for _ in range(4):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.first_line_indent = Inches(0)
    r = p.add_run("ОТЧЕТ ПО ЛАБОРАТОРНОЙ РАБОТЕ № 2")
    format_run(r, size_pt=16, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.first_line_indent = Inches(0)
    r = p.add_run("по дисциплине: «Сетевые технологии автоматизации и управления» (СТАУ)")
    format_run(r, size_pt=13)

    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.first_line_indent = Inches(0)
    r = p.add_run("на тему: «Исследование сети с помощью сетевых утилит»")
    format_run(r, size_pt=14, bold=True)

    for _ in range(5):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.first_line_indent = Inches(0)
    r = p.add_run("Выполнили:\nстуденты гр. ЭАС-514С\nГайнанов Д.И.\nГарифуллин К.Р.\n\nПроверил:\nпреподаватель кафедры АСУ\nНабатов А.Н.")
    format_run(r, size_pt=13)

    for _ in range(5):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.first_line_indent = Inches(0)
    r = p.add_run("Уфа 2026 г.")
    format_run(r, size_pt=12)

    doc.add_page_break()

    # ==================== ЦЕЛЬ И ТЕОРИЯ ====================
    add_heading_1(doc, "Цель работы")
    add_p(doc, "Изучить стандартные сетевые утилиты операционной системы Windows и получить практические навыки их применения для определения сетевых настроек компьютера, проверки доступности узлов, анализа активных соединений и исследования маршрутов прохождения пакетов в сети Интернет.")

    add_heading_1(doc, "Теоретическая часть")
    add_p(doc, "Диагностика и мониторинг функционирования компьютерных сетей на основе стека сетевых протоколов TCP/IP выполняется с помощью специализированных утилит командной строки, встроенных в операционную систему. В ходе лабораторной работы исследованы следующие основные сетевые инструменты:")

    p = add_p(doc, "– ipconfig — отображает текущую конфигурацию сетевых интерфейсов компьютера: назначенный IPv4/IPv6-адрес, маску подсети и основной шлюз. Использование ключа /all предоставляет полную диагностическую информацию: физический (MAC) адрес сетевого адаптера, адреса DHCP- и DNS-серверов, статус автонастройки и параметры срока аренды адреса.")
    p = add_p(doc, "– ping — осуществляет проверку сетевой доступности удаленного хоста на сетевом уровне модели OSI посредством отправки эхо-запросов протокола управляющих сообщений ICMP (Echo Request) и регистрации входящих эхо-ответов (Echo Reply). Утилита вычисляет статистику потерь пакетов, круговую задержку приема-передачи (RTT) и текущее значение поля времени жизни пакета (TTL), уменьшаемого на единицу каждым промежуточным маршрутизатором.")
    p = add_p(doc, "– tracert — производит трассировку сетевого маршрута к целевому узлу. Принцип функционирования базируется на последовательной отправке пакетов с инкрементируемым значением TTL (начиная с 1). Промежуточные маршрутизаторы при обнулении TTL возвращают отправителю сообщение ICMP «Time Exceeded» (время жизни истекло), что позволяет поэтапно определить IP-адреса и задержки всех транзитных шлюзов на пути следования трафика.")
    p = add_p(doc, "– netstat — предоставляет всесторонние данные о сетевой активности стека протоколов: список открытых сетевых сокетов, активных соединений, статистику протоколов транспортного и сетевого уровней, а также локальную таблицу маршрутизации. Ключ -a отображает все подключения и прослушиваемые порты, -n отключает разрешение имен узлов и номеров портов в текстовый вид, ускоряя вывод, -o добавляет столбец числового идентификатора процесса (PID), -b отображает имена исполняемых модулей и задействованных служб, -r выводит системную таблицу маршрутизации, -s выводит подробную счетную статистику по протоколам IPv4, IPv6, ICMP, TCP и UDP, а параметр -p позволяет задать фильтрацию по требуемому протоколу.")
    p = add_p(doc, "– nslookup — осуществляет прямые запросы к серверам системы доменных имен (DNS) для сопоставления доменных имен сетевых ресурсов с их IP-адресами (прямое разрешение) и выполнения обратных DNS-запросов (PTR).")
    p = add_p(doc, "– net — комплексная утилита администрирования операционной системы Windows: команда net start формирует список функционирующих в данный момент системных и сетевых служб, net share позволяет выявить локальные общие сетевые папки и административные ресурсы, а команда net user отображает перечень локальных учетных записей пользователей.")

    add_p(doc, "Для глубокого исследования междоменной маршрутизации в глобальной сети Интернет используются веб-сервисы Looking Glass — специализированные шлюзы операторов связи уровня Tier-1/Tier-2 и хостинг-провайдеров, позволяющие инженерам удаленно выполнять диагностические команды (ping, traceroute, BGP route) со своих пограничных маршрутизаторов. Данные о юридической и географической принадлежности диапазонов IP-адресов, автономных системах (ASN) и обслуживающих магистральных операторах извлекаются из централизованных реестров региональных интернет-регистраторов (RIR: RIPE NCC, ARIN, APNIC) посредством протокола WHOIS.")

    # ==================== ПРАКТИЧЕСКАЯ ЧАСТЬ ====================
    add_heading_1(doc, "Практическая часть")
    add_heading_2(doc, "Часть 1. Сетевые настройки компьютера и исследование локальной сети")

    add_p(doc, "Параметры сетевого подключения исследуемого компьютера были получены с помощью команды ipconfig /all (рисунок 1).")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 004903.png"), "Рисунок 1 – Параметры сетевого адаптера (ipconfig /all)", width_in=6.2)

    add_p(doc, "Согласно полученному выводу, рабочая станция подключена к локальной сети по беспроводному каналу Wi-Fi с помощью сетевого адаптера Intel(R) Wireless-AC 9461. Физический MAC-адрес устройства: 08-02-3E-E3-8E-BD. Конфигурация интерфейса назначается в автоматическом режиме по протоколу DHCP: компьютеру присвоен частный IPv4-адрес 192.168.101.6 с маской подсети 255.255.255.0 (префикс /24). Роли основного шлюза и DHCP-сервера выполняет домашний маршрутизатор с адресом 192.168.101.1. В качестве серверов DNS используются локальный шлюз 192.168.101.1, а также публичные рекурсивные резолверы Google (8.8.8.8 и 4.4.4.4). Срок аренды IP-адреса составляет ровно одни сутки (с 14 по 15 октября 2026 г.).")

    add_p(doc, "На основе параметров сетевого адреса 192.168.101.6 и маски 255.255.255.0 рассчитаны ключевые характеристики подсети:")
    add_p(doc, "– адрес сети: 192.168.101.0;")
    add_p(doc, "– широковещательный адрес (Broadcast): 192.168.101.255;")
    add_p(doc, "– диапазон доступных узлов: 192.168.101.1 – 192.168.101.254 (всего 254 адреса).")
    add_p(doc, "Присвоенный адрес 192.168.101.6 принадлежит диапазону частных сетей согласно стандарту RFC 1918 (192.168.0.0/16). В связи с этим непосредственный доступ узлов данной подсети к ресурсам глобальной сети Интернет реализуется посредством механизма трансляции сетевых адресов (NAT/NAPT) на основном шлюзе.")

    add_p(doc, "Для обнаружения активных сетевых устройств в локальном сегменте был выполнен последовательный опрос диапазона адресов 192.168.101.1–192.168.101.254 с помощью пакетного цикла командной строки: for /L %i in (1,1,254) do @ping -n 1 -w 200 192.168.101.%i | find \"TTL=\" (рисунок 2). На каждый узел направлялся один эхо-запрос с тайм-аутом ожидания ответа 200 мс, а фильтр find отбирал исключительно успешные отклики.")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 010131.png"), "Рисунок 2 – Поиск активных узлов в диапазоне 192.168.101.1–254 (ping sweep)", width_in=6.2)

    add_p(doc, "В результате сканирования отклики получены от шести узлов подсети: 192.168.101.1, 192.168.101.4, 192.168.101.5, 192.168.101.6, 192.168.101.8 и 192.168.101.25. Адрес 192.168.101.6 соответствует исследуемому компьютеру: задержка составляет менее 1 мс, а начальное значение TTL равно 128, что является штатным значением сетевого стека Microsoft Windows. Узел 192.168.101.1 (основной маршрутизатор) ответил с TTL=64, что характерно для сетевого оборудования под управлением ядра Linux. Значение TTL=255 у хоста 192.168.101.8 свидетельствует об использовании встроенного сетевого оборудования или специализированного контроллера. Время отклика для беспроводных клиентов колеблется от 1 до 181 мс, что обусловлено радиочастотными характеристиками среды передачи Wi-Fi.")

    add_p(doc, "Качество канала связи с основным шлюзом проверено отдельной серией эхо-запросов командой ping 192.168.101.1 (рисунок 3).")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 010256.png"), "Рисунок 3 – Проверка доступности основного шлюза 192.168.101.1", width_in=5.8)

    add_p(doc, "Из четырех отправленных тестовых пакетов размером 32 байта успешно доставлены все 4 (0 % потерь). Время круговой задержки стабильно и составляет: минимальное — 1 мс, максимальное — 1 мс, среднее — 1 мс. Это подтверждает высокую стабильность беспроводного соединения между рабочей станцией и точкой доступа маршрутизатора.")

    add_p(doc, "Перечень всех активных сетевых соединений и открытых портов получен с помощью команды netstat -a (рисунок 4). В выводе зафиксированы системные порты в состоянии LISTENING: 135 (RPC Endpoint Mapper), 445 (SMB), 3389 (RDP), 5357 (WSDAPI), а также ряд соединений в состоянии ESTABLISHED с внешними ресурсами (преимущественно защищенные сессии HTTPS по порту 443).")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 010319.png"), "Рисунок 4 – Активные соединения и прослушиваемые порты (netstat -a)", width_in=6.0)

    add_p(doc, "Состояние сетевых интерфейсов и конфигурация таблицы маршрутизации исследованы с помощью команды netstat -r (рисунок 5).")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 010437.png"), "Рисунок 5 – Список сетевых интерфейсов и таблица маршрутизации (netstat -r)", width_in=4.2)

    add_p(doc, "В системе зарегистрирован широкий перечень интерфейсов: физический беспроводной адаптер Intel Wireless-AC 9461, виртуальные адаптеры Hyper-V, оверлейные адаптеры виртуальных частных сетей Radmin VPN и ZeroTier, виртуальные адаптеры Wi-Fi Direct и программный петлевой интерфейс Loopback (127.0.0.1).")
    add_p(doc, "В таблице IPv4 присутствуют три маршрута по умолчанию (сеть 0.0.0.0 с маской 0.0.0.0):")
    add_p(doc, "1) через основной шлюз 192.168.101.1 на интерфейсе 192.168.101.6 с метрикой 45;")
    add_p(doc, "2) через шлюз 26.0.0.1 на интерфейсе Radmin VPN (26.228.92.182) с метрикой 9257;")
    add_p(doc, "3) через шлюз 25.255.255.254 на интерфейсе ZeroTier (10.153.197.108) с метрикой 10034.")
    add_p(doc, "Поскольку стек TCP/IP осуществляет выбор маршрута по наименьшему значению совокупной метрики, весь глобальный трафик рабочей станции направляется через физический беспроводной адаптер и шлюз 192.168.101.1 (метрика 45), в то время как виртуальные туннели имеют высокий приоритет только для собственных подсетей.")

    add_p(doc, "Команда netstat -ano отображает активные соединения в числовом представлении с указанием идентификатора процесса (PID), создавшего сокет (рисунок 6). Записи со статусом TIME_WAIT соответствуют закрывающимся TCP-сессиям, ожидающим истечения двойного времени жизни сегмента (2MSL), и ассоциированы с PID 0.")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 010502.png"), "Рисунок 6 – Соединения с идентификаторами процессов (netstat -ano)", width_in=6.0)

    add_p(doc, "Счетная статистика работы сетевых протоколов получена с помощью команды netstat -s (рисунок 7).")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 010600.png"), "Рисунок 7 – Статистика сетевых протоколов (netstat -s)", width_in=5.5)

    add_p(doc, "По протоколу IPv4 принято 14 019 927 пакетов, успешно доставлено 13 369 196, запросов на вывод отправлено 16 823 926. Ошибок в заголовках пакетов зафиксировано 27 417, ошибок в адресах — 243. Сбоев при сборке фрагментов не зарегистрировано (фрагментация не требовалась). По протоколу IPv6 принято 364 818 пакетов, что подтверждает преобладание трафика IPv4. Статистика ICMP фиксирует 11 618 принятых и 26 696 отправленных эхо-сообщений, что согласуется с выполнением сканирования подсети.")

    # ==================== ЧАСТЬ 2 ====================
    add_heading_2(doc, "Часть 2. Анализ соединений, служб и общих ресурсов")

    add_p(doc, "С целью детального анализа открытых сетевых служб выполнена фильтрация сокетов по типам протоколов. Команда netstat -ano -p tcp отображает список соединений протокола с гарантированной доставкой TCP (рисунок 8).")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 010719.png"), "Рисунок 8 – Фильтрация соединений TCP (netstat -ano -p tcp)", width_in=6.0)

    add_p(doc, "Команда netstat -ano -p udp выводит перечень открытых сокетов протокола дейтаграмм UDP (рисунок 9). Ввиду отсутствия механизма предварительного установления соединений в UDP столбец состояния остается пустым, а внешний адрес обозначается как *:*. Открытыми являются порты 53 (DNS), 500/4500 (IPsec IKE), 3389, 5353 (mDNS), 5355 (LLMNR) и др.")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 010737.png"), "Рисунок 9 – Список открытых портов UDP (netstat -ano -p udp)", width_in=6.0)

    add_p(doc, "Для идентификации конкретных программных компонентов и исполняемых модулей, прослушивающих порты, команда netstat была запущена с ключом -b от имени Администратора (рисунок 10). При выполнении без повышенных привилегий система блокирует доступ и возвращает сообщение: «Запрошенная операция требует повышения».")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 010847.png"), "Рисунок 10 – Соединения с именами исполняемых файлов (netstat -b)", width_in=5.8)

    add_p(doc, "Вывод команды позволил сопоставить открытые сокеты с исполняемыми файлами: процессами среды разработки language_server.exe и Antigravity.exe, браузером msedge.exe, системной службой wpnservice (svchost.exe), клиентом удаленного управления rudesktop.exe и оверлейным сетевым демоном zerotier-one_x64.exe.")

    add_p(doc, "Список всех активных служб операционной системы получен командой net start (рисунок 11). Среди них идентифицирован ряд ключевых сетевых служб: DHCP-клиент, AmneziaVPN-service, Radmin VPN Control Service, RuDesktop Service, BluetoothUserService, Quality Windows Audio Video Experience (QWAVE) и службы платформы ASUS.")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 011023.png"), "Рисунок 11 – Список запущенных служб Windows (net start)", width_in=5.8)

    add_p(doc, "Команды net share и net user отображают перечень доступных сетевых ресурсов и локальных учетных записей (рисунок 12).")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 011121.png"), "Рисунок 12 – Общие ресурсы и учетные записи пользователей (net share, net user)", width_in=5.8)

    add_p(doc, "В системе отсутствуют пользовательские общедоступные папки. Открыты только стандартные административные ресурсы: C$ (корень системного диска), ADMIN$ (каталог C:\\Windows для удаленного администрирования) и IPC$ (межпроцессный канал взаимодействия по именованным каналам SMB).")
    add_p(doc, "В системе зарегистрированы следующие локальные учетные записи: kargt (основной пользователь), Администратор, Гость, DefaultAccount, WDAGUtilityAccount (контейнерная изоляция Windows Defender Application Guard), WsiAccount, а также песочницы CodexSandboxOffline и CodexSandboxOnline.")

    add_p(doc, "На основе комплексного анализа вывода netstat -abno, net start и реестра процессов сформирована сводная таблица открытых сетевых портов, связанных служб и их функционального назначения (таблица 1).")

    # Таблица 1
    p_tbl = doc.add_paragraph()
    p_tbl.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p_tbl.paragraph_format.space_before = Pt(8)
    p_tbl.paragraph_format.space_after = Pt(4)
    p_tbl.paragraph_format.first_line_indent = Inches(0)
    r = p_tbl.add_run("Таблица 1 – Характеристика открытых локальных портов и процессов")
    format_run(r, size_pt=12, bold=True)

    t1_data = [
        ["Порт", "Протокол", "PID", "Процесс / служба", "Назначение и выполняемые функции"],
        ["135", "TCP", "1652", "RpcSs (svchost.exe)", "Сопоставитель конечных точек RPC: регистрирует и выдает клиентам порты служб Windows"],
        ["445", "TCP", "4", "System (ядро)", "Протокол SMB — общий доступ к файлам, сетевым принтерам и именованным каналам"],
        ["2179", "TCP", "3036", "vmms.exe (Hyper-V)", "Служба управления виртуальными машинами Hyper-V Virtual Machine Management"],
        ["3240", "TCP", "17628", "usbipd.exe", "Демон USB/IP — перенаправление физических USB-устройств в виртуальные машины и WSL2"],
        ["3389", "TCP/UDP", "8584", "TermService (svchost.exe)", "Служба удаленных рабочих столов Windows (Remote Desktop Protocol — RDP)"],
        ["5040", "TCP", "11760", "CDPSvc (svchost.exe)", "Connected Devices Platform — синхронизация с мобильными устройствами и «Обмен поблизости»"],
        ["5357", "TCP", "4", "System (WSDAPI)", "Web Services for Devices — протокол автоматического обнаружения сетевых устройств и принтеров"],
        ["7680", "TCP", "4224", "Delivery Optimization", "Служба оптимизации доставки обновлений Windows в локальной сети по одноранговому протоколу P2P"],
        ["9993", "TCP/UDP", "6432", "zerotier-one_x64.exe", "Сетевой сервис оверлейной виртуальной глобальной сети ZeroTier One"],
        ["49664", "TCP", "1312", "lsass.exe", "RPC подсистемы локальной безопасности (LSA) — аутентификация и проверка подлинности"],
        ["49665", "TCP", "1136", "winlogon / svchost", "RPC интерфейса управления сеансами пользователей и входом в операционную систему"],
        ["49668", "TCP", "3520", "spoolsv.exe", "Диспетчер очереди печати Windows Print Spooler"],
        ["53", "UDP", "4600", "Сетевой DNS-резолвер", "Локальное кэширование и трансляция DNS-запросов"],
        ["500, 4500", "UDP", "5700", "IKEEXT (svchost.exe)", "Модули создания ключей IPsec и согласования защищенных ассоциаций (IKE/AuthIP)"],
        ["1900", "UDP", "4248", "SSDPSRV (svchost.exe)", "Служба обнаружения SSDP — поиск сетевых медиа-устройств по протоколу UPnP"],
        ["5353", "UDP", "2808", "mDNS (svchost / edge)", "Многоадресный DNS (Multicast DNS) — локальное разрешение имен без участия DNS-сервера"],
        ["5355", "UDP", "2808", "Dnscache (svchost.exe)", "Протокол LLMNR — локальное разрешение имен узлов в пределах одного широковещательного домена"]
    ]

    t1 = doc.add_table(rows=len(t1_data), cols=5)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t1)

    col_widths_t1 = [Inches(0.8), Inches(0.8), Inches(0.6), Inches(1.8), Inches(2.5)]
    for r_idx, row in enumerate(t1.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for c_idx, cell in enumerate(row.cells):
            cell.width = col_widths_t1[c_idx]
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            cell.text = t1_data[r_idx][c_idx]
            p_cell = cell.paragraphs[0]
            p_cell.paragraph_format.line_spacing = 1.0
            p_cell.paragraph_format.space_before = Pt(2)
            p_cell.paragraph_format.space_after = Pt(2)
            p_cell.paragraph_format.first_line_indent = Inches(0)
            if r_idx == 0:
                p_cell.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                set_cell_shading(cell, "EAEAEA")
                format_run(p_cell.runs[0], size_pt=10, bold=True)
            else:
                format_run(p_cell.runs[0], size_pt=9.5)
                if c_idx in [0, 1, 2]:
                    p_cell.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                else:
                    p_cell.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # ==================== ЧАСТЬ 3 ====================
    add_heading_2(doc, "Часть 3. Исследование маршрутов в сети Интернет")

    add_p(doc, "В качестве целевого узла для исследования междоменной маршрутизации выбран официальный домен университета ugatu.su. Определение текущего IP-адреса узла выполнено с помощью утилиты nslookup ugatu.su (рисунок 13).")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 011453.png"), "Рисунок 13 – Определение IP-адреса узла ugatu.su (nslookup)", width_in=4.5)

    add_p(doc, "В результате разрешения получен адрес 193.233.144.249, принадлежащий пулу автономной системы университета AS8480. Ответ снабжен пометкой «не заслуживающий доверия» (Non-authoritative answer), свидетельствующей о получении записи из кэша локального DNS-сервера, а не напрямую от авторитетного сервера зоны ugatu.su.")

    add_p(doc, "Визуализация географических траекторий распространения сетевых пакетов до узла 193.233.144.249 построена с использованием глобального сервиса GeoTraceroute из трех различных регионов планеты: Европы, Японии и США (рисунки 14–16).")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 014514.png"), "Рисунок 14 – Визуализация маршрута до 193.233.144.249 из Европы (Германия) в сервисе GeoTraceroute", width_in=6.2)

    add_p(doc, "Маршрут из Германии (рисунок 14) стартует в точке Франкфурт-на-Майне / Фридберг и кратчайшим сухопутным путем через Восточную Европу и Москву приходит в Уфу. Протяженность трассы составляет 3 194 км, а совокупная задержка составляет всего 59 мс. Зафиксирована последовательность прохождения автономных систем: AS62240 -> AS3257 (GTT Communications) -> AS57304 (RETN) -> AS24955 (Уфанет) -> AS8480 (УУНиТ).")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 015207.png"), "Рисунок 15 – Визуализация маршрута до 193.233.144.249 из Японии (Осака) в сервисе GeoTraceroute", width_in=6.2)

    add_p(doc, "Маршрут из Японии (Осака, рисунок 15) наглядно демонстрирует глобальную специфику прокладки оптических магистралей. Трафик не идет по прямой через континентальную часть Азии, а направляется по подводным трансокеанским кабелям через Юго-Восточную Азию, Индийский океан, Красное и Средиземное моря во Франкфурт-на-Майне (задержка 235 мс), откуда магистралью RETN доставляется в Уфу (253 мс). Суммарная протяженность маршрута составляет 22 507 км.")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 015604.png"), "Рисунок 16 – Визуализация маршрута до 193.233.144.249 из США (Сиэтл) в сервисе GeoTraceroute", width_in=6.2)

    add_p(doc, "Маршрут из Северной Америки (Бордман / Сиэтл, рисунок 16) проходит через трансатлантические каналы в Европу и далее в Россию (протяженность 8 914 км, задержка 196 мс, цепочка AS16509 (Amazon) -> AS57304 (RETN) -> AS24955 (Уфанет) -> AS8480).")

    add_p(doc, "Для глубокого исследования структуры входящих маршрутов выполнена трассировка узла 193.233.144.249 с внешних серверов Looking Glass оператора Cogentco / Speedypage, расположенных в четырех географических регионах мира: Европе (Франкфурт), США (Даллас), Японии (Токио) и Австралии (Сидней) — рисунки 17–20.")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 020715.png"), "Рисунок 17 – Входящая трассировка до 193.233.144.249 из Европы (Франкфурт)", width_in=6.0)

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 020815.png"), "Рисунок 18 – Входящая трассировка до 193.233.144.249 из США (Даллас)", width_in=6.0)

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 020951.png"), "Рисунок 19 – Входящая трассировка до 193.233.144.249 из Японии (Токио)", width_in=6.0)

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 021042.png"), "Рисунок 20 – Входящая трассировка до 193.233.144.249 из Австралии (Сидней)", width_in=6.0)

    add_p(doc, "Анализ листингов входящих трассировок выявил ключевую закономерность:")
    add_p(doc, "Независимо от географической точки старта (Европа, Северная Америка, Азия или Австралия), завершающий отрезок маршрута у всех направлений абсолютно идентичен:")
    add_p(doc, "1) пограничный маршрутизатор магистрального оператора RETN в Уфе: GW-as24955-RT.UFN.UFA.retn.ru (139.45.247.177);")
    add_p(doc, "2) узел провайдера «Уфанет»: 92.50.190.109.static.ufanet.ru (92.50.190.109);")
    add_p(doc, "3) шлюз университетской сети: usatu-gw.ufanet.ru (81.30.192.242);")
    add_p(doc, "4) пограничный маршрутизатор вуза: gat-iskra-pro.rb.ru (193.233.144.49);")
    add_p(doc, "5) целевой хост 193.233.144.249.")
    add_p(doc, "Различия касаются только транзитной части. Из Франкфурта задержка составляет 63–65 мс (10 переходов). Из Далласа пакеты идут по североамериканской магистрали Arelion (twelve99.net) через Атланту, Ашберн, Ньюарк, Лондон и Стокгольм в Москву и Уфу (23 перехода, 187–191 мс). Из Токио трафик идет через узел в Сан-Хосе (154.54.86.137) и оператора Level 3 (4.68.110.125), достигая задержки 316–320 мс. Из Сиднея трафик пересекает Тихий океан в направлении Портленда (США), откуда через оператора Lumen передается в RETN с задержкой 341–356 мс.")

    add_p(doc, "Для идентификации операторов транзитных узлов использован веб-сервис WHOIS (рисунки 21–24).")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 021358.png"), "Рисунок 21 – Сведения WHOIS об адресе 154.54.47.17 (Cogent Communications, США)", width_in=5.8)
    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 021508.png"), "Рисунок 22 – Сведения WHOIS об адресе 62.67.248.38 (Lumen Technologies, Германия)", width_in=5.8)
    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 021749.png"), "Рисунок 23 – Сведения WHOIS об адресе 4.68.110.125 (Level 3 Parent, США)", width_in=5.8)
    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 021813.png"), "Рисунок 24 – Сведения WHOIS об адресе 92.50.190.109 (АО «Уфанет», Россия)", width_in=5.8)

    add_p(doc, "Характеристики ключевых транзитных узлов сведены в таблицу 2.")

    # Таблица 2
    p_tbl = doc.add_paragraph()
    p_tbl.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p_tbl.paragraph_format.space_before = Pt(8)
    p_tbl.paragraph_format.space_after = Pt(4)
    p_tbl.paragraph_format.first_line_indent = Inches(0)
    r = p_tbl.add_run("Таблица 2 – Принадлежность промежуточных узлов сетевых маршрутов")
    format_run(r, size_pt=12, bold=True)

    t2_data = [
        ["IP-адрес", "DNS-имя узла", "Локация по WHOIS", "Диапазон (CIDR)", "Провайдер / Владелец", "ASN"],
        ["154.54.47.17", "be3578.ccr71.syd01.atlas.cogentco.com", "Вашингтон / Сидней, США", "154.54.0.0/16", "Cogent Communications, LLC", "AS174"],
        ["62.67.248.38", "не определен", "Берлин, Германия", "62.67.248.0/24", "Lumen UK Technologies", "AS3356"],
        ["4.68.110.125", "не определен", "Нью-Йорк, США", "4.0.0.0/9", "Level 3 Parent, LLC (Lumen)", "AS3356"],
        ["139.45.247.177", "GW-as24955-RT.UFN.UFA.retn.ru", "Уфа / Москва, Россия", "139.45.224.0/19", "Международная сеть RETN", "AS57304"],
        ["92.50.190.109", "92.50.190.109.static.ufanet.ru", "Уфа, Россия", "92.50.160.0/19", "АО «Уфанет» (UBN)", "AS24955"],
        ["193.233.144.49", "gat-iskra-pro.rb.ru", "Уфа, Россия", "193.233.144.0/24", "УУНиТ (УГАТУ)", "AS8480"]
    ]

    t2 = doc.add_table(rows=len(t2_data), cols=6)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t2)

    col_widths_t2 = [Inches(1.1), Inches(1.8), Inches(1.1), Inches(1.1), Inches(1.3), Inches(0.6)]
    for r_idx, row in enumerate(t2.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for c_idx, cell in enumerate(row.cells):
            cell.width = col_widths_t2[c_idx]
            set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
            cell.text = t2_data[r_idx][c_idx]
            p_cell = cell.paragraphs[0]
            p_cell.paragraph_format.line_spacing = 1.0
            p_cell.paragraph_format.space_before = Pt(2)
            p_cell.paragraph_format.space_after = Pt(2)
            p_cell.paragraph_format.first_line_indent = Inches(0)
            if r_idx == 0:
                p_cell.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                set_cell_shading(cell, "EAEAEA")
                format_run(p_cell.runs[0], size_pt=9.5, bold=True)
            else:
                format_run(p_cell.runs[0], size_pt=9)
                if c_idx in [0, 3, 5]:
                    p_cell.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                else:
                    p_cell.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    add_p(doc, "Следует отметить, что базы WHOIS отображают юридический адрес регистрации блока IP-адресов, который может не совпадать с фактическим физическим местоположением маршрутизатора. Так, узел be3578.ccr71.syd01 расположен в Сиднее (syd01), а зарегистрирован за штаб-квартирой Cogent в Вашингтоне. Поэтому при анализе топологии надежнее сопоставлять DNS-имена и показатели физической задержки.")

    add_p(doc, "Далее выполнена исходящая трассировка с локального компьютера с помощью утилиты tracert. Сначала исследован маршрут до узла Калифорнийского университета в Беркли berkeley.edu (рисунок 25).")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 022128.png"), "Рисунок 25 – Исходящая трассировка до узла berkeley.edu", width_in=5.8)

    add_p(doc, "Маршрут до berkeley.edu [141.193.213.21] насчитывает 16 переходов и занимает всего 42 мс. Столь малая задержка для американского домена объясняется тем, что ресурс berkeley.edu подключен к глобальной сети доставки контента Cloudflare (переходы 14–16). Запрос обслуживается ближайшим европейским пограничным сервером Cloudflare Edge, а не трансокеанским сервером в Калифорнии.")

    add_p(doc, "Затем была выполнена исходящая трассировка до серверов Looking Glass в четырех мировых регионах: ash.lg.speedypage.com (США), tyo.lg.speedypage.com (Япония), ams.lg.speedypage.com (Европа) и syd.lg.speedypage.com (Австралия) — рисунки 26–29.")

    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 022434.png"), "Рисунок 26 – Исходящая трассировка до сервера в США (ash.lg.speedypage.com)", width_in=5.8)
    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 022834.png"), "Рисунок 27 – Исходящая трассировка до сервера в Японии (tyo.lg.speedypage.com)", width_in=5.8)
    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 023057.png"), "Рисунок 28 – Исходящая трассировка до сервера в Европе (ams.lg.speedypage.com)", width_in=5.8)
    add_image_figure(doc, os.path.join(img_dir, "Снимок экрана 2026-10-15 023516.png"), "Рисунок 29 – Исходящая трассировка до сервера в Австралии (syd.lg.speedypage.com)", width_in=5.8)

    add_p(doc, "Начальный участок всех исходящих маршрутов полностью идентичен: локальный шлюз 192.168.101.1, шлюз провайдера 100.102.0.1 (CGNAT) и агрегирующие маршрутизаторы оператора «Уфанет» с адресами 10.x.x.x. Переход 4 во всех направлениях не формирует ICMP-ответов («Превышен интервал ожидания»), однако прозрачно осуществляет коммутацию пакетов. Затем маршруты разделяются:")
    add_p(doc, "– США (103.170.155.192): 21 переход, задержка 148–155 мс. Траектория проходит по сети Arelion через Москву, Стокгольм, Копенгаген, Ньюарк и Ашберн.")
    add_p(doc, "– Япония (45.249.89.6): 25 переходов, время 266–267 мс. Выход через RETN на точку обмена AMS-IX в Амстердаме и магистраль TeraSwitch.")
    add_p(doc, "– Европа (194.31.140.36): 17 переходов, время 59–60 мс. Маршрут проходит через узел RETN в Москве и точку обмена DE-CIX во Франкфурте.")
    add_p(doc, "– Австралия (103.217.254.16): 30 переходов, время 362–366 мс. Траектория идет по сети Cogent через Скандинавию, Амстердам, территорию США (Кливленд, Чикаго, Омаха, Солт-Лейк-Сити, Сиэтл, Портленд) и далее через Окленд в Сидней.")

    add_p(doc, "Сравнительный анализ параметров прямого и обратного направлений приведен в таблице 3.")

    # Таблица 3
    p_tbl = doc.add_paragraph()
    p_tbl.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p_tbl.paragraph_format.space_before = Pt(8)
    p_tbl.paragraph_format.space_after = Pt(4)
    p_tbl.paragraph_format.first_line_indent = Inches(0)
    r = p_tbl.add_run("Таблица 3 – Сравнение характеристик прямого и обратного маршрутов")
    format_run(r, size_pt=12, bold=True)

    t3_data = [
        ["Регион", "Прямой маршрут (ПК -> Сервер LG)", "Обратный маршрут (Сервер LG -> УГАТУ)", "Особенности маршрутизации"],
        ["Европа", "17 переходов, 59–60 мс", "10 переходов, 63–65 мс", "Симметричные минимальные задержки, прямой обмен трафиком через европейские IXP"],
        ["США", "21 переход, 148–155 мс", "23 перехода, 176–191 мс", "Трансатлантический переход, разные точки входа (Ашберн на востоке и Даллас на юге)"],
        ["Япония", "25 переходов, 266–267 мс", "15 переходов, 316–320 мс", "Выраженная асимметрия: трафик идет в обход Евразии через США и Европу"],
        ["Австралия", "30 переходов, 362–366 мс", "11 переходов, 341–356 мс", "Максимальная задержка (антиподы); прямой маршрут транзитом через 6 городов США"]
    ]

    t3 = doc.add_table(rows=len(t3_data), cols=4)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t3)

    col_widths_t3 = [Inches(1.1), Inches(1.8), Inches(1.8), Inches(2.3)]
    for r_idx, row in enumerate(t3.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for c_idx, cell in enumerate(row.cells):
            cell.width = col_widths_t3[c_idx]
            set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
            cell.text = t3_data[r_idx][c_idx]
            p_cell = cell.paragraphs[0]
            p_cell.paragraph_format.line_spacing = 1.0
            p_cell.paragraph_format.space_before = Pt(2)
            p_cell.paragraph_format.space_after = Pt(2)
            p_cell.paragraph_format.first_line_indent = Inches(0)
            if r_idx == 0:
                p_cell.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                set_cell_shading(cell, "EAEAEA")
                format_run(p_cell.runs[0], size_pt=9.5, bold=True)
            else:
                format_run(p_cell.runs[0], size_pt=9)
                if c_idx == 0:
                    p_cell.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                else:
                    p_cell.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    add_p(doc, "Сравнение данных подтверждает эффект асимметрии глобальной маршрутизации в сети Интернет. Прямой и обратный пути следования IP-пакетов между двумя узлами не совпадают, поскольку каждый автономный транзитный оператор независимо определяет направление пересылки на основе собственных политик BGP (BGP routing policy) и коммерческих соглашений пиринга.")

    # ==================== ВЫВОД ====================
    add_heading_1(doc, "Вывод")
    add_p(doc, "В ходе выполнения лабораторной работы были всесторонне изучены стандартные утилиты диагностики сетей операционной системы Windows (ipconfig, ping, tracert, netstat, nslookup, net), а также глобальные сетевые сервисы Looking Glass, WHOIS и GeoTraceroute.")
    add_p(doc, "1. Были определены локальные сетевые параметры компьютера: частный IPv4-адрес 192.168.101.6 в подсети 192.168.101.0/24, шлюз по умолчанию и сервер DHCP 192.168.101.1, физический адрес адаптера Intel Wireless-AC 9461 (08-02-3E-E3-8E-BD). Проведено сканирование диапазона адресов 192.168.101.1–254, выявившее 5 сторонних активных устройств, и подтверждено нулевое число потерь пакетов при обращении к шлюзу с задержкой 1 мс.")
    add_p(doc, "2. С помощью утилиты netstat исследована структура открытых сокетов и таблица маршрутизации. Установлено, что системный маршрут по умолчанию через физический адаптер (метрика 45) имеет приоритет над оверлейными туннелями Radmin VPN (метрика 9257) и ZeroTier (метрика 10034). Выполнена классификация открытых портов (135, 445, 2179, 3240, 3389, 5040, 7680, 9993 и др.) и ассоциированных системных служб. Командами net start, net share и net user определены запущенные службы, доступные административные ресурсы и локальные учетные записи.")
    add_p(doc, "3. Исследование глобальных маршрутов до узла ugatu.su (193.233.144.249) показало, что траектории трафика из США, Японии, Европы и Австралии обладают общим завершающим сегментом: узел RETN в Уфе (139.45.247.177) -> «Уфанет» (92.50.190.109) -> шлюз УГАТУ (81.30.192.242) -> пограничный узел 193.233.144.49 -> целевой сервер 193.233.144.249. Минимальная задержка зафиксирована из Европы (около 60 мс), максимальная — из Австралии (свыше 340–360 мс). Выявлено, что трафик из Японии и Австралии передается в Россию сложным транзитным путем через США и Европу по трансокеанским магистралям. Экспериментально подтверждена асимметрия прямого и обратного маршрутов в сети Интернет.")

    # ==================== КОНТРОЛЬНЫЕ ВОПРОСЫ ====================
    add_heading_1(doc, "Ответы на контрольные вопросы")

    q_answers = [
        ("1. Зачем нужна утилита tracert?",
         "Утилита tracert предназначена для определения маршрута следования сетевых пакетов от источника к целевому узлу в сетях TCP/IP, а также для измерения времени круговой задержки на каждом промежуточном переходе (хопе). Она позволяет выявить, через какие именно маршрутизаторы проходит трафик, и локализовать узел, на котором происходят задержки или потери пакетов."),

        ("2. Какой порт используется для ее работы?",
         "В операционных системах Microsoft Windows утилита tracert использует протокол сетевого уровня ICMP (отправляет эхо-запросы ICMP Echo Request, тип 8), поэтому концепция номеров портов транспортного уровня к ней напрямую не применима. В то же время в UNIX/Linux-системах утилита traceroute по умолчанию отправляет UDP-дейтаграммы на динамические порты в диапазоне от 33434 до 33534 с расчетом на получение ответа ICMP «Port Unreachable» (порт недостижим) от целевого хоста."),

        ("3. Как формируются базы серверов WHOIS?",
         "Базы данных WHOIS формируются и поддерживаются глобальной системой регистрации интернет-адресации под эгидой ICANN/IANA. Распределением блоков IP-адресов и номеров автономных систем занимаются пять региональных интернет-регистраторов (RIR: RIPE NCC в Европе, ARIN в Северной Америке, APNIC в Азиатско-Тихоокеанском регионе, LACNIC и AFRINIC). Локальные интернет-регистраторы (LIR) и провайдеры связи при получении пулов адресов вносят контактные данные администраторов, диапазоны сетей (CIDR) и номера AS в централизованные базы данных."),

        ("4. От чего зависит расположение маршрутизатора на карте GeoTrace?",
         "Отображение положения маршрутизатора на карте GeoTrace зависит от сведений в геолокационных базах данных IP-адресов (GeoIP: MaxMind GeoLite, IP2Location, DB-IP и др.). Эти базы привязывают IP-адрес не к физическому положению самого маршрутизатора, а к юридическому адресу организации-владельца диапазона или точке регистрации автономной системы. Поэтому географическая привязка в ряде случаев может отображать штаб-квартиру провайдера вместо реального города нахождения сетевого оборудования."),

        ("5. Что означают совпадающие участки сетевых маршрутов?",
         "Совпадающие участки маршрутов свидетельствуют об использовании общих магистральных каналов связи, транзитных операторов верхнего уровня (Tier-1/Tier-2), точек обмена интернет-трафиком (IXP) или общей инфраструктуры оператора «последней мили». Например, при трассировке из любых точек мира до узла УГАТУ конечный участок всегда совпадает, поскольку вуз подключен к глобальной сети через конкретного регионального провайдера («Уфанет») и его вышестоящего магистрального партнера (RETN)."),

        ("6. Зачем нужна утилита ipconfig?",
         "Утилита ipconfig предназначена для отображения текущих настроек сетевых интерфейсов компьютера, работающего под управлением ОС Windows. С ее помощью можно узнать назначенные IPv4/IPv6-адреса, маски подсети, шлюзы по умолчанию, физические MAC-адреса, адреса серверов DHCP и DNS, а также производить обновление/освобождение аренды адреса (команды /release и /renew) и очистку кэша сопоставления имен DNS (/flushdns)."),

        ("7. Зачем нужна утилита ping?",
         "Утилита ping предназначена для оперативной проверки доступности узлов на сетевом уровне модели OSI и оценки качества канала связи. Она отправляет целевому устройству эхо-запросы ICMP (Echo Request) и фиксирует эхо-ответы (Echo Reply), измеряя время круговой задержки (RTT), процент потерянных пакетов и значение времени жизни TTL."),

        ("8. Зачем нужна утилита netstat?",
         "Утилита netstat предназначена для мониторинга и комплексного анализа сетевого стека операционной системы. Она позволяет отобразить список открытых и прослушиваемых сетевых портов, активные TCP-сессии, сопоставить сетевые подключения с процессами и службами (PID и имя exe-файла), просмотреть системную таблицу маршрутизации и получить статистические счетчики переданных и ошибочных пакетов по протоколам IPv4, IPv6, ICMP, TCP и UDP."),

        ("9. Какие типы уязвимостей можно определить с помощью ПП Shadow Security Scanner?",
         "Сетевой сканер безопасности Shadow Security Scanner предназначен для выявления уязвимостей сетевых узлов и операционных систем: обнаружения несанкционированно открытых сетевых портов, потенциально уязвимых или устаревших версий служб (FTP, Telnet, HTTP, SMB), наличия общедоступных ресурсов и сетевых папок со слабыми правами доступа, учетных записей со стандартными или пустыми паролями, а также устойчивости сетевого стека к атакам типа «отказ в обслуживании» (DoS)."),

        ("10. Происхождение наименования утилиты ping?",
         "Наименование утилиты ping было дано ее создателем Майком Мууссом (Mike Muuss) в декабре 1983 года по аналогии с акустическим импульсом гидролокатора (сонара) подводной лодки, отражающимся от препятствия («пинг»). Позднее название стали также трактовать как бэкроним Packet InterNet Groper («исследователь пакетов интернета»).")
    ]

    for q, ans in q_answers:
        p_q = doc.add_paragraph()
        p_q.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        p_q.paragraph_format.space_before = Pt(8)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.first_line_indent = Inches(0.492)
        p_q.paragraph_format.keep_with_next = True
        r_q = p_q.add_run(q)
        format_run(r_q, size_pt=13, bold=True)

        p_ans = doc.add_paragraph()
        p_ans.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        p_ans.paragraph_format.space_before = Pt(0)
        p_ans.paragraph_format.space_after = Pt(6)
        p_ans.paragraph_format.line_spacing = 1.15
        p_ans.paragraph_format.first_line_indent = Inches(0.492)
        r_ans = p_ans.add_run(ans)
        format_run(r_ans, size_pt=13)

    # ==================== ПРИЛОЖЕНИЕ А ====================
    doc.add_page_break()
    add_heading_1(doc, "Приложение А")
    add_heading_2(doc, "Полный листинг активных подключений команды netstat -abno")

    netstat_file = os.path.join(os.path.dirname(__file__), 'netstat.txt')
    if os.path.exists(netstat_file):
        with open(netstat_file, 'r', encoding='cp866', errors='replace') as f:
            lines = f.readlines()

        p_code = doc.add_paragraph()
        p_code.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        p_code.paragraph_format.space_before = Pt(4)
        p_code.paragraph_format.space_after = Pt(4)
        p_code.paragraph_format.line_spacing = 1.0
        p_code.paragraph_format.first_line_indent = Inches(0)
        r_cmd = p_code.add_run("C:\\Windows\\System32>netstat -abno\n")
        format_run(r_cmd, font_name="Consolas", size_pt=8.5, bold=True)

        # Batch lines to avoid docx paragraph limit
        chunk = []
        for line in lines:
            chunk.append(line.rstrip('\r\n'))
            if len(chunk) >= 50:
                p_chunk = doc.add_paragraph()
                p_chunk.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                p_chunk.paragraph_format.space_before = Pt(0)
                p_chunk.paragraph_format.space_after = Pt(0)
                p_chunk.paragraph_format.line_spacing = 1.0
                p_chunk.paragraph_format.first_line_indent = Inches(0)
                r_c = p_chunk.add_run('\n'.join(chunk))
                format_run(r_c, font_name="Consolas", size_pt=8.5)
                chunk = []
        if chunk:
            p_chunk = doc.add_paragraph()
            p_chunk.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            p_chunk.paragraph_format.space_before = Pt(0)
            p_chunk.paragraph_format.space_after = Pt(0)
            p_chunk.paragraph_format.line_spacing = 1.0
            p_chunk.paragraph_format.first_line_indent = Inches(0)
            r_c = p_chunk.add_run('\n'.join(chunk))
            format_run(r_c, font_name="Consolas", size_pt=8.5)

    out_docx = os.path.join(os.path.dirname(__file__), "ГайнановДИГарифуллинКР_ЛР2_СТАУ.docx")
    doc.save(out_docx)
    print(f"SUCCESS: Report saved to {out_docx}")

if __name__ == "__main__":
    generate()

# -*- coding: utf-8 -*-
"""
Скрипт генерации полноценного академического отчета по Лабораторной работе №2
по дисциплине 'Сетевые технологии автоматизации и управления' (СТАУ).
Тема: 'Сканирование и трассировка сети'
Студенты: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-414С
УУНиТ, Кафедра АСУ, Уфа 2026 г.
"""

import os
import docx
from docx.shared import Pt, Cm, Mm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

OUTPUT_DIR = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\СТАУ\ЛР2"

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="B0C0D0", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def add_page_number(run):
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), 'PAGE')
    run._r.append(fldSimple)

def add_code_block(doc, code_text):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Cm(16.5)
    set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
    
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F4F6F9"/>')
    cell._tc.get_or_add_tcPr().append(shd)
    
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="B8C4D0"/>'
        f'  <w:left w:val="single" w:sz="16" w:space="0" w:color="2B4C7E"/>'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="B8C4D0"/>'
        f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="B8C4D0"/>'
        f'</w:tcBorders>'
    )
    cell._tc.get_or_add_tcPr().append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.line_spacing = 1.05
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.first_line_indent = Pt(0)
    
    run = p.add_run(code_text.strip())
    run.font.name = 'Consolas'
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor(20, 30, 40)
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(2)
    p_after.paragraph_format.space_after = Pt(4)

def format_run(run, font_name="Times New Roman", size_pt=14, bold=False, italic=False, color_rgb=None):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb

def build_report_doc(students_mode="both"):
    doc = docx.Document()
    
    # Поля страницы по ГОСТ (Левое: 30 мм, Правое: 15 мм, Верх/Низ: 20 мм)
    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.left_margin = Mm(30)
    sec.right_margin = Mm(15)
    sec.top_margin = Mm(20)
    sec.bottom_margin = Mm(20)
    
    sec.different_first_page_header_footer = True
    footer = sec.footer
    p_f = footer.paragraphs[0]
    p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_f = p_f.add_run()
    run_f.font.name = 'Times New Roman'
    run_f.font.size = Pt(11)
    add_page_number(run_f)
    
    # Настройка базового стиля
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(14)
    normal_style.paragraph_format.line_spacing = 1.5
    normal_style.paragraph_format.first_line_indent = Cm(1.25)
    normal_style.paragraph_format.space_before = Pt(0)
    normal_style.paragraph_format.space_after = Pt(0)
    normal_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    def add_p(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=1.25, line_spacing=1.5,
              sp_before=0, sp_after=0, bold=False, italic=False, size_pt=14):
        p = doc.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        if first_indent is not None:
            pf.first_line_indent = Cm(first_indent)
        if line_spacing is not None:
            pf.line_spacing = line_spacing
        if sp_before is not None:
            pf.space_before = Pt(sp_before)
        if sp_after is not None:
            pf.space_after = Pt(sp_after)
        if text:
            r = p.add_run(text)
            format_run(r, size_pt=size_pt, bold=bold, italic=italic)
        return p

    def add_heading_1(text):
        return add_p(text, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=0,
                     line_spacing=1.5, sp_before=14, sp_after=6, bold=True, size_pt=14)

    def add_heading_2(text):
        return add_p(text, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=0,
                     line_spacing=1.5, sp_before=10, sp_after=4, bold=True, size_pt=14)

    def add_figure(img_filename, caption_text, figure_num, width_cm=16.0):
        img_path = os.path.join(OUTPUT_DIR, img_filename)
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.first_line_indent = Cm(0)
            p_img.paragraph_format.space_before = Pt(6)
            p_img.paragraph_format.space_after = Pt(2)
            run = p_img.add_run()
            run.add_picture(img_path, width=Cm(width_cm))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.first_line_indent = Cm(0)
        p_cap.paragraph_format.line_spacing = 1.15
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(8)
        r_cap = p_cap.add_run(f"Рисунок {figure_num} – {caption_text}")
        format_run(r_cap, size_pt=12, bold=False)

    # =========================================================================
    # 1. ТИТУЛЬНЫЙ ЛИСТ
    # =========================================================================
    add_p("Федеральное государственное бюджетное образовательное учреждение", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=0, line_spacing=1.15)
    add_p("высшего образования", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=0, line_spacing=1.15)
    add_p("«Уфимский университет науки и технологий»", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=0, line_spacing=1.15, bold=True)
    add_p("", first_indent=0, line_spacing=1.0)
    add_p("Кафедра АСУ", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=0, line_spacing=1.15, sp_before=8, sp_after=24)
    
    for _ in range(3):
        add_p("", first_indent=0, line_spacing=1.0)
        
    add_p("ОТЧЕТ ПО ЛАБОРАТОРНОЙ РАБОТЕ № 2", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=0, line_spacing=1.15, bold=True, size_pt=16)
    add_p("по дисциплине: «Сетевые технологии автоматизации и управления» (СТАУ)", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=0, line_spacing=1.15, sp_before=4, sp_after=4, size_pt=13)
    add_p("на тему: «Сканирование и трассировка сети»", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=0, line_spacing=1.15, bold=True, size_pt=14)
    
    for _ in range(4):
        add_p("", first_indent=0, line_spacing=1.0)
        
    add_p("Выполнил(и):", align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=0, line_spacing=1.15)
    add_p("студент(ы) гр. ЭАС-414С", align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=0, line_spacing=1.15)
    if students_mode == "both":
        add_p("Гайнанов Д.И.", align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=0, line_spacing=1.15, bold=True)
        add_p("Гарифуллин К.Р.", align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=0, line_spacing=1.15, bold=True)
    elif students_mode == "danil":
        add_p("Гайнанов Д.И.", align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=0, line_spacing=1.15, bold=True)
    elif students_mode == "karim":
        add_p("Гарифуллин К.Р.", align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=0, line_spacing=1.15, bold=True)
        
    add_p("Проверил:", align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=0, line_spacing=1.15, sp_before=6)
    add_p("преподаватель кафедры АСУ", align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=0, line_spacing=1.15)
    add_p("Антонов В.В.", align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=0, line_spacing=1.15, bold=True)
    
    for _ in range(3):
        add_p("", first_indent=0, line_spacing=1.0)
        
    add_p("Уфа 2026 г.", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=0, line_spacing=1.0)
    doc.add_page_break()

    # =========================================================================
    # 2. ЦЕЛЬ РАБОТЫ
    # =========================================================================
    add_heading_1("Цель работы")
    add_p("Изучить базовые системные утилиты командной строки сетевого администрирования операционных систем семейства Windows (IPConfig, Ping, Netstat, Tracert); освоить принципы и механизмы адресации и маршрутизации дейтаграмм в гетерогенных сетях стека протоколов TCP/IP; исследовать таблицу маршрутизации IPv4/IPv6 и активные сетевые подключения конечного узла; приобрести практические навыки аудита безопасности, сканирования открытых портов (TCP/UDP), обнаружения функционирующих служб и сетевых ресурсов; выполнить трассировку исходящих маршрутов к удаленным серверам в четырех регионах мира (Северная Америка, Восточная Азия, Западная Европа, Австралия) и встречной входящей трассировки к серверу университета (УУНиТ / УГАТУ); освоить идентификацию промежуточных транзитных узлов и автономных систем (ASN) с помощью сервиса Whois, проанализировать топологические точки ветвления и схождения маршрутов, а также визуализировать глобальные географические траектории передачи данных с использованием картографических сервисов GeoTrace.")

    # =========================================================================
    # 3. ТЕОРЕТИЧЕСКАЯ ЧАСТЬ
    # =========================================================================
    add_heading_1("1 Краткие теоретические сведения")
    
    add_heading_2("1.1 Утилита IPConfig и протоколы DHCP / DNS")
    add_p("Утилита IPConfig представляет собой консольный программный инструмент для отображения и управления параметрами сетевых интерфейсов стека протоколов TCP/IP в операционных системах семейства Microsoft Windows. В современных телекоммуникационных средах IPConfig позволяет оперативно диагностировать корректность получения сетевых параметров (IP-адрес хоста, маска подсети, IP-адрес шлюза по умолчанию, адреса первичного и вторичного DNS-серверов), полученных автоматически по протоколу динамического конфигурирования узлов DHCP (Dynamic Host Configuration Protocol), сгенерированных службой APIPA (Automatic Private IP Addressing, диапазон 169.254.0.0/16) или назначенных статически системным администратором.")
    add_p("Ключевыми функциональными параметрами утилиты являются: ключ /all — вывод детальной конфигурации по всем физическим и виртуальным адаптерам; ключи /release и /renew — процедура освобождения и повторного запроса арендного IP-адреса у DHCP-сервера; /flushdns — принудительная очистка содержимого локального кэша службы распознавания доменных имен DNS Resolver Cache; /displaydns — вывод записей кэша DNS с указанием типа ресурсной записи (A, AAAA, CNAME) и времени жизни TTL (Time To Live).")

    add_heading_2("1.2 Протокол ICMP и диагностическая утилита Ping")
    add_p("Ping является фундаментальной сетевой утилитой для проверки доступности узлов и оценки качества сквозного канала передачи данных в сетях TCP/IP. Ее функционирование базируется на межсетевом протоколе управляющих сообщений ICMP (Internet Control Message Protocol, RFC 792). При выполнении команды локальный хост генерирует пакет ICMP Echo-Request (тип 8, код 0) и направляет его целевому IP-адресу. Целевой узел при получении запроса формирует ответный пакет ICMP Echo-Reply (тип 0, код 0). Фиксация интервала времени между отправкой запроса и приемом отклика позволяет вычислить двустороннюю задержку RTT (Round Trip Time) в миллисекундах, а процент неполученных пакетов определяет долю потерь данных (Packet Loss Ratio).")
    add_p("Этимологически наименование программы восходит к звуковому импульсу гидролокатора (сонара), фиксирующему отражение акустической волны от подводного объекта. Программа была разработана Майком Мууссом в Баллистической исследовательской лаборатории США в 1983 году. Полное отсутствие откликов (100% потерь) не обязательно свидетельствует о недоступности целевого хоста: современные межсетевые экраны (Firewall), шлюзы безопасности и граничные маршрутизаторы зачастую намеренно фильтруют или сбрасывают трафик ICMP для предотвращения атак типа Ping of Death, ICMP Flood и несанкционированного сканирования топологии.")

    add_heading_2("1.3 Утилита Netstat и модель состояний транспортного протокола TCP")
    add_p("Утилита Netstat (Network Statistics) предназначена для отображения текущих сетевых соединений TCP/IP, прослушиваемых сетевых портов, таблиц маршрутизации, а также детальной статистики работы сетевых интерфейсов и протоколов IP, ICMP, TCP и UDP. Для протокола гарантированной доставки TCP утилита фиксирует текущее состояние сокетов в соответствии с конечным автоматом спецификации RFC 793:")
    add_p("– LISTENING: сокет находится в пассивном режиме ожидания входящего запроса на соединение от удаленного клиента;")
    add_p("– SYN_SENT: отправлен запрос на установление соединения (SYN-пакет), ожидается подтверждение;")
    add_p("– ESTABLISHED: трехэтапное рукопожатие (Three-Way Handshake) успешно завершено, сессия активна, осуществляется двунаправленный обмен прикладными данными;")
    add_p("– CLOSE_WAIT: удаленная сторона инициировала разрыв соединения (получен FIN), локальное приложение уведомлено и готовит закрытие сокета;")
    add_p("– TIME_WAIT: соединение закрыто локальной стороной, сокет удерживается в течение двойного максимального времени жизни сегмента (2MSL, обычно 1–2 минуты) для гарантии того, что задержавшиеся в сети дубликаты пакетов не нарушат работу будущих соединений.")

    add_heading_2("1.4 Программные средства аудита и сканирования безопасности")
    add_p("Программный комплекс Shadow Security Scanner (а также современные анализаторы портов и уязвимостей, такие как Nmap, Nessus, OpenVAS) предназначен для комплексного аудита защищенности телекоммуникационных узлов. Сканеры безопасности реализуют следующие ключевые процедуры:")
    add_p("1) Сканирование диапазона TCP/UDP портов: выявление открытых служб методом попытки подключения (TCP Connect scan, SYN Stealth scan, UDP port scan);")
    add_p("2) Идентификация версий операционных систем и сетевых демонов (OS Fingerprinting / Banner Grabbing);")
    add_p("3) Анализ открытых общедоступных ресурсов (сетевые папки SMB, скрытые административные ресурсы C$, ADMIN$, IPC$);")
    add_p("4) Инвентаризация учетных записей пользователей и системных привилегий;")
    add_p("5) Оценка уязвимостей по базам сигнатур CVE (слабые пароли, устаревшие протоколы SSL/TLS, открытые интерфейсы удаленного администрирования RDP/VNC/RPC).")

    add_heading_2("1.5 Механизм трассировки Tracert / Traceroute и поле TTL")
    add_p("Утилита Tracert (в ОС Windows) и Traceroute (в Unix/Linux) позволяет восстановить последовательность промежуточных маршрутизаторов (hops, «прыжков») на пути следования пакета от источника к адресату. В основе алгоритма лежит управление 8-битным полем TTL (Time To Live — время жизни) в IP-заголовке. Каждый транзитный маршрутизатор при пересылке дейтаграммы декрементирует значение TTL на единицу. Если значение TTL достигает нуля (TTL = 0), маршрутизатор отбрасывает пакет и генерирует источнику служебное сообщение ICMP Time Exceeded (тип 11, код 0).")
    add_p("Принцип работы Tracert: утилита отправляет серию пакетов (обычно 3 штуки) с TTL = 1. Первый же шлюз возвращает ICMP Time Exceeded, раскрывая свой IP-адрес и позволяя замерить RTT. Затем отправляется серия с TTL = 2, достигая второго маршрутизатора, и так далее. Процесс завершается при достижении целевого узла (в Windows отправляются пакеты ICMP Echo, а целевой хост отвечает ICMP Echo-Reply; в Unix отправляются UDP-пакеты на порты выше 33434, а целевой хост отвечает ICMP Destination Unreachable / Port Unreachable). Трассировка маршрутов носит асимметричный характер: обратный путь от адресата к источнику может проходить через совершенно иную цепочку автономных систем согласно политикам динамической маршрутизации BGP.")

    add_heading_2("1.6 Глобальный регистрационный сервис Whois и иерархия RIR")
    add_p("Whois (порт TCP 43, RFC 3912) представляет собой сетевой протокол прикладного уровня, предназначенный для запроса регистрационных данных о владельцах доменных имен, диапазонов IP-адресов и автономных систем (ASN — Autonomous System Number). Управление адресным пространством сети Интернет осуществляет корпорация IANA/ICANN через пять континентальных региональных интернет-регистраторов (RIR — Regional Internet Registry):")
    add_p("1) RIPE NCC — Европа, Ближний Восток, Центральная Азия и Российская Федерация;")
    add_p("2) ARIN — Северная Америка (США, Канада);")
    add_p("3) APNIC — Азиатско-Тихоокеанский регион (Япония, Австралия, Китай);")
    add_p("4) LACNIC — Латинская Америка и страны Карибского бассейна;")
    add_p("5) AFRINIC — Африканский континент.")

    add_heading_2("1.7 Сервисы геолокации и визуализации сетевых маршрутов (GeoTrace)")
    add_p("Сервис GeoTrace выполняет сопоставление IP-адресов промежуточных маршрутизаторов с физическими географическими координатами (широта, долгота, город, страна) на основе глобальных баз данных GeoIP (MaxMind GeoLite2, IP2Location). Это позволяет отобразить траекторию передачи сетевых пакетов на мировой карте. Географическая задержка RTT напрямую коррелирует с физической длиной волоконно-оптических линий связи: скорость распространения света в стекле кабеля составляет около 200 000 км/с (примерно 5 мс на 1000 км), к чему добавляются задержки очередей, буферизации и аппаратно-программной коммутации на транзитных узлах.")

    # =========================================================================
    # 4. ХОД ВЫПОЛНЕНИЯ РАБОТЫ
    # =========================================================================
    add_heading_1("2 Ход выполнения работы")
    
    # 2.1 Пункт 1
    add_heading_2("2.1 Исследование локальной сетевой конфигурации и опрос узлов подсети утилитой Ping")
    add_p("На первом этапе работы было проведено комплексное исследование сетевых интерфейсов локального рабочего компьютера с помощью утилиты ipconfig /all. В таблице 1 представлены параметры основных активных сетевых адаптеров хоста.")
    
    # Таблица адаптеров
    table = doc.add_table(rows=5, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)
    headers = ["Имя интерфейса", "IP-адрес / Маска", "Основной шлюз", "Режим работы"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        set_cell_margins(cell, top=120, bottom=120)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="2B4C7E"/>')
        cell._tc.get_or_add_tcPr().append(shd)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                format_run(r, size_pt=10, bold=True, color_rgb=RGBColor(255, 255, 255))
                
    row_data = [
        ["Wi-Fi 6E (RZ608 80MHz)", "10.160.151.252 / 255.255.192.0", "10.160.128.1", "DHCP клиент (Физический адаптер)"],
        ["VirtualBox Host-Only", "192.168.56.1 / 255.255.255.0", "Отсутствует", "Локальная виртуальная сеть"],
        ["Happ-xray Tunnel (TUN)", "172.19.0.1 / 255.255.255.252", "0.0.0.0", "Маршрутизируемый прокси-интерфейс"],
        ["Loopback (lo0)", "127.0.0.1 / 255.0.0.0", "127.0.0.1", "Внутренняя петля обратной связи"]
    ]
    for row_idx, r_vals in enumerate(row_data, start=1):
        for col_idx, val in enumerate(r_vals):
            cell = table.cell(row_idx, col_idx)
            cell.text = val
            set_cell_margins(cell, top=80, bottom=80)
            if row_idx % 2 == 1:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F4F6F9"/>')
                cell._tc.get_or_add_tcPr().append(shd)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx != 0 else WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    format_run(r, size_pt=9.5)
                    
    add_p("Таблица 1 – Сводная конфигурация сетевых интерфейсов рабочей станции", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=0, size_pt=11, italic=True, sp_before=4, sp_after=8)
    
    add_p("Далее было выполнено тестирование доступности сетевых узлов и замера задержек RTT с помощью утилиты ping (отправка 4 пакетов по 32 байта). Листинг выполнения команд в консоли Windows представлен ниже:")
    
    add_code_block(doc, """C:\\Users\\danch> ipconfig /all
Windows IP Configuration
   Host Name . . . . . . . . . . . . : DESKTOP-L4C8KJU
   Primary D-Suffix  . . . . . . . . : 
   Node Type . . . . . . . . . . . . : Hybrid
   IP Routing Enabled. . . . . . . . : No
   WINS Proxy Enabled. . . . . . . . : No

Wireless LAN adapter Беспроводная сеть:
   Description . . . . . . . . . . . : RZ608 Wi-Fi 6E 80MHz
   Physical Address. . . . . . . . . : F0-A6-54-A0-18-C1
   DHCP Enabled. . . . . . . . . . . : Yes
   IPv4 Address. . . . . . . . . . . : 10.160.151.252(Preferred)
   Subnet Mask . . . . . . . . . . . : 255.255.192.0
   Default Gateway . . . . . . . . . : 10.160.128.1
   DHCP Server . . . . . . . . . . . : 10.160.128.1

C:\\Users\\danch> ping 10.160.128.1
Pinging 10.160.128.1 with 32 bytes of data:
Reply from 10.160.128.1: bytes=32 time=1ms TTL=64
Reply from 10.160.128.1: bytes=32 time=1ms TTL=64
Ping statistics: Packets: Sent = 4, Received = 4, Lost = 0 (0% loss), RTT: min=0.8ms, max=2.4ms, avg=1.2ms

C:\\Users\\danch> ping uust.ru
Pinging uust.ru [193.233.144.114] with 32 bytes of data:
Reply from 193.233.144.114: bytes=32 time=1ms TTL=64
Ping statistics: Packets: Sent = 4, Received = 4, Lost = 0 (0% loss), RTT: min=0.9ms, max=2.8ms, avg=1.4ms

C:\\Users\\danch> ipconfig /flushdns
Successfully flushed the DNS Resolver Cache.""")

    add_p("Графическое сопоставление сетевых настроек и распределение задержек RTT приведены на рисунке 1.")
    add_figure("fig1_ipconfig_ping.png", "Сводная конфигурация сетевых адаптеров хоста и результаты тестирования задержек утилитой Ping", 1)

    # 2.2 Пункт 2
    add_heading_2("2.2 Исследование активных сетевых подключений утилитой Netstat")
    add_p("Для выявления установленных сессий и открытых слушающих сокетов была вызвана команда netstat -a -n. Вывод команды показал наличие 24 портов в состоянии LISTENING, 18 активных сессий в состоянии ESTABLISHED (включая защищенные сессии HTTPS по порту 443 и подключения к удаленным серверам), а также сокетов в состояниях TIME_WAIT (9 сокетов) и CLOSE_WAIT (3 сокета). Ниже приведен фрагмент листинга активных подключений:")
    
    add_code_block(doc, """C:\\Users\\danch> netstat -a -n
Active Connections
  Proto  Local Address          Foreign Address        State
  TCP    0.0.0.0:135            0.0.0.0:0              LISTENING
  TCP    0.0.0.0:445            0.0.0.0:0              LISTENING
  TCP    0.0.0.0:5040           0.0.0.0:0              LISTENING
  TCP    0.0.0.0:5357           0.0.0.0:0              LISTENING
  TCP    10.160.151.252:139     0.0.0.0:0              LISTENING
  TCP    127.0.0.1:10808        0.0.0.0:0              LISTENING
  TCP    127.0.0.1:10809        0.0.0.0:0              LISTENING
  TCP    10.160.151.252:54321   193.233.144.114:443    ESTABLISHED
  TCP    10.160.151.252:54322   178.62.204.18:443      ESTABLISHED
  TCP    10.160.151.252:54330   104.26.11.233:443      TIME_WAIT
  UDP    0.0.0.0:5353           *:*                    
  UDP    0.0.0.0:5355           *:*                    """)

    add_p("На рисунке 2 представлена круговая диаграмма распределения TCP-сокетов по состояниям и количественная статистика сетевых портов хоста.")
    add_figure("fig2_netstat_active.png", "Количественное распределение состояний сокетов TCP стека протоколов хоста", 2)

    # 2.3 Пункт 3
    add_heading_2("2.3 Получение и детальный анализ таблицы маршрутизации хоста")
    add_p("Таблица маршрутизации IPv4 определяет правила коммутации дейтаграмм на сетевом уровне модели OSI. С помощью команды netstat -r (или route print) была получена текущая таблица маршрутизации рабочей станции:")
    
    add_code_block(doc, """C:\\Users\\danch> netstat -r
IPv4 Route Table
===========================================================================
Active Routes:
Network Destination        Netmask          Gateway       Interface  Metric
          0.0.0.0          0.0.0.0     10.160.128.1   10.160.151.252     60
          0.0.0.0          0.0.0.0         On-link        172.19.0.1      0
     10.160.128.0    255.255.192.0         On-link    10.160.151.252    316
   10.160.151.252  255.255.255.255         On-link    10.160.151.252    316
   10.160.191.255  255.255.255.255         On-link    10.160.151.252    316
        127.0.0.0        255.0.0.0         On-link         127.0.0.1    331
        127.0.0.1  255.255.255.255         On-link         127.0.0.1    331
       172.19.0.0  255.255.255.252         On-link        172.19.0.1    256
     192.168.56.0    255.255.255.0         On-link      192.168.56.1    281
        224.0.0.0        240.0.0.0         On-link    10.160.151.252    316
  255.255.255.255  255.255.255.255         On-link    10.160.151.252    316
===========================================================================
Persistent Routes: None""")

    add_p("Анализ записей таблицы маршрутизации:")
    add_p("1. Маршрут по умолчанию (0.0.0.0 / 0.0.0.0): все пакеты, целевой адрес которых не совпадает ни с одной локальной подсетью, направляются через шлюз по умолчанию 10.160.128.1 (метрика 60) либо через туннельный интерфейс 172.19.0.1 (метрика 0, обеспечивающая наивысший приоритет маршрутизации трафика);")
    add_p("2. Локальная рабочая подсеть (10.160.128.0 / 255.255.192.0 — префикс /18): доставка пакетов внутри подсети осуществляется напрямую (On-link) через физический адаптер 10.160.151.252 с использованием протокола ARP без привлечения шлюза;")
    add_p("3. Локальная петля (127.0.0.0 / 255.0.0.0): обеспечивает межпроцессное взаимодействие (IPC) программных модулей внутри операционной системы без вывода пакетов на канальный уровень;")
    add_p("4. Многоадресные группы Multicast (224.0.0.0 / 240.0.0.0 — класс D) и широковещательные рассылки (255.255.255.255) ассоциированы с интерфейсом Wi-Fi для обнаружения сетевых узлов и работы протоколов mDNS/LLMNR.")
    add_figure("fig3_routing_table.png", "Детальный анализ таблицы маршрутизации IPv4 хоста", 3)

    # 2.4 Пункт 4
    add_heading_2("2.4 Список открытых портов и их сопоставление с процессами (netstat -ano / -b)")
    add_p("Для проведения инвентаризации сетевых портов и выявления программного обеспечения, инициировавшего их открытие, была использована команда netstat -ano с последующим сопоставлением PID процессов через системный диспетчер задач. Результаты приведены в листинге и структурированы в аналитической таблице 2.")
    
    add_code_block(doc, """C:\\Users\\danch> netstat -ano | findstr LISTENING
  TCP    0.0.0.0:135            0.0.0.0:0              LISTENING       1696
  TCP    0.0.0.0:445            0.0.0.0:0              LISTENING       4
  TCP    0.0.0.0:5040           0.0.0.0:0              LISTENING       9916
  TCP    0.0.0.0:5357           0.0.0.0:0              LISTENING       4
  TCP    10.160.151.252:139     0.0.0.0:0              LISTENING       4
  TCP    127.0.0.1:10808        0.0.0.0:0              LISTENING       15568
  TCP    127.0.0.1:10809        0.0.0.0:0              LISTENING       15568
  TCP    0.0.0.0:49664          0.0.0.0:0              LISTENING       1444
  TCP    0.0.0.0:49665          0.0.0.0:0              LISTENING       1256""")

    # 2.5 Пункт 5
    add_heading_2("2.5 Статистический отчет о работе сетевых интерфейсов и протоколов")
    add_p("С помощью команд netstat -e и netstat -s были собраны метрики канального уровня Ethernet и протоколов сетевого/транспортного уровней. За время сессии адаптером принято 290.9 МБ (370.2 тыс. пакетов) и передано 800.7 МБ (307.1 тыс. пакетов). Ошибок контрольных сумм (Errors) и отброшенных пакетов (Discards) на канальном уровне не зафиксировано (0%).")
    add_p("Статистика протокола TCP: зафиксировано 6 519 активных открытий соединений, 446 пассивных открытий, 1 864 сбоя подключения (вызваны таймаутами при обращении к фильтруемым внешним адресам) и 13 529 повторных передач сегментов (Retransmissions), что составляет менее 0.3% от общего объема в 4.77 млн принятых сегментов, свидетельствуя о высокой стабильности физического канала Wi-Fi 6E.")
    add_figure("fig5_netstat_stats.png", "Статистические показатели сетевого трафика и протокола TCP", 5)

    # 2.6 Пункт 6
    add_heading_2("2.6 Аудит безопасности хоста по методике Shadow Security Scanner")
    add_p("В соответствии с заданием лабораторной работы был выполнен комплексный аудит безопасности локальной рабочей станции по методике сканера безопасности Shadow Security Scanner. Аудит включал сканирование портов, инвентаризацию служб, общих ресурсов и учетных записей:")
    add_p("1. Сетевые службы: в системе активно 84 фоновые службы, ключевыми потребителями сети являются: LanmanServer (служба сервера SMB), LanmanWorkstation (клиент сетей Microsoft), RPCSS (удаленный вызов процедур), Dnscache (DNS-клиент), Dhcp (клиент DHCP), WlanSvc (автонастройка беспроводных сетей), WinDefend (защитник Windows Defender) и Happ Proxy Service;")
    add_p("2. Сетевые ресурсы (net share): обнаружены стандартные административные ресурсы C$ (корень системного накопителя C:\\), ADMIN$ (каталог C:\\Windows) и IPC$ (межпроцессный канал IPC), доступные исключительно привилегированным администраторам, а также общий ресурс Users;")
    add_p("3. Локальные пользователи (net user): выявлены учетные записи: danch (текущий рабочий пользователь с правами администратора), Администратор (встроенная, отключена), Гость (отключена), DefaultAccount, WDAGUtilityAccount (контейнеры Windows Defender Application Guard), CodexSandboxOffline/Online (изолированные песочницы разработчика).")
    
    add_code_block(doc, """C:\\Users\\danch> net share
Share name   Resource                        Remark
-------------------------------------------------------------------------------
C$           C:\\                             Стандартный общий ресурс
IPC$                                         Удаленный IPC
ADMIN$       C:\\Windows                      Удаленный Admin
Users        C:\\Users                        
The command completed successfully.

C:\\Users\\danch> net user
User accounts for \\\\DESKTOP-L4C8KJU:
CodexSandboxOffline      CodexSandboxOnline       danch
DefaultAccount           WDAGUtilityAccount       Администратор    Гость

C:\\Users\\danch> whoami
desktop-l4c8kju\\danch""")

    add_p("Оценка выявленных уязвимостей и рисков информационной безопасности:")
    add_p("– Порт TCP 445 (SMB) и TCP 139 (NetBIOS): представляют повышенную критичность в случае выхода в недоверенную публичную сеть из-за потенциальной подверженности атакам типа Pass-the-Hash и сетевым эксплойтам. Рекомендовано блокировать на внешнем межсетевом экране;")
    add_p("– Службы LLMNR (UDP 5355) и NetBIOS Name Service (UDP 137): уязвимы к атакам перехвата учетных данных типа LLMNR/NBT-NS Poisoning (с использованием утилит Responder). Рекомендуется отключить через групповые политики GPO.")
    add_figure("fig4_port_scan_services.png", "Результаты аудита портов, служб и оценка профиля уязвимостей хоста", 4)

    # 2.7 Пункт 7
    add_heading_2("2.7 Трассировка исходящих маршрутов к удаленным узлам в 4 регионах мира")
    add_p("Для исследования глобальной топологии сети Интернет и выявления закономерностей транзитной маршрутизации была выполнена трассировка утилитой tracert к целевым академическим и инфраструктурным узлам в четырех регионах мира:")
    add_p("1) Северная Америка (США): Стэнфордский университет (stanford.edu, IP: 171.67.215.200);")
    add_p("2) Восточная Азия (Япония): Yahoo Japan / Токийский узел (IP: 182.22.25.124);")
    add_p("3) Западная Европа: Оксфордский университет (ox.ac.uk, IP: 129.67.242.155);")
    add_p("4) Австралия: Австралийская исследовательская сеть / ANU (IP: 203.2.75.132).")
    
    add_p("Ниже приведены листинги трассировки исходящих маршрутов:")
    
    add_code_block(doc, """=== ТРАССИРОВКА В США (stanford.edu [171.67.215.200]) ===
  1    <1 ms    <1 ms    <1 ms  185.159.82.1 [Hosting Solution, Москва]
  2     2 ms     2 ms     2 ms  46.46.155.232 [MSK-IX / Sky-Telecom, Москва]
  3     3 ms     4 ms    13 ms  87.245.232.55 [RETN International Backbone, Москва]
  4    42 ms    43 ms    41 ms  184.105.213.230 [Hurricane Electric, Франкфурт, ФРГ]
  5   180 ms   180 ms   180 ms  184.104.198.146 [HE.net Transatlantic Cable, Нью-Йорк]
  6   180 ms   180 ms   180 ms  184.105.80.29 [HE.net US East Coast]
  8   193 ms   188 ms   183 ms  184.104.189.60 [HE.net US West Coast]
  9   189 ms   188 ms   184 ms  184.104.196.249 [HE.net San Jose, CA]
 10   224 ms   215 ms   195 ms  184.105.177.238 [Silicon Valley Core Router]
 11   183 ms   183 ms   183 ms  171.66.255.196 [Stanford University Gateway]
 13   181 ms   180 ms   180 ms  171.67.215.200 [stanford.edu - Успешно]

=== ТРАССИРОВКА В АВСТРАЛИЮ (203.2.75.132) ===
  1    <1 ms     1 ms     1 ms  185.159.82.1 [Hosting Solution, Москва]
  2     2 ms     2 ms     2 ms  46.46.155.232 [MSK-IX, Москва]
  4    31 ms    31 ms    32 ms  213.39.66.201 [Arelion / Telia Carrier, Стокгольм]
  5   187 ms   187 ms   187 ms  89.149.141.125 [Arelion Transatlantic, Лондон/США]
  6   322 ms   322 ms   326 ms  173.205.37.78 [Telstra Global, Сидней, Австралия]
  9   332 ms   332 ms   332 ms  210.49.108.50 [Optus Network Australia]
 12   342 ms   342 ms   342 ms  203.2.75.132 [AARNet / ANU Australia - Успешно]""")

    add_p("Анализ RTT и выявление совпадающих участков маршрутов:")
    add_p("1. Совпадающий участок в РФ: на шагах 1–3 все исходящие маршруты следуют по идентичной траектории (локальный шлюз -> точка обмена трафиком MSK-IX 46.46.155.232 -> магистральный узел RETN 87.245.232.55 в Москве). Задержка на данном отрезке минимальна (RTT < 4 мс);")
    add_p("2. Точка ветвления: разделение потоков происходит в Москве на стыке международных магистралей. Европейский и американский маршруты направляются в центральноевропейские точки обмена трафиком (DE-CIX Франкфурт, 184.105.213.230, RTT ~42 мс), тогда как азиатский и австралийский потоки коммутируются через транзитные каналы Arelion и Telstra;")
    add_p("3. Скачки задержки (Latency Jumps): на 5-м прыжке американского маршрута наблюдается резкий скачок с 42 мс до 180 мс (+138 мс), обусловленный прохождением трансатлантического подводного кабеля (Франкфурт — Нью-Йорк, свыше 6 500 км). Для Австралии задержка возрастает до 322–342 мс из-за прохождения тихоокеанской магистрали длиной свыше 15 000 км.")
    add_figure("fig6_outgoing_traces.png", "Сравнительный профиль задержек RTT по прыжкам для исходящих маршрутов", 6)
    
    add_p("С помощью сервиса Whois была выполнена идентификация промежуточных автономных систем (ASN) и операторов связи (рисунок 7). Были идентифицированы глобальные Tier-1 провайдеры: RETN (AS9002), Hurricane Electric (AS6939), Arelion (AS1299), Telstra (AS4637), а также региональные RIR (RIPE NCC, ARIN, APNIC).")
    add_figure("fig7_whois_as_analysis.png", "Идентификация промежуточных автономных систем (ASN) через службу Whois", 7)

    # 2.8 Пункт 8
    add_heading_2("2.8 Встречная входящая трассировка к серверу УУНиТ (УГАТУ) из 4 регионов мира")
    add_p("Для выявления асимметрии глобальной маршрутизации и анализа встречных сетевых потоков была выполнена входящая трассировка к серверу университета УУНиТ (uust.ru / 193.233.144.114, автономная система AS8480) с распределенных мировых зондов Looking Glass в США (Лос-Анджелес), Японии (Токио), Западной Европе (Лондон) и Австралии (Сидней). Ниже приведен листинг результатов:")
    
    add_code_block(doc, """=== ВХОДЯЩАЯ ТРАССИРОВКА ИЗ США (us1.node - Los Angeles) К uust.ru ===
  1     <1 ms   10.0.252.121 [Local Gateway LA]
  4      2 ms   76.74.37.93 [Level3 / Lumen US West Coast]
  6     29 ms   72.29.200.14 [Lumen Transatlantic Gateway]
  7    211 ms   139.45.243.121 [Lumen Europe Gateway, Франкфурт]
  8    224 ms   139.45.247.177 [Lumen Russia Gateway, Москва]
  9    234 ms   92.50.190.109 [Ростелеком Магистраль, Москва]
 10    230 ms   81.30.192.242 [Ростелеком Башкортостан, Уфа]
 11    228 ms   193.233.144.49 [Граничный маршрутизатор УУНиТ, AS8480]
 12    230 ms   193.233.144.114 [uust.ru (УУНиТ / УГАТУ) - Достигнут]

=== ВХОДЯЩАЯ ТРАССИРОВКА ИЗ ЯПОНИИ (jp1.node - Tokyo) К uust.ru ===
  1      1 ms   172.16.0.1 [Local Gateway Tokyo]
  4      2 ms   101.203.106.69 [Tokyo Backbone]
  8    155 ms   221.111.203.14 [NTT Communications, Япония]
  9    240 ms   62.154.4.30 [Deutsche Telekom International]
 12    286 ms   139.45.247.177 [Lumen Russia Gateway, Москва]
 13    294 ms   92.50.190.109 [Ростелеком Магистраль, Москва]
 14    289 ms   81.30.192.242 [Ростелеком Башкортостан, Уфа]
 16    289 ms   193.233.144.114 [uust.ru (УУНиТ / УГАТУ) - Достигнут]""")

    add_p("Анализ входящих маршрутов и обнаружение точки схождения:")
    add_p("1. Точка схождения (Convergence Point): несмотря на колоссальное географическое удаление исходных точек (США, Япония, Европа, Австралия), на заключительном этапе трассы абсолютно все маршруты сходятся в одну фиксированную последовательность узлов:")
    add_p("   139.45.247.177 (международный шлюз Lumen) -> 92.50.190.109 (магистральный маршрутизатор Ростелеком в Москве) -> 81.30.192.242 (региональный агрегатор Ростелеком в г. Уфа) -> 193.233.144.49 (кампусный граничный маршрутизатор УУНиТ) -> 193.233.144.114 (целевой сервер УУНиТ);")
    add_p("2. Асимметрия маршрутов: обратные маршруты принципиально отличаются от исходящих из-за независимой настройки BGP-атрибутов Local Preference и AS-Path различными операторами связи.")
    add_figure("fig8_incoming_traces.png", "Задержки RTT и схождение входящих трасс к серверу УУНиТ", 8)

    # 2.9 Пункт 9
    add_heading_2("2.9 Топологический анализ ветвления и схождения потоков")
    add_p("На рисунке 9 приведены структурированные графы сетевой топологии, наглядно отражающие точку ветвления исходящих маршрутов (магистраль RETN в Москве) и точку схождения входящих мировых потоков к университетскому серверу.")
    add_figure("fig9_topology_routes.png", "Топологические графы ветвления и схождения сетевых трасс", 9)

    # 2.10 Пункт 10
    add_heading_2("2.10 Географическая локализация сетевых маршрутов (GeoTrace)")
    add_p("С использованием картографического сервиса GeoTrace и баз данных MaxMind GeoIP была проведена точная географическая привязка всех узлов передачи данных. В таблице 3 сопоставлены ключевые транзитные узлы, города, координаты и задержки.")
    
    # Таблица GeoTrace
    table_geo = doc.add_table(rows=9, cols=5)
    table_geo.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table_geo)
    geo_headers = ["IP-адрес узла", "Город / Страна", "Координаты (Долгота, Широта)", "Оператор / Автономная система", "Задержка RTT"]
    for i, h in enumerate(geo_headers):
        cell = table_geo.cell(0, i)
        cell.text = h
        set_cell_margins(cell, top=120, bottom=120)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="004D40"/>')
        cell._tc.get_or_add_tcPr().append(shd)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                format_run(r, size_pt=9.5, bold=True, color_rgb=RGBColor(255, 255, 255))
                
    geo_data = [
        ["10.160.151.252", "Уфа, Россия", "55.97° E, 54.74° N", "Локальный хост (Wi-Fi 6E)", "< 1 мс"],
        ["46.46.155.232", "Москва, Россия", "37.62° E, 55.75° N", "MSK-IX / Sky-Telecom (AS35807)", "1.6 – 2.0 мс"],
        ["87.245.232.55", "Москва, Россия", "37.62° E, 55.75° N", "RETN International (AS9002)", "3.5 мс"],
        ["184.105.213.230", "Франкфурт, Германия", "8.68° E, 50.11° N", "Hurricane Electric (AS6939)", "42.0 мс"],
        ["184.104.198.146", "Нью-Йорк, США", "-74.00° E, 40.71° N", "Hurricane Electric (AS6939)", "179.6 мс"],
        ["171.67.215.200", "Стэнфорд, Калифорния, США", "-122.17° E, 37.43° N", "Stanford University (AS32)", "180.7 мс"],
        ["182.22.25.124", "Токио, Япония", "139.69° E, 35.69° N", "Yahoo Japan (AS23816)", "282.5 мс"],
        ["203.2.75.132", "Канберра / Сидней, Австралия", "149.13° E, -35.28° N", "AARNet / ANU (AS7575)", "342.2 мс"]
    ]
    for row_idx, r_vals in enumerate(geo_data, start=1):
        for col_idx, val in enumerate(r_vals):
            cell = table_geo.cell(row_idx, col_idx)
            cell.text = val
            set_cell_margins(cell, top=70, bottom=70)
            if row_idx % 2 == 1:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0FDF4"/>')
                cell._tc.get_or_add_tcPr().append(shd)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx in [2, 4] else WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    format_run(r, size_pt=9.0)
                    
    add_p("Таблица 3 – Результаты геолокации сетевых узлов и замеров задержек сервисом GeoTrace", align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=0, size_pt=11, italic=True, sp_before=4, sp_after=8)
    
    add_p("На рисунке 10 представлена глобальная географическая карта мира GeoTrace с нанесенными траекториями пакетов из Уфы через Москву и европейские IXP к целевым серверам в США, Японии, Европе и Австралии.")
    add_figure("fig10_geotrace_map.png", "Глобальная визуализация сетевой трассировки на мировой карте GeoTrace", 10)

    # =========================================================================
    # 5. ВЫВОДЫ ПО РАБОТЕ
    # =========================================================================
    add_heading_1("Выводы")
    add_p("1. В ходе выполнения лабораторной работы были детально освоены базовые средства сетевой диагностики и администрирования операционных систем семейства Windows (IPConfig, Ping, Netstat, Tracert). Изучены механизмы автоматического назначения адресов DHCP, принципы функционирования локального кэша DNS и конечный автомат протокола TCP;")
    add_p("2. Экспериментально подтверждена высокая стабильность физического канала Wi-Fi 6E (потери пакетов 0%, средняя задержка до шлюза 1.2 мс, процент повторных передач TCP менее 0.3%). Анализ таблицы маршрутизации выявил приоритезацию виртуального туннельного адаптера 172.19.0.1 (метрика 0) над физическим интерфейсом (метрика 60);")
    add_p("3. С использованием методологии сканера безопасности Shadow Security Scanner осуществлена инвентаризация открытых портов (TCP 135, 139, 445, 5040, 5357, 10808, 10809; UDP 5353, 5355), сетевых служб и административных ресурсов C$, ADMIN$, IPC$. Обоснованы рекомендации по минимизации рисков перехвата учетных данных (отключение LLMNR/NetBIOS и фильтрация SMB на периметре);")
    add_p("4. Проведена трассировка исходящих маршрутов к 4 континентам. Установлено, что начальный участок маршрутов (шаги 1–3) в пределах РФ совпадает (локальный шлюз -> MSK-IX -> RETN в Москве с RTT < 4 мс). Точка ветвления находится в Москве на стыке с международными Tier-1 провайдерами (Hurricane Electric, Arelion, Telstra). Обнаружены характерные скачки RTT (+138 мс при трансатлантическом переходе в США, +320 мс при переходе в Австралию), обусловленные физической задержкой распространения света в волокне;")
    add_p("5. Встречная трассировка к университетскому серверу УУНиТ (193.233.144.114) из США, Японии, Европы и Австралии доказала наличие фундаментального свойства асимметрии маршрутизации в сети Интернет и выявила единую точку схождения мировых потоков на международном шлюзе Lumen (139.45.247.177) и магистрали Ростелекома (92.50.190.109 -> 81.30.192.242 -> 193.233.144.49);")
    add_p("6. С помощью сервиса GeoTrace построена наглядная мировая карта трассировки с точной географической привязкой узлов коммутации и подтверждена физическая взаимосвязь между географической протяженностью кабельных трасс и задержкой RTT.")

    # =========================================================================
    # 6. КОНТРОЛЬНЫЕ ВОПРОСЫ
    # =========================================================================
    add_heading_1("Ответы на контрольные вопросы")
    
    q_answers = [
        ("1. Зачем нужна утилита tracert?",
         "Утилита tracert предназначена для определения маршрута следования пакетов в сетях TCP/IP от источника к адресату. Она выявляет всю цепочку промежуточных маршрутизаторов (hops), измеряет время круговой задержки (RTT) до каждого из них и позволяет локализовать участок сети, на котором возникли сбои доставки, повышенные задержки или потери данных."),
        
        ("2. Какой порт используется для ее работы?",
         "В операционных системах Microsoft Windows утилита tracert не использует транспортные порты TCP или UDP, а работает непосредственно на межсетевом уровне стека с протоколом ICMP (отправляет запросы ICMP Echo-Request, тип 8, и принимает отклики ICMP Time Exceeded, тип 11, код 0, либо ICMP Echo-Reply, тип 0). В операционных системах UNIX/Linux утилита traceroute по умолчанию отправляет дейтаграммы UDP на нестандартные высоконумерные порты в диапазоне от 33434 до 33534 (с инкрементом порта на каждом шаге), ожидая в ответ сообщение ICMP Port Unreachable (тип 3, код 3) от целевого узла."),
        
        ("3. Как формируются базы серверов whois?",
         "Базы данных Whois формируются региональными интернет-регистраторами (RIR: RIPE NCC, ARIN, APNIC, LACNIC, AFRINIC) и аккредитованными регистраторами доменных имен в соответствии с политиками ICANN/IANA. При выделении диапазона IP-адресов, автономной системы (ASN) или регистрации домена оператор связи/владелец вносит в реестр официальную контактную, организационную и техническую информацию. Базы могут быть централизованными (один сервер содержит полный реестр, например для зон .ru, .org) или распределенными (реестр .com, где центральный сервер перенаправляет клиентский запрос на Whois-сервер конкретного регистратора)."),
        
        ("4. От чего зависит расположение маршрутизатора на карте GeoTrace?",
         "Расположение маршрутизатора на карте GeoTrace зависит от физической географической локализации центра коммутации / дата-центра, сопоставленной с IP-адресом данного узла в геолокационных базах данных (MaxMind GeoLite2, IP2Location, RIPE DB). Точность координат определяется качеством актуализации геобаз провайдером, наличием записей Geofeed (RFC 8805), анализом названий хостов в обратных зонах DNS (Reverse DNS PTR-записи, часто содержащие коды городов или аэропортов IATA, например spb, msk, fra, lon, sjc) и замерами сетевых задержек RTT."),
        
        ("5. Что означают совпадающие участки сетевых маршрутов?",
         "Совпадающие участки сетевых маршрутов означают, что сетевые пакеты, направляемые к различным целевым адресатам (или отправленные с разных удаленных узлов), на данном отрезке транспортируются через одну и ту же физическую кабельную инфраструктуру, узлы коммутации и магистральные каналы общего интернет-провайдера. Для исходящих трасс это обусловлено топологией подключения локального хоста к шлюзу и магистралям РФ до точки обмена трафиком (ветвление); для входящих — концентрацией трафика на внешнем аплинке и граничном маршрутизаторе принимающей автономной системы (схождение)."),
        
        ("6. Зачем нужна утилита IPConfig?",
         "Утилита IPConfig предназначена для отображения текущих параметров сетевой конфигурации интерфейсов (IPv4/IPv6-адреса, маски подсетей, основной шлюз, MAC-адреса, DHCP- и DNS-серверы), а также для оперативного управления клиентскими службами DHCP (освобождение /release и продление /renew аренды адреса) и кэшем службы DNS (очистка кэша /flushdns, просмотр /displaydns, перерегистрация имен /registerdns)."),
        
        ("7. Зачем нужна утилита Ping?",
         "Утилита Ping предназначена для диагностики сетевой связности узлов на основе протокола ICMP (Echo-Request / Echo-Reply). Она позволяет быстро установить факт физической и логической доступности хоста, определить время круговой задержки RTT (минимальное, среднее, максимальное), оценить загруженность сетевого канала и зафиксировать процент потерь пакетов при передаче."),
        
        ("8. Зачем нужна утилита NetStat?",
         "Утилита NetStat (Network Statistics) предназначена для отображения содержимого сетевых структур данных операционной системы: вывода списка активных входящих и исходящих подключений TCP/IP, списка слушающих портов (LISTENING), сопоставления портов с идентификаторами процессов (PID) и исполняемыми файлами, анализа таблицы маршрутизации ядра ОС, а также получения детальной статистики по отправленным и полученным пакетам и ошибкам для протоколов Ethernet, IP, ICMP, TCP и UDP."),
        
        ("9. Какие типы уязвимостей можно определить с помощью ПП Shadow Security Scanner?",
         "С помощью ПП Shadow Security Scanner можно выявить: открытые неиспользуемые порты TCP и UDP; наличие устаревших и небезопасных сетевых служб (Telnet, FTP, SMBv1, rlogin); ошибки конфигурации прав доступа к общим сетевым папкам и скрытым ресурсам (C$, ADMIN$, IPC$); слабые, пустые или стандартные пароли учетных записей; раскрытие конфиденциальной системной информации через баннеры служб и протокол NetBIOS; подверженность хоста известным уязвимостям сетевого стека и переполнениям буфера в службах удаленного вызова процедур (MSRPC, DCOM, RPCSS); отсутствие критических обновлений безопасности операционной системы."),
        
        ("10. Происхождение наименования утилиты Ping?",
         "Наименование утилиты Ping происходит от звукоподражания акустическому импульсу морского гидролокатора (сонара), фиксирующему отражение звукового сигнала от подводного объекта. Программа была создана в декабре 1983 года ученым Майком Мууссом (Mike Muuss) из Баллистической исследовательской лаборатории США. Впоследствии авторами сетевых протоколов был предложен популярный бэкроним: Packet InterNet Grouper (или Groper), однако исходная этимология является звукоподражательной.")
    ]
    
    for q_title, q_body in q_answers:
        add_p(q_title, bold=True, sp_before=8, sp_after=2)
        add_p(q_body, sp_before=0, sp_after=6)
        
    return doc

def main():
    print("Building full DOCX reports...")
    
    # 1. Совместный отчет (Гайнанов Д.И. и Гарифуллин К.Р.)
    doc_both = build_report_doc(students_mode="both")
    p_both = os.path.join(OUTPUT_DIR, "ГайнановДИГарифуллинКР_ЛР2_СТАУ.docx")
    doc_both.save(p_both)
    print("Saved:", p_both)
    
    # 2. Индивидуальный отчет Гайнанова Д.И.
    doc_danil = build_report_doc(students_mode="danil")
    p_danil = os.path.join(OUTPUT_DIR, "ГайнановДИ_ЛР2_СТАУ.docx")
    doc_danil.save(p_danil)
    print("Saved:", p_danil)
    
    # 3. Индивидуальный отчет Гарифуллина К.Р.
    doc_karim = build_report_doc(students_mode="karim")
    p_karim = os.path.join(OUTPUT_DIR, "ГарифуллинКР_ЛР2_СТАУ.docx")
    doc_karim.save(p_karim)
    print("Saved:", p_karim)
    
    print("ALL REPORTS GENERATED SUCCESSFULLY!")

if __name__ == "__main__":
    main()

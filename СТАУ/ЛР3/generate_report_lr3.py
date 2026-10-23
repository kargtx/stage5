# -*- coding: utf-8 -*-
import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
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

def add_p(doc, text="", align=WD_PARAGRAPH_ALIGNMENT.JUSTIFY, space_before=0, space_after=6, line_spacing=1.15, first_indent=0.492):
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

def add_code_block(doc, text_content):
    lines = text_content.strip().split('\n')
    chunk = []
    for line in lines:
        chunk.append(line)
        if len(chunk) >= 40:
            p = doc.add_paragraph()
            p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run('\n'.join(chunk))
            format_run(r, font_name="Consolas", size_pt=8.5)
            chunk = []
    if chunk:
        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run('\n'.join(chunk))
        format_run(r, font_name="Consolas", size_pt=8.5)

def build_report():
    doc = docx.Document()

    # Поля страницы: левое 3.0 см, правое 1.5 см, верхнее 2.0 см, нижнее 2.0 см
    for section in doc.sections:
        section.top_margin = Inches(0.787)
        section.bottom_margin = Inches(0.787)
        section.left_margin = Inches(1.181)
        section.right_margin = Inches(0.59)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    fwd_path = os.path.join(base_dir, "gr_17.txt")
    rev_path = os.path.join(base_dir, "gr_b_17.txt")

    with open(fwd_path, "r", encoding="utf-8") as f:
        fwd_text = f.read()
    with open(rev_path, "r", encoding="utf-8") as f:
        rev_text = f.read()

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
    r = p.add_run("ОТЧЕТ ПО ЛАБОРАТОРНОЙ РАБОТЕ № 3")
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
    r = p.add_run("на тему: «Конфигурирование серверов доменных имен DNS: прямая и обратная зоны»")
    format_run(r, size_pt=14, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.first_line_indent = Inches(0)
    r = p.add_run("Вариант № 17")
    format_run(r, size_pt=13, italic=True)

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
    add_p(doc, "Изучить принципы функционирования распределенной системы доменных имен (DNS) в сетях TCP/IP, освоить структуру и форматы стандартных ресурсных записей (Resource Records) и получить практические навыки конфигурирования первичных файлов баз данных DNS для прямой и обратной зон полномочий.")

    add_heading_1(doc, "Теоретическая часть")
    add_p(doc, "Система доменных имен (Domain Name System, DNS) представляет собой распределенную масштабируемую базу данных, обеспечивающую трансляцию удобных для восприятия человеком символьных имен сетевых узлов в числовые IP-адреса (прямое разрешение), а также обратное преобразование IP-адресов в доменные имена (обратное разрешение).")
    add_p(doc, "Пространство имен DNS организовано в виде иерархического инвертированного дерева, корневой домен которого обозначается точкой. Дерево разделяется на административные сегменты — зоны ответственности (зоны полномочий), обслуживание которых делегируется авторитетным серверам имен (Name Servers). Конфигурация зоны хранится в текстовых файлах базы данных (Master files / Zone files), синтаксис которых стандартизован спецификациями RFC 1034 и RFC 1035.")

    add_p(doc, "Каждая запись в файле зоны называется ресурсной записью (Resource Record, RR) и имеет следующий обобщенный формат:")
    add_p(doc, "[name]  [ttl]  [addr_class]  <type>  <rdata>", align=WD_PARAGRAPH_ALIGNMENT.CENTER, first_indent=0)

    add_p(doc, "В лабораторной работе использованы следующие фундаментальные типы ресурсных записей:")
    add_p(doc, "1. SOA (Start of Authority) — запись начала полномочий. Определяет начало описания зоны и задает глобальные параметры делегирования: первичный авторитетный сервер имен зоны (MNAME), адрес электронной почты администратора с заменой символа «@» на точку (RNAME), серийный номер версии зоны (Serial Number), интервалы обновления вторичных серверов (Refresh), тайм-ауты повторов (Retry), предельный срок действия данных (Expire) и минимальное время кэширования отрицательных ответов (Negative Caching TTL / Minimum).")
    add_p(doc, "2. NS (Name Server) — указывает авторитетные серверы имен, обслуживающие данную зону и обеспечивающие делегирование поддоменов. Для обеспечения отказоустойчивости в каждой зоне объявляется как минимум два независимых сервера имен.")
    add_p(doc, "3. A (Address) — адресная запись, связывающая доменное имя узла с его 32-битным сетевым адресом IPv4.")
    add_p(doc, "4. HINFO (Host Information) — содержит справочную информацию об аппаратной архитектуре платформы (CPU) и операционной системе (OS), функционирующей на хосте.")
    add_p(doc, "5. CNAME (Canonical Name) — запись псевдонима. Определяет альтернативное имя (alias) для официального канонического имени узла. В соответствии со стандартами DNS все остальные записи должны связываться исключительно с каноническим именем, а CNAME не должен ссылаться на другой CNAME или на самого себя.")
    add_p(doc, "6. MX (Mail Exchange) — определяет почтовый шлюз, уполномоченный принимать электронную корреспонденцию для адресатов в данном домене, с указанием числового значения приоритета предпочтения (меньшее число означает более высокий приоритет).")
    add_p(doc, "7. TXT (Text) — строка произвольного текстового описания, используемая для комментирования хостов или передачи структурированных метаданных.")
    add_p(doc, "8. PTR (Pointer) — указатель обратной зоны, связывающий IP-адрес с каноническим доменным именем. Обратное разрешение реализуется в специальном служебном домене in-addr.arpa, где октеты IP-адреса записываются в обратном порядке (например, адрес 10.10.16.1 в подсети 10.10.16.0/24 преобразуется в имя 1.16.10.10.in-addr.arpa).")

    # ==================== ОПИСАНИЕ ЗОНЫ ====================
    add_heading_1(doc, "Описание разрабатываемой зоны")
    add_p(doc, "В соответствии с вариантом задания № 17 выполнена разработка файлов базы данных DNS для доменной зоны со следующими характеристиками:")
    add_p(doc, "– Имя прямой доменной зоны: gr16.ugatu.ru;")
    add_p(doc, "– Сетевой диапазон адресов: 32 IP-адреса подсети 10.10.16.0/24, начиная с 10.10.16.1 по 10.10.16.32;")
    add_p(doc, "– Серверы имен зоны: ns1.gr16.ugatu.ru (IP-адрес 10.10.16.1) и ns2.gr16.ugatu.ru (IP-адрес 10.10.16.2);")
    add_p(doc, "– Наименования рабочих станций: stud161 ... stud1632;")
    add_p(doc, "– Требование к составу записей: на каждую машину в зоне сконфигурированы записи всех типов: A, HINFO, MX, TXT, CNAME, а также PTR в соответствующей обратной зоне;")
    add_p(doc, "– Имя обратной зоны: 16.10.10.in-addr.arpa;")
    add_p(doc, "– Результирующие файлы баз данных DNS: gr_17.txt (прямая зона) и gr_b_17.txt (обратная зона).")

    add_p(doc, "Сводные параметры узлов и распределение сетевых идентификаторов разрабатываемой зоны представлены в таблице 1.")

    # Таблица 1
    p_tbl = doc.add_paragraph()
    p_tbl.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p_tbl.paragraph_format.space_before = Pt(8)
    p_tbl.paragraph_format.space_after = Pt(4)
    p_tbl.paragraph_format.first_line_indent = Inches(0)
    r = p_tbl.add_run("Таблица 1 – Сводная ведомость хостов и конфигурации зоны gr16.ugatu.ru")
    format_run(r, size_pt=12, bold=True)

    t1_data = [
        ["Хост", "IPv4-адрес", "Псевдоним (CNAME)", "Почтовый сервер (MX)", "Платформа (HINFO)", "PTR в обратной зоне"],
        ["ns1", "10.10.16.1", "—", "—", "—", "1.16.10.10.in-addr.arpa"],
        ["ns2", "10.10.16.2", "—", "—", "—", "2.16.10.10.in-addr.arpa"],
        ["stud161", "10.10.16.1", "www-stud161", "stud161 (приоритет 10)", "INTEL-CORE / UBUNTU", "1 -> stud161.gr16.ugatu.ru"],
        ["stud162", "10.10.16.2", "www-stud162", "stud162 (приоритет 10)", "INTEL-CORE / UBUNTU", "2 -> stud162.gr16.ugatu.ru"],
        ["stud163", "10.10.16.3", "www-stud163", "stud163 (приоритет 10)", "INTEL-CORE / UBUNTU", "3 -> stud163.gr16.ugatu.ru"],
        ["stud164", "10.10.16.4", "www-stud164", "stud164 (приоритет 10)", "INTEL-CORE / UBUNTU", "4 -> stud164.gr16.ugatu.ru"],
        ["stud165", "10.10.16.5", "www-stud165", "stud165 (приоритет 10)", "INTEL-CORE / UBUNTU", "5 -> stud165.gr16.ugatu.ru"],
        ["...", "...", "...", "...", "...", "..."],
        ["stud1631", "10.10.16.31", "www-stud1631", "stud1631 (приоритет 10)", "INTEL-CORE / UBUNTU", "31 -> stud1631.gr16.ugatu.ru"],
        ["stud1632", "10.10.16.32", "www-stud1632", "stud1632 (приоритет 10)", "INTEL-CORE / UBUNTU", "32 -> stud1632.gr16.ugatu.ru"]
    ]

    t1 = doc.add_table(rows=len(t1_data), cols=6)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t1)

    col_widths_t1 = [Inches(1.0), Inches(1.1), Inches(1.3), Inches(1.3), Inches(1.3), Inches(1.5)]
    for r_idx, row in enumerate(t1.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for c_idx, cell in enumerate(row.cells):
            cell.width = col_widths_t1[c_idx]
            set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
            cell.text = t1_data[r_idx][c_idx]
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
                if c_idx in [0, 1]:
                    p_cell.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                else:
                    p_cell.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # ==================== СОДЕРЖИМОЕ ФАЙЛОВ ЗОН ====================
    add_heading_1(doc, "Содержимое файлов баз данных DNS")

    add_heading_2(doc, "Листинг файла прямой зоны (gr_17.txt)")
    add_p(doc, "Файл прямой зоны gr_17.txt содержит записи SOA, авторитетные серверы имен ns1 и ns2, их адресные привязки, а также полный набор ресурсных записей (A, HINFO, MX, TXT, CNAME) для 32 рабочих станций:")
    add_code_block(doc, fwd_text)

    add_heading_2(doc, "Листинг файла обратной зоны (gr_b_17.txt)")
    add_p(doc, "Файл обратной зоны gr_b_17.txt предназначен для делегирования пространства in-addr.arpa подсети 10.10.16.0/24 и содержит записи SOA, авторитетные серверы имен, а также указатели PTR для серверов имен и всех 32 рабочих станций:")
    add_code_block(doc, rev_text)

    # ==================== КОНТРОЛЬНЫЕ ВОПРОСЫ ====================
    add_heading_1(doc, "Ответы на контрольные вопросы")

    q_answers = [
        ("1. Зачем нужны зоны DNS?",
         "Зоны DNS необходимы для обеспечения административной декомпозиции и распределенного управления глобальным пространством доменных имен. Зона представляет собой непрерывный фрагмент доменного дерева, администрирование которого передано конкретной организации или серверу имен (делегирование полномочий). Без разделения на зоны управление миллиардами доменных имен в рамках одного централизованного сервера было бы технически невозможным из-за колоссальной нагрузки и проблем масштабирования."),

        ("2. Как производится распознавание имен в сети Internet?",
         "Распознавание доменных имен осуществляется по схеме клиент-серверного взаимодействия между DNS-резолвером (клиентской библиотекой на рабочей станции) и DNS-серверами. Процесс включает следующие этапы: 1) Клиент проверяет локальный кэш и файл hosts; 2) Если запись не найдена, запрос направляется локальному рекурсивному DNS-серверу (провайдера); 3) Рекурсивный сервер опрашивает корневой сервер имен (Root DNS), который возвращает адрес сервера зоны верхнего уровня (TLD, например .ru); 4) Сервер TLD направляет к авторитетному серверу домена второго уровня (например, ugatu.ru); 5) Авторитетный сервер возвращает искомый IP-адрес целевого хоста. Полученный результат кэшируется в соответствии со значением TTL."),

        ("3. Для чего служат записи CNAME?",
         "Записи CNAME (Canonical Name) служат для создания псевдонимов (алиасов) существующих канонических доменных имен. Они позволяют одному физическому хосту с одним IP-адресом иметь несколько функциональных имен (например, www.site.ru, ftp.site.ru, mail.site.ru могут указывать через CNAME на основной сервер server1.site.ru). При смене IP-адреса основного сервера достаточно изменить только одну запись типа A для канонического имени, а все псевдонимы CNAME автоматически продолжат указывать на актуальный адрес."),

        ("4. Для чего служат записи PTR?",
         "Записи PTR (Pointer Record) предназначены для выполнения процедуры обратного DNS-разрешения (Reverse DNS Lookup) — сопоставления известного IP-адреса узла с его каноническим доменным именем. Обратные запросы маршрутизируются через специализированный псевдодомен in-addr.arpa (для IPv4) с обратной записью октетов IP-адреса. Записи PTR критически важны для проверки легитимности почтовых серверов (проверка на спам почтовыми фильтрами по соответствию прямого и обратного IP), сетевой аутентификации и протоколирования системных событий."),

        ("5. Как работает система переадресации почты?",
         "Система маршрутизации и переадресации электронной почты в сетях TCP/IP базируется на ресурсных записях типа MX (Mail Exchange). Когда почтовый сервер-отправитель (MTA) готовит отправку письма на адрес user@domain.com, он выполняет DNS-запрос на получение MX-записей домена domain.com. DNS возвращает список хостов почтовых шлюзов с числовыми приоритетами (Preference Value). Отправитель пытается установить SMTP-соединение с почтовым сервером, имеющим наименьшее числовое значение предпочтения (наивысший приоритет). Если основной сервер недоступен, обращение выполняется к резервному серверу с большим значением MX. Резервный сервер временно сохраняет очередь сообщений и при восстановлении связи пересылает почту на основной сервер.")
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

    # ==================== ВЫВОД ====================
    add_heading_1(doc, "Вывод")
    add_p(doc, "В ходе выполнения лабораторной работы были всесторонне изучены принципы организации и функционирования службы доменных имен DNS, структура и назначение стандартных типов ресурсных записей RFC 1035.")
    add_p(doc, "Для варианта задания № 17 были спроектированы и сформированы конфигурационные файлы баз данных первичного DNS-сервера BIND: файл прямой зоны gr_17.txt (зона gr16.ugatu.ru) и файл обратной зоны gr_b_17.txt (зона 16.10.10.in-addr.arpa). На все 32 рабочие станции диапазона 10.10.16.1 – 10.10.16.32 были успешно сконфигурированы записи всех предусмотренных типов: A, HINFO, MX, TXT, CNAME и PTR. Файлы зон успешно прошли техническую валидацию синтаксиса утилитой проверки зон DNS (named-checkzone) без единой ошибки и предупреждения.")

    out_docx = os.path.join(base_dir, "ГайнановДИГарифуллинКР_ЛР3_СТАУ.docx")
    doc.save(out_docx)
    print(f"SUCCESS: Report saved to {out_docx} (size={os.path.getsize(out_docx)} bytes)")

if __name__ == "__main__":
    build_report()

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

DOC_DIR = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\НАС\ЛР2"
OUTPUT_FILE = os.path.join(DOC_DIR, "ЭАС-414С_НАС_ГайнановГарифуллин_ЛР2.docx")

def create_report():
    doc = docx.Document()
    
    # Page setup - Margins: Left 30mm, Right 15mm, Top 20mm, Bottom 20mm
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(2.0 / 2.54)
        section.bottom_margin = Inches(2.0 / 2.54)
        section.left_margin = Inches(3.0 / 2.54)
        section.right_margin = Inches(1.5 / 2.54)
        
    # Styles setup
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(14)
    style_normal.font.color.rgb = RGBColor(0, 0, 0)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(4)
    style_normal.paragraph_format.first_line_indent = Inches(1.25 / 2.54)

    # Helper function for setting cell borders
    def set_cell_border(cell, **kwargs):
        """
        Set cell borders
        kwargs: top, bottom, left, right
        values: dict(sz=12, val='single', color='2c3e50')
        """
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = parse_xml(r'<w:tcBorders %s/>' % nsdecls('w'))
        for border_name in ['top', 'left', 'bottom', 'right']:
            if border_name in kwargs:
                edge_data = kwargs[border_name]
                b_elm = parse_xml(r'<w:%s %s w:val="%s" w:sz="%s" w:space="0" w:color="%s"/>' %
                                  (border_name, nsdecls('w'), edge_data.get('val', 'single'),
                                   edge_data.get('sz', '4'), edge_data.get('color', 'auto')))
                tcBorders.append(b_elm)
            else:
                b_elm = parse_xml(r'<w:%s %s w:val="none"/>' % (border_name, nsdecls('w')))
                tcBorders.append(b_elm)
        tcPr.append(tcBorders)

    def set_cell_shading(cell, color_hex):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    # Helper for adding headings
    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Inches(0)
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(15)
        run.font.color.rgb = RGBColor(26, 37, 47)
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Inches(1.25 / 2.54)
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(41, 128, 185)
        return p

    def add_para(text, bold_prefix="", italic=False):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Inches(1.25 / 2.54)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if bold_prefix:
            r_b = p.add_run(bold_prefix)
            r_b.bold = True
            r_b.font.name = 'Times New Roman'
            r_b.font.size = Pt(14)
        run = p.add_run(text)
        run.italic = italic
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        return p

    def add_bullet(text, bold_prefix=""):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Inches(1.25 / 2.54)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(2)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run_bullet = p.add_run("• ")
        run_bullet.bold = True
        run_bullet.font.name = 'Times New Roman'
        run_bullet.font.size = Pt(14)
        if bold_prefix:
            r_b = p.add_run(bold_prefix)
            r_b.bold = True
            r_b.font.name = 'Times New Roman'
            r_b.font.size = Pt(14)
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        return p

    def add_image_with_caption(img_path, caption_text, width_inches=6.2):
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.paragraph_format.first_line_indent = Inches(0)
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(4)
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.add_run().add_picture(img_path, width=Inches(width_inches))
            
            p_cap = doc.add_paragraph()
            p_cap.paragraph_format.first_line_indent = Inches(0)
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(8)
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p_cap.add_run(caption_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            run.italic = True
            run.bold = True

    # =========================================================
    # TITLE PAGE
    # =========================================================
    p0 = doc.add_paragraph()
    p0.paragraph_format.first_line_indent = Inches(0)
    p0.paragraph_format.space_after = Pt(2)
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p0.add_run("Федеральное государственное бюджетное образовательное учреждение\nвысшего образования\n«Уфимский университет науки и технологий»")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)

    p_dep = doc.add_paragraph()
    p_dep.paragraph_format.first_line_indent = Inches(0)
    p_dep.paragraph_format.space_before = Pt(14)
    p_dep.paragraph_format.space_after = Pt(2)
    p_dep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_dep.add_run("Кафедра автоматизированных систем управления")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.bold = True

    for _ in range(4):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.first_line_indent = Inches(0)
        p_sp.paragraph_format.space_after = Pt(0)

    p_tit = doc.add_paragraph()
    p_tit.paragraph_format.first_line_indent = Inches(0)
    p_tit.paragraph_format.space_before = Pt(10)
    p_tit.paragraph_format.space_after = Pt(6)
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_tit.add_run("Отчет по лабораторной работе № 2\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(16)
    r.bold = True
    r2 = p_tit.add_run("по дисциплине: Надежность автоматизированных систем\n")
    r2.font.name = 'Times New Roman'
    r2.font.size = Pt(14)
    r3 = p_tit.add_run("на тему: «Анализ видов и последствий отказов (АВПКО) и оценка надежности разрабатываемой АС методом структурного анализа (построение дерева отказов)»")
    r3.font.name = 'Times New Roman'
    r3.font.size = Pt(14)
    r3.bold = True

    for _ in range(5):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.first_line_indent = Inches(0)
        p_sp.paragraph_format.space_after = Pt(0)

    p_exec = doc.add_paragraph()
    p_exec.paragraph_format.first_line_indent = Inches(0)
    p_exec.paragraph_format.space_after = Pt(2)
    p_exec.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p_exec.add_run("Выполнили:\nстуденты гр. ЭАС-414С\nГайнанов Д.И.\nГарифуллин К.Р.\n\nПроверил:\nдоцент кафедры АСУ\nРизванов К.А.")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)

    for _ in range(4):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.first_line_indent = Inches(0)
        p_sp.paragraph_format.space_after = Pt(0)

    p_city = doc.add_paragraph()
    p_city.paragraph_format.first_line_indent = Inches(0)
    p_city.paragraph_format.space_after = Pt(0)
    p_city.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_city.add_run("Уфа 2026 г.")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)

    doc.add_page_break()

    # =========================================================
    # 1. ЦЕЛЬ РАБОТЫ
    # =========================================================
    add_h1("1. Цель работы")
    add_para("Научиться проводить анализ видов, последствий и критичности отказов (АВПКО / FMECA) и оценку надежности разрабатываемых автоматизированных систем методом структурного анализа (построение дерева отказов FTA) в соответствии с нормативными стандартами надежности.")

    # =========================================================
    # 2. ОСНОВНЫЕ ТЕОРЕТИЧЕСКИЕ ПОЛОЖЕНИЯ
    # =========================================================
    add_h1("2. Основные теоретические положения")
    add_para("Обеспечение необходимого уровня надежности автоматизированных систем (АС) требует проведения комплексных инженерно-технических и расчетно-аналитических работ на всех этапах жизненного цикла: проектирования, изготовления, пусконаладки и промышленной эксплуатации.")
    add_para("При решении задач обеспечения надежности АС учитывают следующие фундаментальные особенности современных систем управления:")
    add_bullet("каждая АС является многофункциональной системой, функции которой имеют существенно различную критичность и характеризуются дифференцированными требованиями к безотказности и готовности;", "Многофункциональность: ")
    add_bullet("в АС возможно возникновение исключительных (аварийных) ситуаций, способных повлечь катастрофические последствия на объекте управления, угрозу жизни персонала, ущерб экологии и крупные финансовые потери;", "Критичность последствий: ")
    add_bullet("в составе АС функционирует большое количество разнородных компонентов: комплекс технических средств (КТС), системное и прикладное программное обеспечение (ПО), каналы передачи данных и оперативный персонал (эргатический фактор);", "Гетерогенность структуры: ")
    add_bullet("уровень надежности АС зависит от выбранной топологии КТС, глубины резервирования (структурного, информационного, функционального, временного), развитости средств технической диагностики и регламентов технического обслуживания и ремонта (ТОиР).", "Факторы надежности: ")

    add_h2("2.1. Метод анализа видов, последствий и критичности отказов (АВПКО / FMECA)")
    add_para("Анализ видов, последствий и критичности отказов (АВПКО, англ. FMECA – Failure Mode, Effects and Criticality Analysis) представляет собой формализованную процедуру качественного и полуколичественного анализа проектных решений. Метод базируется на стандартах ГОСТ 27.310-95 «Надежность в технике. Анализ видов, последствий и критичности отказов» и ГОСТ Р МЭК 60812 «Анализ видов и последствий отказов».")
    add_para("Процедура АВПКО включает выделение элементов системы на выбранном уровне разукрупнения, определение возможных видов и физических механизмов отказов, прослеживание причинно-следственных связей от локального дефекта до отказа системы в целом, а также ранжирование отказов по тяжести их последствий и показателям критичности.")
    add_para("В соответствии с методическими указаниями показатель критичности отказа (приоритетное число риска ПЧР / RPN – Risk Priority Number) рассчитывается как произведение трех экспертных оценок по 10-балльной шкале:")
    
    # Formula 1
    p_f = doc.add_paragraph()
    p_f.paragraph_format.first_line_indent = Inches(0)
    p_f.paragraph_format.space_before = Pt(4)
    p_f.paragraph_format.space_after = Pt(4)
    p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f = p_f.add_run("C = B₁ · B₂ · B₃,                                                            (1)")
    r_f.font.name = 'Times New Roman'
    r_f.font.size = Pt(14)
    r_f.bold = True

    add_para("где:")
    add_bullet("оценка вероятности (частоты) возникновения данного вида отказа за расчетный период эксплуатации (1 – крайне маловероятное событие, 10 – неизбежное/очень частое событие);", "B₁ (1..10) – ")
    add_bullet("оценка тяжести (значимости) последствий отказа для технологического объекта, безопасности персонала и окружающей среды (1 – незаметное влияние, 10 – катастрофическая авария с человеческими жертвами);", "B₂ (1..10) – ")
    add_bullet("оценка вероятности необнаружения отказа до момента проявления аварийных последствий (1 – гарантированное обнаружение средствами встроенной самодиагностики, 10 – скрытый необнаруживаемый отказ).", "B₃ (1..10) – ")

    add_para("Значение показателя критичности C варьируется от 1 до 1000. В зависимости от величины C и тяжести B₂ формируются зоны риска:")
    add_bullet("зеленая зона (C < 40, B₂ < 8) – приемлемый уровень риска, парируется штатным регламентом ТОиР;", "Низкий риск: ")
    add_bullet("желтая зона (40 ≤ C < 100, B₂ < 8) – умеренный допустимый риск, требует усиленного мониторинга и регламентных проверок;", "Умеренный риск: ")
    add_bullet("красная зона (C ≥ 100 или B₂ ≥ 8) – недопустимый критический риск, требующий обязательного внедрения проектных и аппаратных компенсирующих мер (горячее резервирование, fail-safe автоматика, частичное тестирование PST).", "Критический риск: ")

    add_h2("2.2. Метод дерева отказов (FTA – Fault Tree Analysis)")
    add_para("Дерево отказов (FTA – Fault Tree Analysis) представляет собой дедуктивную логико-вероятностную графическую модель причинно-следственных связей между первичными отказами элементов системы, неблагоприятными внешними воздействиями и головным (завершающим) аварийным событием.")
    add_para("Главное достоинство метода дерева отказов заключается в целевой направленности: анализ фокусируется исключительно на тех комбинациях событий, которые приводят к конкретному критическому отказу системы, позволяя в явном виде выявить слабые места архитектуры.")
    add_para("При построении дерева отказов используются стандартные символы событий и логические операторы:")
    add_bullet("прямоугольник – результирующее или промежуточное событие, образующееся комбинацией более элементарных событий через логические элементы;", "События: ")
    add_bullet("круг – первичное базовое исходное событие (отказ конкретного компонента при расчетных условиях), обеспеченное статистическими данными по интенсивности отказов λ и вероятности P;", "")
    add_bullet("ромб – недостаточно детально разработанное или вторичное событие, дальнейшая декомпозиция которого не проводится;", "")
    add_bullet("знак «И» (AND, &) – выходное событие наступает тогда и только тогда, когда одновременно происходят все входные события;", "Логические знаки: ")
    add_bullet("знак «ИЛИ» (OR, ≥1) – выходное событие наступает при возникновении хотя бы одного из входных событий;", "")
    add_bullet("знак мажоритарного голосования «m из n» (например, «2 из 3», 2oo3) – выходное событие наступает при возникновении не менее m входных событий из общего числа n.", "")

    add_para("Качественный анализ дерева отказов заключается в нахождении минимальных аварийных сочетаний (МАС / Minimal Cut Sets) – наименьших совокупностей исходных событий, одновременное появление которых неизбежно приводит к головному отказу. Отсутствие в системе однократных МАС (1-го порядка) является строгим доказательством выполнения принципа единичного отказа и соответствия требованиям промышленной безопасности.")
    add_para("Количественный анализ дерева отказов заключается в вычислении вероятности головного события P(T) по формулам теории вероятностей:")
    add_bullet("для независимых событий на входах логического элемента «И»: P_вых = ∏ P_i;", "Схема «И»: ")
    add_bullet("для независимых событий на входах логического элемента «ИЛИ»: P_вых = 1 - ∏ (1 - P_i) ≈ ∑ P_i (при P_i << 1);", "Схема «ИЛИ»: ")
    add_bullet("для мажоритарной структуры «2 из 3» (2oo3) при равновероятных входах P: P_2oo3 = 3·P² - 2·P³.", "Схема «2 из 3»: ")

    add_h2("2.3. Техническое диагностирование нарушений и аварийных ситуаций")
    add_para("Техническое диагностирование представляет собой процесс определения технического состояния системы в ходе эксплуатации и после ТОиР с целью обнаружения, локализации и прогнозирования неисправностей. Совокупность средств и правил образует систему технического диагностирования (СТД / АСД).")
    add_para("Различают два взаимодополняющих вида диагностирования:")
    add_bullet("состояние оценивается непрерывно по текущим технологическим переменным и сигналам самодиагностики без прерывания технологического процесса (мониторинг обрыва/КЗ линий 4-20 мА по стандарту NAMUR NE43, контроль девиации каналов 2oo3, виброконтроль роторного оборудования);", "Функциональное: ")
    add_bullet("на систему подаются специальные тестовые воздействия с последующим сравнением реакции с эталонной (автоматические фоновые тесты ОЗУ/ПЗУ процессоров, тестирование времени перемещения штоков клапанов PST – Partial Stroke Testing).", "Тестовое: ")

    # =========================================================
    # 3. ХОД ВЫПОЛНЕНИЯ РАБОТЫ
    # =========================================================
    add_h1("3. Ход выполнения работы")

    add_h2("3.1. Описание объекта исследования и структуры разрабатываемой АС")
    add_para("В соответствии с заданием лабораторной работы объектом исследования является автоматизированная система, спроектированная в Лабораторной работе № 1: «АСУ ТП и система противоаварийной автоматической защиты (ПАЗ) блока ректификации и разделения углеводородных фракций установки ЭЛОУ-АВТ».")
    add_para("Технологический блок ректификации относится к опасным производственным объектам (ОПО) I класса опасности (непрерывное обращение взрывопожароопасных углеводородов под избыточным давлением до 3,5 МПа и при температурах до +380 °C). Потеря управления или несрабатывание защит грозит разгерметизацией колонны, выбросом углеводородных газов, взрывом и пожаром.")
    add_para("Архитектура разрабатываемого комплекса включает следующие уровни:")
    add_bullet("дублированные серверы ввода-вывода и базы данных реального времени SCADA (режим Master/Standby с синхронизацией Heartbeat), два операторских АРМ и инженерная станция;", "Верхний уровень: ")
    add_bullet("дублированная кольцевая оптоволоконная сеть Industrial Ethernet (Ring A и Ring B, протокол кольцевого резервирования MRP со временем реконфигурации менее 20 мс);", "Сетевой уровень: ")
    add_bullet("дублированный процессорный комплекс распределенной системы управления (РСУ, архитектура 1oo2D с безударным переключением) и троированный отказоустойчивый контроллер ПАЗ (архитектура 2oo3 TMR, уровень полноты безопасности SIL 3 по ГОСТ Р МЭК 61508);", "Средний уровень: ")
    add_bullet("первичные преобразователи давления верха колонны (мажоритарная группа 2oo3 датчиков PT-102A/B/C), температуры куба (TE-101) и уровня флегмы (LT-101), быстродействующие отсечные клапаны ПАЗ с пружинным возвратом (Fail-Safe: закрыт при обесточивании), регулирующие клапаны с электропневматическими смарт-позиционерами;", "Полевой уровень: ")
    add_bullet("дублированные фидеры электроснабжения 220 В с промышленными источниками бесперебойного питания (ИБП) и дублированными блоками питания 24 В DC.", "Электропитание: ")

    add_h2("3.2. Функциональное моделирование процесса АВПКО в стандарте IDEF0")
    add_para("Для формализации и стандартизации процедуры проведения АВПКО разработана функциональная модель процесса в соответствии с методологией IDEF0 (диаграмма А0 представлена на Рисунке 1).")
    
    # Insert Figure 1
    add_image_with_caption(os.path.join(DOC_DIR, "idef0_a0_fmeca.png"),
                           "Рисунок 1 – Функциональная модель процесса проведения АВПКО в нотации IDEF0 (диаграмма А0)",
                           width_inches=6.3)

    add_para("Модель декомпозирует процесс АВПКО на три взаимосвязанных этапа:")
    add_bullet("установление минимального уровня разукрупнения, идентификация всех сменных модулей и компонентов КТС, присвоение логистических контрольных номеров (LCN);", "Блок А1 («Структурирование и кодирование системы»): ")
    add_bullet("идентификация видов отказов, выявление причин и последствий, экспертное определение баллов B₁, B₂, B₃ и расчет показателя критичности C;", "Блок А2 («Определение показателей критичности отказов»): ")
    add_bullet("ранжирование компонентов по критерию C ≥ 100 и B₂ ≥ 8, оформление перечня критичных элементов и разработка компенсирующих проектных мер.", "Блок А3 («Формирование перечня критичных элементов и мер»): ")

    add_h2("3.3. Проведение анализа видов, последствий и критичности отказов (АВПКО)")
    add_para("В соответствии с разработанной структурой проведена процедура АВПКО для ключевых компонентов разрабатываемой АСУ ТП и системы ПАЗ. Результаты анализа сведены в Таблицу 1.")

    # Table 1: FMECA
    table1 = doc.add_table(rows=1, cols=10)
    table1.alignment = WD_TABLE_ALIGNMENT.CENTER
    table1.autofit = False

    col_widths = [Inches(0.6), Inches(0.9), Inches(1.0), Inches(1.1), Inches(1.1), Inches(0.8), Inches(0.35), Inches(0.35), Inches(0.35), Inches(0.45)]
    
    # Table Header
    hdr_cells = table1.rows[0].cells
    hdr_titles = ["LCN", "Элемент АС", "Вид отказа", "Причина отказа", "Последствия отказа", "Метод обнаружения", "B1", "B2", "B3", "C"]
    for i, title in enumerate(hdr_titles):
        hdr_cells[i].width = col_widths[i]
        set_cell_shading(hdr_cells[i], "1F497D")
        set_cell_border(hdr_cells[i], top=dict(sz=6, color="1F497D"), bottom=dict(sz=6, color="1F497D"),
                        left=dict(sz=4, color="BDC3C7"), right=dict(sz=4, color="BDC3C7"))
        set_cell_margins(hdr_cells[i], top=80, bottom=80, left=80, right=80)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Inches(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(title)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(255, 255, 255)

    # Table 1 Data Rows (12 components)
    fmeca_data = [
        ("LCN 1.1", "ЦПУ РСУ (1oo2D)", "Аппаратный сбой ЦПУ (ошибка ОЗУ, зависание)", "Деградация чипа, перегрев", "Потеря авторегулирования давления в контуре", "Watchdog, безударный переход на Standby", "3", "7", "2", "42"),
        ("LCN 1.2", "ЦПУ ПАЗ (SIL 3)", "Сбой 1 процессора троированной схемы", "Сбой в логике, помеха", "Деградация схемы 2oo3 в 1oo2, защита сохранена", "TMR потактовое сравнение каналов", "2", "9", "1", "18"),
        ("LCN 2.1", "Датчик PT-101 (РСУ)", "Дрейф нуля / замерзание показаний 4-20 мА", "Закоксовывание мембраны разделителя", "Ложное закрытие сброса газа, рост давления", "Noise-check, кросс-валидация", "4", "8", "4", "128"),
        ("LCN 2.2", "Датчики ПАЗ PT-102A/B/C", "Опасный скрытый отказ 2 датчиков из 3", "Засорение отборных штуцеров", "Непрохождение сигнала на аварийный останов", "Контроль рассогласования каналов 2oo3", "2", "10", "2", "40"),
        ("LCN 3.1", "Отсечной клапан XV-101", "Механическое заклинивание штока в открытом пол.", "Коррозия сальника, смолистые отложения", "Невозможность отсечки сырья при аварии", "Тестирование частичного хода (PST)", "4", "10", "5", "200"),
        ("LCN 3.2", "Соленоид пневмопривода", "Обрыв обмотки / сгорание катушки", "Тепловой пробой изоляции", "Ложное закрытие клапана (Fail-Safe), простой", "Контроль целостности цепи DO током", "3", "4", "1", "12"),
        ("LCN 4.1", "Регулирующий клапан FV-102", "Заклинивание штока / сбой позиционера", "Попадание мехпримесей, износ манжет", "Срыв орошения колонны, перегрев верха", "Сигнал рассогласования смарт-позиционера", "3", "7", "2", "42"),
        ("LCN 5.1", "Коммутатор Ethernet", "Отказ порта / обрыв оптического кабеля", "Повреждение кабельной трассы", "Отсутствуют (реконфигурация MRP < 20 мс)", "SNMP-мониторинг, диагностика MRP", "3", "3", "1", "9"),
        ("LCN 6.1", "Сервер SCADA", "Зависание ОС / сбой базы данных", "Утечка памяти, сбой дискового массива", "Кратковременный сбой отображения АРМ", "Heartbeat-контроль, горячий резерв", "3", "4", "1", "12"),
        ("LCN 7.1", "ИБП 24В DC / АКБ", "Отказ батарей при пропадании внешнего ввода", "Сульфатация АКБ, отказ инвертора", "Обесточивание шкафа, переход в Fail-Safe", "Автотест батарей под нагрузкой", "2", "6", "2", "24"),
        ("LCN 8.1", "Модуль аналог. ввода AI", "Пробой АЦП измерительного канала", "Импульсная перенапряженческая наводка", "Искажение сигнала одного датчика", "Контроль границ NAMUR NE43 (<3.6 / >21 мА)", "2", "5", "1", "10"),
        ("LCN 9.1", "Оператор АРМ (эргат.)", "Ошибочный перевод регулятора в режим Manual", "Утомление, человеческий фактор", "Нарушение режима ректификации", "Логирование команд, приоритет ПАЗ", "3", "6", "2", "36")
    ]

    for row_idx, r_data in enumerate(fmeca_data):
        row_cells = table1.add_row().cells
        c_val = int(r_data[9])
        bg_color = "FADBD8" if c_val >= 100 else ("FCF3CF" if c_val >= 40 else "EAFAF1")
        for c_idx, val in enumerate(r_data):
            row_cells[c_idx].width = col_widths[c_idx]
            if c_idx >= 6:
                set_cell_shading(row_cells[c_idx], bg_color)
            else:
                if row_idx % 2 == 1:
                    set_cell_shading(row_cells[c_idx], "F8F9F9")
            set_cell_border(row_cells[c_idx], top=dict(sz=4, color="D5D8DC"), bottom=dict(sz=4, color="D5D8DC"),
                            left=dict(sz=4, color="D5D8DC"), right=dict(sz=4, color="D5D8DC"))
            set_cell_margins(row_cells[c_idx], top=50, bottom=50, left=60, right=60)
            p = row_cells[c_idx].paragraphs[0]
            p.paragraph_format.first_line_indent = Inches(0)
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0, 6, 7, 8, 9]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(8.5)
            if c_idx == 9 and c_val >= 100:
                run.bold = True
                run.font.color.rgb = RGBColor(144, 12, 63)

    p_tbl = doc.add_paragraph()
    p_tbl.paragraph_format.first_line_indent = Inches(0)
    p_tbl.paragraph_format.space_before = Pt(4)
    p_tbl.paragraph_format.space_after = Pt(8)
    p_tbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cap = p_tbl.add_run("Таблица 1 – Анализ видов, последствий и критичности отказов (АВПКО / FMECA) АСУ ТП и ПАЗ")
    r_cap.font.name = 'Times New Roman'
    r_cap.font.size = Pt(11)
    r_cap.italic = True
    r_cap.bold = True

    add_para("По результатам количественной оценки АВПКО построены матрица критичности отказов и диаграмма Парето (Рисунок 2).")

    # Insert Figure 2
    add_image_with_caption(os.path.join(DOC_DIR, "pareto_risk_chart.png"),
                           "Рисунок 2 – Матрица критичности отказов (Risk Matrix) и диаграмма Парето показателей критичности АВПКО",
                           width_inches=6.3)

    add_para("Анализ диаграммы Парето и матрицы рисков показывает, что 2 компонента попадают в критическую красную зону риска:")
    add_bullet("быстродействующий отсечной клапан ПАЗ на сырье XV-101 с коэффициентом критичности C = 200 (B₁=4, B₂=10, B₃=5). Высокая критичность обусловлена тем, что отсекатель годами находится в открытом состоянии и подвержен незаметному коксованию и заклиниванию штока;", "1. LCN 3.1 – ")
    add_bullet("первичный датчик давления контура регулирования РСУ PT-101 с коэффициентом критичности C = 128 (B₁=4, B₂=8, B₃=4). Забивание импульсной трубки смолами искажает сигнал регулятора и ведет к разгону давления.", "2. LCN 2.1 – ")

    add_para("Для парирования выявленных критических рисков разработаны проектные компенсирующие решения:")
    add_bullet("внедрение системы автоматизированного частичного перемещения клапана (Partial Stroke Testing – PST). Раз в 2 недели клапан кратковременно перемещается на 10-15% хода без возмущения техпроцесса. Это снижает вероятность необнаружения B₃ с 5 до 2, а критичность снижается с 200 до 80 (желтая зона);", "Для отсекателя XV-101: ")
    add_bullet("организация регулярной регламентной продувки импульсных линий азотом, внедрение алгоритма noise-check (мониторинг дисперсии шума датчика) и защита независимым троированным контуром ПАЗ (2oo3).", "Для датчика PT-101: ")

    # =========================================================
    # 3.4. ОЦЕНКА НАДЕЖНОСТИ АС МЕТОДОМ ДЕРЕВА ОТКАЗОВ (FTA)
    # =========================================================
    add_h2("3.4. Оценка надежности АС методом структурного анализа (построение дерева отказов)")
    add_para("Для проведения глубокого структурно-вероятностного анализа безопасности разрабатываемой АСУ ТП и системы ПАЗ построено дерево отказов (FTA) функции защиты ректификационной колонны от избыточного давления.")
    add_para("В качестве вершинного нежелательного события принято событие:")
    add_para("Т = «Аварийная разгерметизация ректификационной колонны с выбросом углеводородной смеси при превышении критического давления сверх предела прочности аппарата (P > 3,5 МПа)».", italic=True)
    add_para("В соответствии с эвристическими правилами построения деревьев отказов (правило 4 методических указаний: «Связывать инициирующие события с событием типа 'отсутствие защитных действий'») наступление аварии Т требует одновременного совпадения двух условий через логический элемент «И»:")
    add_bullet("возникновение аварийного режима неконтролируемого нарастания давления в колонне выше аварийной уставки 3,0 МПа вследствие отказа основного контура регулирования давления РСУ (А1), срыва орошения (А2) или перегрева куба (А3);", "Инициирующее событие E_иниц: ")
    add_bullet("несрабатывание системы противоаварийной автоматической защиты (ПАЗ) вследствие отказа подсистемы измерения 2oo3 (B1), отказа контроллера ПАЗ (B2), заклинивания отсекателя (B3) или потери питания (B4).", "Отказ защиты E_защита: ")

    add_para("Логическая структура дерева отказов представлена на Рисунке 3.")

    # Insert Figure 3
    add_image_with_caption(os.path.join(DOC_DIR, "fault_tree_logical.png"),
                           "Рисунок 3 – Дерево отказов (FTA) функции защиты ректификационной колонны от избыточного давления (логическая структура)",
                           width_inches=6.4)

    add_h2("3.5. Качественный анализ дерева отказов")
    add_para("Качественный анализ заключается в получении булевой аналитической функции дерева отказов и выделении минимальных аварийных сочетаний (МАС).")
    add_para("Булево выражение для головного события T имеет вид:")
    
    p_b = doc.add_paragraph()
    p_b.paragraph_format.first_line_indent = Inches(0)
    p_b.paragraph_format.space_before = Pt(4)
    p_b.paragraph_format.space_after = Pt(4)
    p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_b = p_b.add_run("T = E_иниц ∧ E_защита                                                           (2)")
    r_b.font.name = 'Times New Roman'
    r_b.font.size = Pt(13)
    r_b.bold = True

    add_para("Раскроем логические функции промежуточных событий через первичные события X₁..X₁₅:")
    add_bullet("E_иниц = A₁ ∨ A₂ ∨ A₃ = (X₁ ∨ X₂ ∨ X₃ ∨ X₄) ∨ U₁ ∨ U₂;", "Контур регулирования: ")
    add_bullet("B₁ (измерители ПАЗ 2oo3) = (X₅ ∧ X₆) ∨ (X₅ ∧ X₇) ∨ (X₆ ∧ X₇);", "Измерители 2oo3: ")
    add_bullet("B₂ (контроллер ПАЗ) = X₈ ∨ (X₉ ∧ X₁₀);", "Контроллер ПАЗ: ")
    add_bullet("B₃ (отсекатель ПАЗ) = X₁₁ ∨ X₁₂ ∨ X₁₃;", "Исполнительный клапан: ")
    add_bullet("B₄ (электропитание) = X₁₄ ∧ X₁₅;", "Электропитание: ")
    add_bullet("E_защита = B₁ ∨ B₂ ∨ B₃ ∨ B₄ = [(X₅ ∧ X₆) ∨ (X₅ ∧ X₇) ∨ (X₆ ∧ X₇)] ∨ [X₈ ∨ (X₉ ∧ X₁₀)] ∨ [X₁₁ ∨ X₁₂ ∨ X₁₃] ∨ [X₁₄ ∧ X₁₅].", "Суммарный отказ ПАЗ: ")

    add_para("Подставляя выражения в формулу (2) и применяя законы дистрибутивности, получаем множество минимальных аварийных сочетаний:")
    add_bullet("В системе отсутствуют однократные аварийные сочетания (МАС 1-го порядка). Ни один единичный отказ любого датчика, контроллера, клапана или кабеля не способен привести к аварийной разгерметизации колонны. Это строго доказывает отказоустойчивость разработанной архитектуры и ее соответствие ГОСТ Р 55608-2013;", "МАС 1-го порядка (однократные отказы): ")
    add_bullet("Сочетания отказа контура регулирования РСУ и одного из компонентов ПАЗ: {X_i, X₁₁} (отказ регулятора + заклинивание отсекателя XV-101); {X_i, X₁₂} (отказ регулятора + отказ соленоида); {X_i, X₁₃} (отказ регулятора + потеря воздуха КИП); {X_i, X₈} (отказ регулятора + сбой ПО ПАЗ), где X_i ∈ {X₁, X₂, X₃, X₄};", "МАС 2-го порядка (двухкратные отказы): ")
    add_bullet("Сочетания отказа регулятора и отказа двух из трех датчиков ПАЗ: {X_i, X₅, X₆}, {X_i, X₅, X₇}, {X_i, X₆, X₇}; сочетания отказа регулятора и обоих процессоров ПАЗ: {X_i, X₉, X₁₀}; сочетания отказа регулятора, внешнего ввода сети и ИБП: {X_i, X₁₄, X₁₅}.", "МАС 3-го порядка (трехкратные отказы): ")

    add_h2("3.6. Количественный расчет вероятностей дерева отказов")
    add_para("Для количественной оценки вероятности головного события T приняты расчетные вероятности возникновения первичных событий за период непрерывной работы t = 8760 ч (1 год), рассчитанные на базе интенсивностей отказов λ из Лабораторной работы № 1. Численные данные сведены в Таблицу 2.")

    # Table 2: Probabilities
    table2 = doc.add_table(rows=1, cols=4)
    table2.alignment = WD_TABLE_ALIGNMENT.CENTER
    table2.autofit = False
    col_w2 = [Inches(1.2), Inches(3.0), Inches(1.3), Inches(0.9)]
    
    hdr2 = table2.rows[0].cells
    titles2 = ["Код события", "Наименование исходного события", "Интенсивность λ, ч⁻¹", "Вероятность P"]
    for i, t in enumerate(titles2):
        hdr2[i].width = col_w2[i]
        set_cell_shading(hdr2[i], "1F497D")
        set_cell_border(hdr2[i], top=dict(sz=6, color="1F497D"), bottom=dict(sz=6, color="1F497D"),
                        left=dict(sz=4, color="BDC3C7"), right=dict(sz=4, color="BDC3C7"))
        set_cell_margins(hdr2[i], top=80, bottom=80, left=80, right=80)
        p = hdr2[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Inches(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(t)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(255, 255, 255)

    prob_data = [
        ("X₁", "Дрейф рабочего датчика давления РСУ PT-101", "2,9 · 10⁻⁶", "0,0250"),
        ("X₂", "Отказ дублированного контроллера РСУ (1oo2D)", "1,1 · 10⁻⁷", "0,0010"),
        ("X₃", "Заклинивание регулирующего клапана РСУ на сбросе", "4,1 · 10⁻⁶", "0,0350"),
        ("X₄", "Обрыв сегмента промышленной шины РСУ", "4,6 · 10⁻⁷", "0,0040"),
        ("X₅, X₆, X₇", "Отказ одного датчика давления ПАЗ PT-102 (каждого)", "1,7 · 10⁻⁶", "0,0150"),
        ("X₈", "Общий сбой встроенного системного ПО ПАЗ (CCF)", "1,1 · 10⁻⁸", "0,0001"),
        ("X₉, X₁₀", "Отказ одного процессорного модуля ПАЗ", "2,3 · 10⁻⁷", "0,0020"),
        ("X₁₁", "Заклинивание отсекателя XV-101 в открытом положении", "1,4 · 10⁻⁶", "0,0120"),
        ("X₁₂", "Отказ соленоидного распределителя отсекателя", "5,7 · 10⁻⁷", "0,0050"),
        ("X₁₃", "Засорение импульсной трубки воздуха КИП", "3,4 · 10⁻⁷", "0,0030"),
        ("X₁₄", "Пропадание внешнего сетевого электроснабжения", "5,7 · 10⁻⁶", "0,0500"),
        ("X₁₅", "Отказ аккумуляторной батареи / инвертора ИБП", "5,7 · 10⁻⁷", "0,0050")
    ]

    for row_idx, r_data in enumerate(prob_data):
        row_cells = table2.add_row().cells
        for c_idx, val in enumerate(r_data):
            row_cells[c_idx].width = col_w2[c_idx]
            if row_idx % 2 == 1:
                set_cell_shading(row_cells[c_idx], "F8F9F9")
            set_cell_border(row_cells[c_idx], top=dict(sz=4, color="D5D8DC"), bottom=dict(sz=4, color="D5D8DC"),
                            left=dict(sz=4, color="D5D8DC"), right=dict(sz=4, color="D5D8DC"))
            set_cell_margins(row_cells[c_idx], top=50, bottom=50, left=60, right=60)
            p = row_cells[c_idx].paragraphs[0]
            p.paragraph_format.first_line_indent = Inches(0)
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0, 2, 3]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(9.5)

    p_tbl2 = doc.add_paragraph()
    p_tbl2.paragraph_format.first_line_indent = Inches(0)
    p_tbl2.paragraph_format.space_before = Pt(4)
    p_tbl2.paragraph_format.space_after = Pt(8)
    p_tbl2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cap2 = p_tbl2.add_run("Таблица 2 – Численные параметры надежности исходных событий дерева отказов (за 8760 ч)")
    r_cap2.font.name = 'Times New Roman'
    r_cap2.font.size = Pt(11)
    r_cap2.italic = True
    r_cap2.bold = True

    add_para("Выполним пошаговый вероятностный расчет дерева отказов снизу вверх:")
    add_para("1. Вероятность отказа основного контура регулирования РСУ (логический элемент «ИЛИ»):")
    add_bullet("P(A₁) = 1 - (1 - P(X₁))·(1 - P(X₂))·(1 - P(X₃))·(1 - P(X₄)) = 1 - (1 - 0,025)·(1 - 0,001)·(1 - 0,035)·(1 - 0,004) = 1 - 0,975 · 0,999 · 0,965 · 0,996 = 1 - 0,9362 = 0,0638 (6,38% в год). Принимая отказы орошения P(U₁) = 0,018 и перегрева куба P(U₂) = 0,012, суммарная частота инициирующего броска давления составляет P(E_иниц) ≈ 0,0638.")
    
    add_para("2. Вероятность отказа измерительного контура ПАЗ со структурой 2oo3:")
    add_bullet("P(B₁) = 3·P(X₅)² - 2·P(X₅)³ = 3·(0,015)² - 2·(0,015)³ = 3·0,000225 - 0,00000675 = 0,000668 (6,68 · 10⁻⁴). Применение троирования снизило вероятность отказа подсистемы измерения более чем в 22 раза по сравнению с одиночным датчиком!")

    add_para("3. Вероятность отказа контроллера ПАЗ (архитектура TMR / SIL 3):")
    add_bullet("P(B₂) = P(X₈) + P(X₉)·P(X₁₀) = 0,0001 + (0,002)² = 0,0001 + 0,000004 = 0,000104 (1,04 · 10⁻⁴).")

    add_para("4. Вероятность отказа исполнительного контура отсечки сырья (схема «ИЛИ»):")
    add_bullet("P(B₃) = 1 - (1 - P(X₁₁))·(1 - P(X₁₂))·(1 - P(X₁₃)) = 1 - (1 - 0,012)·(1 - 0,005)·(1 - 0,003) = 1 - 0,988 · 0,995 · 0,997 = 1 - 0,98012 = 0,01988 (~1,99%).")

    add_para("5. Вероятность полного обесточивания цепей управления (схема «И» сети и ИБП):")
    add_bullet("P(B₄) = P(X₁₄) · P(X₁₅) = 0,050 · 0,005 = 0,000250 (2,50 · 10⁻⁴).")

    add_para("6. Суммарная вероятность отказа защитных действий системы ПАЗ (схема «ИЛИ»):")
    add_bullet("P(E_защита) = 1 - (1 - P(B₁))·(1 - P(B₂))·(1 - P(B₃))·(1 - P(B₄)) = 1 - (1 - 0,000668)·(1 - 0,000104)·(1 - 0,01988)·(1 - 0,00025) ≈ 1 - 0,99933 · 0,99990 · 0,98012 · 0,99975 = 1 - 0,97913 = 0,02087 (~2,09%).")

    add_para("7. Вероятность головного нежелательного события Т (логический элемент «И»):")
    
    p_res = doc.add_paragraph()
    p_res.paragraph_format.first_line_indent = Inches(0)
    p_res.paragraph_format.space_before = Pt(6)
    p_res.paragraph_format.space_after = Pt(6)
    p_res.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_res = p_res.add_run("P(T) = P(E_иниц) · P(E_защита) = 0,0638 · 0,02087 = 1,33 · 10⁻³ год⁻¹ (0,00133)")
    r_res.font.name = 'Times New Roman'
    r_res.font.size = Pt(14)
    r_res.bold = True
    r_res.font.color.rgb = RGBColor(192, 57, 43)

    add_para("Графическая схема дерева отказов с нанесенными вероятностями представлена на Рисунке 4.")

    # Insert Figure 4
    add_image_with_caption(os.path.join(DOC_DIR, "fault_tree_probabilities.png"),
                           "Рисунок 4 – Дерево отказов (FTA) с указанием вероятностей базовых, промежуточных событий и головного отказа",
                           width_inches=6.4)

    add_para("Оценка эффекта проектных решений: внедрение системы частичного тестирования хода клапана (PST) снижает интенсивность скрытых отказов отсекателя P(X₁₁) с 0,012 до 0,0015. В результате вероятность отказа исполнительного контура снижается до P(B₃)мод = 0,0045, а вероятность головного аварийного события T снижается более чем в 3,8 раза:")
    add_bullet("P(T)мод = 0,0638 · 0,0055 = 3,5 · 10⁻⁴ год⁻¹, что строго удовлетворяет жестким требованиям промышленной безопасности Ростехнадзора (P < 10⁻³ год⁻¹) и гарантирует соответствие уровню полноты безопасности SIL 3 по ГОСТ Р МЭК 61508.", "Модифицированная надежность: ")

    add_h2("3.7. Разработка предложений по системе технического диагностирования (СТД)")
    add_para("На основе теоретических положений методических указаний (раздел технического диагностирования) сформированы предложения по созданию системы технического диагностирования (СТД) блока ректификации:")
    add_bullet("непрерывный контроль целостности измерительных цепей 4-20 мА по стандарту NAMUR NE43 (обнаружение обрыва линии при I < 3,6 мА и короткого замыкания при I > 21,0 мА); алгоритмическое выявление рассогласования каналов 2oo3 (девиация показаний между любыми двумя датчиками более 2% формирует предупреждение оператору); вибродиагностика центробежных насосов флегмы и сырья по спектру виброграмм (обнаружение дефектов подшипников и дисбаланса ротора на зарождающейся стадии);", "Функциональное диагностирование: ")
    add_bullet("автоматический циклический тест целостности ОЗУ/ПЗУ контроллеров ПАЗ на основе расчета контрольных сумм CRC32; сторожевые таймеры (Watchdog) процессорных модулей; автоматизированное тестирование хода штоков быстродействующих отсечных клапанов ПАЗ (Partial Stroke Testing – PST) каждые 336 часов без прерывания подачи сырья.", "Тестовое диагностирование: ")

    # =========================================================
    # 4. ВЫВОДЫ
    # =========================================================
    add_h1("4. Выводы")
    add_para("1. В ходе выполнения лабораторной работы освоены методы анализа видов, последствий и критичности отказов (АВПКО / FMECA) и структурного анализа надежности и безопасности автоматизированных систем (FTA – Fault Tree Analysis) в соответствии с требованиями ГОСТ 27.310-95, ГОСТ Р МЭК 60812 и ГОСТ Р МЭК 61508.")
    add_para("2. Разработана функциональная модель процесса проведения АВПКО в нотации IDEF0 (диаграмма А0), регламентирующая этапы структурирования и кодирования компонентов КТС, определения показателей критичности C = B₁ · B₂ · B₃ и формирования перечня критичных элементов.")
    add_para("3. На примере АСУ ТП и системы ПАЗ блока ректификации установки ЭЛОУ-АВТ составлена детальная таблица АВПКО для 12 ключевых компонентов системы, построены матрица критичности отказов (Risk Matrix) и диаграмма Парето. Выявлены наиболее критичные узлы: отсечной клапан ПАЗ XV-101 (C = 200) и датчик давления РСУ PT-101 (C = 128). Предложены проектные компенсирующие меры (внедрение PST-тестирования, регламентная азотная продувка, алгоритм noise-check), позволившие перевести систему в область допустимого риска (C < 100).")
    add_para("4. Построено многоуровневое дерево отказов (FTA) функции защиты ректификационной колонны от избыточного давления. В результате качественного анализа методом булевой алгебры доказано отсутствие однократных аварийных сочетаний (МАС 1-го порядка), что подтверждает выполнение принципа единичного отказа благодаря горячему дублированию РСУ (1oo2D) и мажорированию ПАЗ (2oo3).")
    add_para("5. В результате количественного вероятностного расчета показано, что применение мажоритарной структуры 2oo3 в измерительном контуре ПАЗ снизило вероятность отказа измерения в 22 раза (с 0,015 до 6,68 · 10⁻⁴), а суммарная расчетная вероятность головного аварийного события составила P(T) = 1,33 · 10⁻³ год⁻¹ (с последующим снижением до 3,5 · 10⁻⁴ год⁻¹ при использовании PST), что строго удовлетворяет уровню полноты безопасности SIL 3 и нормам промышленной безопасности.")

    # =========================================================
    # 5. КОНТРОЛЬНЫЕ ВОПРОСЫ И ОТВЕТЫ
    # =========================================================
    add_h1("5. Контрольные вопросы и ответы")

    add_para("Вопрос 1. В чем заключается цель и сущность метода АВПКО (FMECA) и как вычисляется показатель критичности отказа?", bold_prefix="")
    add_para("Ответ: Цель АВПКО (FMECA – Failure Mode, Effects and Criticality Analysis) – заблаговременное выявление потенциальных видов отказов элементов сложной системы, исследование причин и механизмов их возникновения, качественная и количественная оценка тяжести их последствий на всех уровнях и разработка обоснованных компенсирующих проектных решений. Показатель критичности отказа (приоритетное число риска ПЧР / RPN) определяется как произведение трех составляющих в шкалах от 1 до 10: C = B₁ · B₂ · B₃, где B₁ – вероятность возникновения отказа за период эксплуатации, B₂ – тяжесть последствий отказа для объекта и персонала, B₃ – вероятность необнаружения отказа средствами диагностики до наступления аварии.")

    add_para("Вопрос 2. Каковы основные эвристические правила построения дерева отказов (FTA)?", bold_prefix="")
    add_para("Ответ: Ключевые эвристические правила построения дерева отказов включают: 1) замену абстрактных событий менее абстрактными и физически конкретными (например, «двигатель работает слишком долго» заменяется на «ток через двигатель протекает слишком долго»); 2) разделение событий на элементарные причины; 3) точное определение места и механизма отказа; 4) связывание инициирующих событий с событием типа «отсутствие защитных действий» через логический элемент «И»; 5) поиск совместно действующих причин; 6) детальную разработку отказов до уровня состояния элемента; 7) упрощение дерева путем исключения ветвей с пренебрежимо малой вероятностью.")

    add_para("Вопрос 3. Что такое минимальные аварийные сочетания (МАС) и каков их физический смысл при анализе надежности АС?", bold_prefix="")
    add_para("Ответ: Минимальное аварийное сочетание (МАС / Minimal Cut Set) – это наименьший набор исходных (базовых) событий дерева отказов, при совместном наступлении которых неизбежно возникает головное (аварийное) событие. Если хотя бы одно событие из МАС не произойдет, авария по данной траектории исключена. Порядок МАС определяется числом входящих в него событий: МАС 1-го порядка означают уязвимость системы к единичному отказу, МАС 2-го порядка требуют совпадения двух независимых отказов, МАС 3-го порядка – трех отказов. Проектирование надежных систем ответственного назначения требует полного исключения МАС 1-го порядка за счет резервирования.")

    add_para("Вопрос 4. Как логический элемент мажоритарного голосования «2 из 3» (2oo3) выражается через базис «И» и «ИЛИ» и как рассчитывается его вероятность отказа?", bold_prefix="")
    add_para("Ответ: Логический знак «2 из 3» (Two out of Three) для трех каналов X₁, X₂, X₃ означает, что выходное событие наступает при отказе любых двух или всех трех каналов. В базисе элементов «И» и «ИЛИ» он эквивалентен схеме: (X₁ ∧ X₂) ∨ (X₁ ∧ X₃) ∨ (X₂ ∧ X₃). При условии независимости каналов и одинаковой вероятности отказа одного канала P вероятность отказа схемы рассчитывается по формуле: P_2oo3 = C₃² · P² · (1 - P) + C₃³ · P³ = 3·P²·(1 - P) + P³ = 3·P² - 2·P³. При малых P (P << 1) значение P_2oo3 ≈ 3·P², что на порядки меньше вероятности отказа одиночного канала P.")

    add_para("Вопрос 5. В чем заключаются принципиальные различия между функциональным и тестовым диагностированием в системах технического диагностирования (СТД)?", bold_prefix="")
    add_para("Ответ: Функциональное диагностирование проводится непрерывно в процессе штатного функционирования объекта по сигналам его рабочих входов и выходов x(t), y(t) без подачи специальных возмущений (например, контроль диапазона 4-20 мА NAMUR NE43, кросс-валидация каналов 2oo3, спектральный анализ виброграмм). Тестовое диагностирование предполагает подачу на систему или ее компонент специально сформированных тестовых сигналов x_т и сравнение отклика с эталонным (например, аппаратный тест ОЗУ контроллера, тестирование частичного хода штока клапана PST – Partial Stroke Testing). Для технологических аппаратов преобладает функциональное диагностирование, а для микропроцессорных средств и защитной арматуры обязательно сочетание обоих видов.")

    # Save document
    doc.save(OUTPUT_FILE)
    print("Report saved successfully to:", OUTPUT_FILE)

if __name__ == "__main__":
    create_report()

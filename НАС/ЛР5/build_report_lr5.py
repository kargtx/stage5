import os
import docx
from docx.shared import Inches, Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

DIR_LR5 = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\НАС\ЛР5"
REPORT_PATH = os.path.join(DIR_LR5, "ЭАС-414С_НАС_ГайнановГарифуллин_ЛР5.docx")
FIG_PATH = os.path.join(DIR_LR5, "fig_stage_gate_system.png")

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_shading(cell, color_hex):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def set_table_borders(table, color="000000", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders_xml = f'''
    <w:tblBorders {nsdecls("w")}>
        <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
    </w:tblBorders>
    '''
    tblPr.append(parse_xml(borders_xml))

def format_paragraph(p, space_before=0, space_after=0, line_spacing=1.15, first_indent=1.25, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if first_indent > 0:
        p.paragraph_format.first_line_indent = Mm(first_indent * 10)
    else:
        p.paragraph_format.first_line_indent = Mm(0)
    p.alignment = align

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, space_before=12, space_after=6, line_spacing=1.15, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    format_paragraph(p, space_before=8, space_after=4, line_spacing=1.15, first_indent=1.25, align=WD_ALIGN_PARAGRAPH.LEFT)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_body_p(doc, text, bold_prefix="", italic_prefix=""):
    p = doc.add_paragraph()
    format_paragraph(p, space_before=0, space_after=0, line_spacing=1.15, first_indent=1.25, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    if bold_prefix:
        r_b = p.add_run(bold_prefix)
        r_b.font.name = 'Times New Roman'
        r_b.font.size = Pt(14)
        r_b.font.bold = True
        r_b.font.color.rgb = RGBColor(0, 0, 0)
    if italic_prefix:
        r_i = p.add_run(italic_prefix)
        r_i.font.name = 'Times New Roman'
        r_i.font.size = Pt(14)
        r_i.font.italic = True
        r_i.font.color.rgb = RGBColor(0, 0, 0)
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def build_report_lr5():
    doc = docx.Document()

    # Поля страницы ГОСТ: левое 30 мм, правое 15 мм, верхнее 20 мм, нижнее 20 мм
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Mm(30)
    section.right_margin = Mm(15)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)

    # ------------------ ТИТУЛЬНЫЙ ЛИСТ ------------------
    p = doc.add_paragraph()
    format_paragraph(p, space_before=0, space_after=2, line_spacing=1.0, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p.add_run("МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ\n"
                  "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ\n"
                  "ВЫСШЕГО ОБРАЗОВАНИЯ\n"
                  "«УФИМСКИЙ УНИВЕРСИТЕТ НАУКИ И ТЕХНОЛОГИЙ»\n"
                  "Кафедра автоматизированных систем управления")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_sp = doc.add_paragraph()
    format_paragraph(p_sp, space_before=70, space_after=0, line_spacing=1.0, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)

    p_rep = doc.add_paragraph()
    format_paragraph(p_rep, space_before=0, space_after=4, line_spacing=1.15, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_rep.add_run("ОТЧЕТ\nПО ЛАБОРАТОРНОЙ РАБОТЕ № 5\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    r_sub = p_rep.add_run("по дисциплине: «Надежность автоматизированных систем»\n"
                          "Тема: «Разработка показателей качества и надежности АСОИУ для гейтовой системы управления качеством на стадиях жизненного цикла»")
    r_sub.font.name = 'Times New Roman'
    r_sub.font.size = Pt(14)
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor(0, 0, 0)

    p_sp2 = doc.add_paragraph()
    format_paragraph(p_sp2, space_before=80, space_after=0, line_spacing=1.0, first_indent=0, align=WD_ALIGN_PARAGRAPH.RIGHT)

    p_sign = doc.add_paragraph()
    format_paragraph(p_sign, space_before=0, space_after=0, line_spacing=1.15, first_indent=0, align=WD_ALIGN_PARAGRAPH.RIGHT)
    r = p_sign.add_run("Выполнили: студенты гр. ЭАС-414С\n"
                       "Гайнанов Д.И., Гарифуллин К.Р.\n\n"
                       "Проверил: доцент кафедры АСУ,\n"
                       "канд. техн. наук Ризванов К.А.")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_sp3 = doc.add_paragraph()
    format_paragraph(p_sp3, space_before=90, space_after=0, line_spacing=1.0, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)

    p_city = doc.add_paragraph()
    format_paragraph(p_city, space_before=0, space_after=0, line_spacing=1.0, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_city.add_run("Уфа – 2026")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0, 0, 0)

    doc.add_page_break()

    # ------------------ ОСНОВНОЙ ТЕКСТ ------------------
    add_heading_1(doc, "1. ЦЕЛЬ РАБОТЫ")
    add_body_p(doc, "Целью настоящей работы является освоение методологии гейтового управления качеством и надежностью (Stage-Gate process) сложных систем на всех стадиях жизненного цикла, разработка сбалансированной номенклатуры единичных и интегральных показателей качества для промышленной АСОИУ (на примере распределенной системы управления и противоаварийной защиты ректификационной колонны разделения углеводородов), а также проведение сравнительного анализа альтернативных вариантов аппаратно-программного комплекса по критериям надежности и эффективности прохождения контрольных шлюзов.")

    add_heading_1(doc, "2. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ ГЕЙТОВОЙ СИСТЕМЫ УПРАВЛЕНИЯ КАЧЕСТВОМ АС")
    add_body_p(doc, "Гейтовая система управления качеством (модель Stage-Gate) представляет собой концептуальный и операционный механизм сквозного управления проектом создания сложной автоматизированной системы. Весь жизненный цикл разбивается на дискретные содержательные стадии (Stages), завершающиеся контрольными шлюзами («воротами качества» – Gates).")
    add_body_p(doc, "Каждый шлюз (Gate) функционирует как точка принятия стратегического решения о дальнейшем развитии проекта. В шлюзе межфункциональная экспертная комиссия оценивает достигнутый уровень единичных и интегральных показателей качества на основе объективных отчетных документов и принимает одно из четырех регламентированных решений:")
    add_body_p(doc, "Go (Проход) – результаты этапа соответствуют ТЗ и нормативным критериям шлюза, проект переходит на следующую стадию;", bold_prefix="   1. ")
    add_body_p(doc, "Kill (Прекращение) – выявлена критическая техническая или экономическая несостоятельность, разработка закрывается;", bold_prefix="   2. ")
    add_body_p(doc, "Hold (Приостановка) – разработка временно замораживается до появления необходимых ресурсов или сертификатов;", bold_prefix="   3. ")
    add_body_p(doc, "Recycle (Доработка) – показатели не достигли пороговых значений, проект возвращается на доработку с корректировкой КД/ТД.", bold_prefix="   4. ")
    add_body_p(doc, "Для проектируемой АСУ ТП и системы ПАЗ ректификационной колонны выделены четыре ключевые стадии жизненного цикла:")
    add_body_p(doc, "1) Стадия М – Маркетинг (Gate 1: «Утверждение концепции и требований ТЗ»);", bold_prefix="   • ")
    add_body_p(doc, "2) Стадия Р – Разработка и проектирование (Gate 2: «Верификация проектных решений и готовность к производству»);", bold_prefix="   • ")
    add_body_p(doc, "3) Стадия П – Производство и комплексирование (Gate 3: «Заводские приемочные испытания FAT и отгрузка»);", bold_prefix="   • ")
    add_body_p(doc, "4) Стадия Э – Эксплуатация (Gate 4: «Шлюз ввода в промышленную эксплуатацию SAT и мониторинг надежности»).", bold_prefix="   • ")
    add_body_p(doc, "На рисунке 1 представлена структурно-логическая схема гейтовой системы управления качеством и надежностью АСУ ТП.")

    # Вставка рисунка
    if os.path.exists(FIG_PATH):
        p_img = doc.add_paragraph()
        format_paragraph(p_img, space_before=6, space_after=2, line_spacing=1.0, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        p_img.add_run().add_picture(FIG_PATH, width=Inches(6.2))

        p_cap = doc.add_paragraph()
        format_paragraph(p_cap, space_before=2, space_after=8, line_spacing=1.0, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p_cap.add_run("Рисунок 1 – Схема гейтовой системы (Stage-Gate) управления качеством и надежностью АСУ ТП на стадиях жизненного цикла")
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.font.color.rgb = RGBColor(0, 0, 0)

    add_body_p(doc, "Для оценки качества на каждом шлюзе рассчитывается аддитивный комплексный (интегральный) показатель качества этапа I_stage по формуле взвешенной суммы нормированных единичных показателей:")
    add_body_p(doc, "I_stage = sum_{i=1}^m [ w_i * q~_i ] / sum_{i=1}^m w_i,", bold_prefix="   ", italic_prefix="")
    add_body_p(doc, "где w_i – весовой коэффициент важности i-го единичного показателя (sum w_i = 1,0); q~_i – безразмерный нормированный единичный показатель качества. Для показателей, стремление которых направлено на увеличение (наработка на отказ, готовность, срок службы, баллы эргономики), нормировка выполняется отношением: q~_i = q_i / q_баз. Для показателей, стремление которых направлено на уменьшение (время восстановления, быстpодействие/задержка, уровень брака, стоимость): q~_i = q_баз / q_i.")
    add_body_p(doc, "По аналогии с методическими указаниями для технико-экономического сопоставления рассчитывается удельный интегральный показатель качества по отношению к приведенным затратам C_затр:")
    add_body_p(doc, "I_уд = I_stage / C_затр.", bold_prefix="   ", italic_prefix="")

    add_heading_1(doc, "3. РАЗРАБОТКА ПОКАЗАТЕЛЕЙ КАЧЕСТВА И НАДЕЖНОСТИ ПО ЭТАПАМ ЖЦ")
    add_body_p(doc, "Проведем сравнительную оценку двух альтернативных вариантов реализации системы автоматизации технологического блока ректификации углеводородов:")
    add_body_p(doc, "Вариант 1 (Отечественная РСУ/ПАЗ): распределенная система на базе сертифицированных ПЛК REGUL R500 / TREI-5B-05 с дублированной архитектурой ЦПУ 1oo2D в ПАЗ, мажоритарными каналами давления 2oo3 и резервированными серверами SCADA;", bold_prefix="   • ")
    add_body_p(doc, "Вариант 2 (Импортная РСУ/ПАЗ): распределенный программно-технический комплекс на базе Yokogawa Centum VP / ProSafe-RS с тройным модульным резервированием (TMR 2oo3).", bold_prefix="   • ")

    # ------------------ ЭТАП 1: МАРКЕТИНГ ------------------
    add_heading_2(doc, "3.1. Этап 1: Маркетинг (Gate 1 «Концепция и ТЗ»)")
    add_body_p(doc, "На этапе маркетинга формулируются требования потребителя (нефтегазоперерабатывающего предприятия), анализируются условия взрывопожароопасности объекта, регламентируются ключевые технико-экономические и эксплуатационные показатели. В таблице 1 представлены единичные показатели качества для этапа Маркетинга.")

    # Таблица Маркетинг
    t1 = doc.add_table(rows=1, cols=6)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t1)

    t1_headers = ["№", "Единичный показатель качества", "Вес, w_i", "Базовое знач.", "Вариант 1 (REGUL)", "Вариант 2 (Yokogawa)"]
    t1_widths = [Mm(10), Mm(60), Mm(20), Mm(25), Mm(28), Mm(28)]
    for j, cell in enumerate(t1.rows[0].cells):
        cell.width = t1_widths[j]
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        set_cell_shading(cell, "EAEAEA")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        format_paragraph(p, space_before=2, space_after=2, line_spacing=1.0, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run(t1_headers[j])
        r.font.name = 'Times New Roman'
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)

    m_data = [
        ("1", "Назначенный срок службы, лет", "0,15", "15", "15", "20"),
        ("2", "Наработка на отказ КТС T0, тыс. ч", "0,20", "50", "65", "85"),
        ("3", "Гарантийный срок обслуживания, мес.", "0,10", "24", "36", "24"),
        ("4", "Время реакции ПАЗ на аварию, мс", "0,15", "100", "80", "50"),
        ("5", "Уровень полноты безопасности SIL", "0,15", "SIL 3", "SIL 3 (3)", "SIL 3 (3)"),
        ("6", "Эргономика HMI (по 10-балльной шкале)", "0,05", "8,0", "8,5", "9,5"),
        ("7", "Освояемость персоналом (10-балльная)", "0,05", "8,0", "9,0", "7,5"),
        ("8", "Наличие регионального сервиса 24/7", "0,10", "1,0", "1,0 (в РФ)", "0,7 (удален.)"),
        ("9", "Затраты на внедрение C_затр, млн руб.", "0,05", "50,0", "42,0", "78,0")
    ]

    for row_data in m_data:
        row = t1.add_row()
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            cell.width = t1_widths[j]
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            align = WD_ALIGN_PARAGRAPH.LEFT if j == 1 else WD_ALIGN_PARAGRAPH.CENTER
            format_paragraph(p, space_before=1, space_after=1, line_spacing=1.0, first_indent=0, align=align)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(0, 0, 0)

    add_body_p(doc, "Методы обеспечения выбранных показателей на этапе Маркетинга:")
    add_body_p(doc, "1) Социологический метод – анкетирование начальников цехов КИПиА и операторов НПЗ для выявления требований к HMI и сервисной поддержке;", bold_prefix="   • ")
    add_body_p(doc, "2) Расчетный метод – технико-экономический анализ стоимости владения (TCO) и предварительный расчет бюджета надежности;", bold_prefix="   • ")
    add_body_p(doc, "3) Экспертный метод – оценка соответствия проектных требований нормам Федеральных норм и правил в области промбезопасности ОПО;", bold_prefix="   • ")
    add_body_p(doc, "4) Статистический метод – сбор ретроспективных данных об аварийности и отказах систем автоматизации на действующих ГФУ.", bold_prefix="   • ")

    add_body_p(doc, "Расчет интегрального показателя качества этапа Маркетинга:")
    add_body_p(doc, "Для Варианта 1 (REGUL): q~1 = 15/15=1,0; q~2 = 65/50=1,30; q~3 = 36/24=1,50; q~4 = 100/80=1,25; q~5 = 3/3=1,0; q~6 = 8,5/8=1,06; q~7 = 9/8=1,125; q~8 = 1,0/1,0=1,0; q~9 = 50/42=1,19. Итого:")
    add_body_p(doc, "I_марк(Вар.1) = 0,15*1,0 + 0,20*1,30 + 0,10*1,50 + 0,15*1,25 + 0,15*1,0 + 0,05*1,06 + 0,05*1,125 + 0,10*1,0 + 0,05*1,19 = 1,215.", bold_prefix="   • ")
    add_body_p(doc, "Для Варианта 2 (Yokogawa): q~1 = 20/15=1,33; q~2 = 85/50=1,70; q~3 = 24/24=1,0; q~4 = 100/50=2,0; q~5 = 3/3=1,0; q~6 = 9,5/8=1,19; q~7 = 7,5/8=0,94; q~8 = 0,7/1,0=0,70; q~9 = 50/78=0,64. Итого:")
    add_body_p(doc, "I_марк(Вар.2) = 0,15*1,33 + 0,20*1,70 + 0,10*1,0 + 0,15*2,0 + 0,15*1,0 + 0,05*1,19 + 0,05*0,94 + 0,10*0,70 + 0,05*0,64 = 1,301.", bold_prefix="   • ")
    add_body_p(doc, "Удельный интегральный показатель с учетом затрат: I_уд(Вар.1) = 1,215 / 42 = 0,0289; I_уд(Вар.2) = 1,301 / 78 = 0,0167. По удельной эффективности отечественный комплекс опережает импортный в 1,73 раза при полном соответствии требованиям шлюза Gate 1 (I >= 0,85). Решение на шлюзе Gate 1: Go (Одобрен переход на стадию разработки).")

    # ------------------ ЭТАП 2: РАЗРАБОТКА ------------------
    add_heading_2(doc, "3.2. Этап 2: Разработка и проектирование (Gate 2 «Проектные решения»)")
    add_body_p(doc, "На этапе разработки выполняются схемотехнические расчеты, синтез алгоритмов ПАЗ, анализ видов и последствий отказов (FMECA) и математическое моделирование надежности. В таблице 2 представлены единичные показатели качества этапа разработки.")

    t2 = doc.add_table(rows=1, cols=6)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t2)

    for j, cell in enumerate(t2.rows[0].cells):
        cell.width = t1_widths[j]
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        set_cell_shading(cell, "EAEAEA")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        format_paragraph(p, space_before=2, space_after=2, line_spacing=1.0, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run(t1_headers[j])
        r.font.name = 'Times New Roman'
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)

    r_data = [
        ("1", "Вероятность безотказности за год P(8760)", "0,20", "0,990", "0,995", "0,998"),
        ("2", "Коэффициент готовности Kг", "0,20", "0,9995", "0,9998", "0,9999"),
        ("3", "Среднее время восстановления Tв, ч", "0,15", "0,50", "0,40", "0,45"),
        ("4", "Охват самодиагностикой C_diag, %", "0,15", "95,0", "99,2", "99,6"),
        ("5", "Запас загрузки процессора ПЛК, %", "0,10", "40,0", "55,0", "65,0"),
        ("6", "Время цикла опроса модулей УСО, мс", "0,10", "50,0", "20,0", "15,0"),
        ("7", "Запас памяти контроллера под расширение, %", "0,10", "30,0", "50,0", "50,0")
    ]

    for row_data in r_data:
        row = t2.add_row()
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            cell.width = t1_widths[j]
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            align = WD_ALIGN_PARAGRAPH.LEFT if j == 1 else WD_ALIGN_PARAGRAPH.CENTER
            format_paragraph(p, space_before=1, space_after=1, line_spacing=1.0, first_indent=0, align=align)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(0, 0, 0)

    add_body_p(doc, "Методы обеспечения выбранных показателей на этапе Разработки:")
    add_body_p(doc, "1) Расчетный аналитический метод – расчет показателей P(t), Kг, PFDavg марковскими цепями и графами состояний;", bold_prefix="   • ")
    add_body_p(doc, "2) Имитационное моделирование – отладка алгоритмов противоаварийной защиты и ПИД-регулирования давления в среде полунатурного стенда HIL (Hardware-in-the-Loop);", bold_prefix="   • ")
    add_body_p(doc, "3) Экспертный метод – верификация проектной документации на соответствие ГОСТ Р 27.310-95 (АВПКО) и нормам промышленной безопасности.", bold_prefix="   • ")

    add_body_p(doc, "Расчет интегрального показателя качества этапа Разработки:")
    add_body_p(doc, "Для Варианта 1 (REGUL): q~1 = 0,995/0,990=1,005; q~2 = 0,9998/0,9995=1,0003; q~3 = 0,50/0,40=1,25; q~4 = 99,2/95=1,044; q~5 = 55/40=1,375; q~6 = 50/20=2,50; q~7 = 50/30=1,667. Итого:")
    add_body_p(doc, "I_разр(Вар.1) = 0,20*1,005 + 0,20*1,0003 + 0,15*1,25 + 0,15*1,044 + 0,10*1,375 + 0,10*2,50 + 0,10*1,667 = 1,299.", bold_prefix="   • ")
    add_body_p(doc, "Для Варианта 2 (Yokogawa): q~1 = 0,998/0,990=1,008; q~2 = 0,9999/0,9995=1,0004; q~3 = 0,50/0,45=1,11; q~4 = 99,6/95=1,048; q~5 = 65/40=1,625; q~6 = 50/15=3,333; q~7 = 50/30=1,667. Итого:")
    add_body_p(doc, "I_разр(Вар.2) = 0,20*1,008 + 0,20*1,0004 + 0,15*1,11 + 0,15*1,048 + 0,10*1,625 + 0,10*3,333 + 0,10*1,667 = 1,448.", bold_prefix="   • ")
    add_body_p(doc, "Оба варианта с большим запасом превосходят критерий Gate 2 (I >= 0,85). Решение на шлюзе Gate 2: Go (Проект утвержден к производству шкафов управления).")

    # ------------------ ЭТАП 3: ПРОИЗВОДСТВО ------------------
    add_heading_2(doc, "3.3. Этап 3: Производство и комплексирование (Gate 3 «Заводской FAT»)")
    add_body_p(doc, "На стадии производства выполняется сборка шкафов КТС, 100% входной контроль ПКИ, климатические испытания, вибропрогон и заводские приемочные испытания FAT. В таблице 3 приведены показатели качества этапа производства.")

    t3 = doc.add_table(rows=1, cols=6)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t3)

    for j, cell in enumerate(t3.rows[0].cells):
        cell.width = t1_widths[j]
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        set_cell_shading(cell, "EAEAEA")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        format_paragraph(p, space_before=2, space_after=2, line_spacing=1.0, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run(t1_headers[j])
        r.font.name = 'Times New Roman'
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)

    p_data = [
        ("1", "Уровень брака монтажа и пайки, %", "0,20", "0,10", "0,04", "0,02"),
        ("2", "Длительность технологич. прогона, ч", "0,15", "48", "72", "96"),
        ("3", "Полнота охвата тестами FAT, %", "0,20", "98,0", "100,0", "100,0"),
        ("4", "Квалификация сборщиков (разряд)", "0,10", "4,5", "5,0", "5,2"),
        ("5", "Сопротивление изоляции цепей, МОм", "0,15", "50", "250", "300"),
        ("6", "Устойчивость к помехам ЭМС (степень)", "0,10", "3", "4", "4"),
        ("7", "Срок поставки и сборки шкафов, мес.", "0,10", "6,0", "3,5", "9,0")
    ]

    for row_data in p_data:
        row = t3.add_row()
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            cell.width = t1_widths[j]
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            align = WD_ALIGN_PARAGRAPH.LEFT if j == 1 else WD_ALIGN_PARAGRAPH.CENTER
            format_paragraph(p, space_before=1, space_after=1, line_spacing=1.0, first_indent=0, align=align)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(0, 0, 0)

    add_body_p(doc, "Методы обеспечения выбранных показателей на этапе Производства:")
    add_body_p(doc, "1) Измерительный метод – инструментальный контроль сопротивления изоляции мегаомметром на 500 В, контроль уровней пульсаций вторичного питания 24 В;", bold_prefix="   • ")
    add_body_p(doc, "2) Регистрационный метод – ведение журналов входного контроля ОТК, протоколирование результатов 72-часового технологического прогона шкафов;", bold_prefix="   • ")
    add_body_p(doc, "3) Статистический метод – контроль технологической стабильности сборочного процесса по картам Шухарта;", bold_prefix="   • ")
    add_body_p(doc, "4) Экспертный метод – приемочная комиссия FAT с подписанием протокола заводских испытаний заказчиком.", bold_prefix="   • ")

    add_body_p(doc, "Расчет интегрального показателя качества этапа Производства:")
    add_body_p(doc, "Для Варианта 1 (REGUL): q~1 = 0,10/0,04=2,50; q~2 = 72/48=1,50; q~3 = 100/98=1,02; q~4 = 5,0/4,5=1,11; q~5 = 250/50=5,0; q~6 = 4/3=1,33; q~7 = 6,0/3,5=1,71. Итого:")
    add_body_p(doc, "I_произв(Вар.1) = 0,20*2,50 + 0,15*1,50 + 0,20*1,02 + 0,10*1,11 + 0,15*5,0 + 0,10*1,33 + 0,10*1,71 = 2,094.", bold_prefix="   • ")
    add_body_p(doc, "Для Варианта 2 (Yokogawa): q~1 = 0,10/0,02=5,0; q~2 = 96/48=2,0; q~3 = 100/98=1,02; q~4 = 5,2/4,5=1,155; q~5 = 300/50=6,0; q~6 = 4/3=1,33; q~7 = 6,0/9,0=0,67. Итого:")
    add_body_p(doc, "I_произв(Вар.2) = 0,20*5,0 + 0,15*2,0 + 0,20*1,02 + 0,10*1,155 + 0,15*6,0 + 0,10*1,33 + 0,10*0,67 = 2,620.", bold_prefix="   • ")
    add_body_p(doc, "Оба варианта успешно преодолели пороговый критерий шлюза Gate 3 (I >= 0,85). Существенным практическим преимуществом Варианта 1 является срок поставки (3,5 месяца против 9 месяцев у зарубежного аналога ввиду ограничений логистики). Решение на шлюзе Gate 3: Go (Оборудование готово к отгрузке и монтажу).")

    # ------------------ ЭТАП 4: ЭКСПЛУАТАЦИЯ ------------------
    add_heading_2(doc, "3.4. Этап 4: Эксплуатация (Gate 4 «Шлюз ввода в промышленную эксплуатацию и надежность»)")
    add_body_p(doc, "На стадии эксплуатации оценивается фактическая безотказность системы на объекте, отсутствие ложных срабатываний ПАЗ, оперативность устранения инцидентов и достаточность ЗИП. В таблице 4 приведены показатели этапа эксплуатации.")

    t4 = doc.add_table(rows=1, cols=6)
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t4)

    for j, cell in enumerate(t4.rows[0].cells):
        cell.width = t1_widths[j]
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        set_cell_shading(cell, "EAEAEA")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        format_paragraph(p, space_before=2, space_after=2, line_spacing=1.0, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run(t1_headers[j])
        r.font.name = 'Times New Roman'
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)

    e_data = [
        ("1", "Фактическая наработка T0, тыс. ч", "0,25", "50", "68", "88"),
        ("2", "Коэффициент готовности Kг, факт", "0,20", "0,9995", "0,99985", "0,99995"),
        ("3", "Число ложных срабатываний ПАЗ / год", "0,15", "0,5", "0,1", "0,05"),
        ("4", "Время локализации инцидента, мин", "0,10", "30", "15", "12"),
        ("5", "Достаточность комплекта ЗИП-О Pзип", "0,10", "0,98", "0,99", "0,995"),
        ("6", "Доля отказов по вине персонала, %", "0,10", "5,0", "2,0", "1,5"),
        ("7", "Удовлетворенность персонала (балл)", "0,10", "8,0", "9,0", "9,4")
    ]

    for row_data in e_data:
        row = t4.add_row()
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            cell.width = t1_widths[j]
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            align = WD_ALIGN_PARAGRAPH.LEFT if j == 1 else WD_ALIGN_PARAGRAPH.CENTER
            format_paragraph(p, space_before=1, space_after=1, line_spacing=1.0, first_indent=0, align=align)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(0, 0, 0)

    add_body_p(doc, "Методы обеспечения выбранных показателей на этапе Эксплуатации:")
    add_body_p(doc, "1) Статистический метод – автоматизированная регистрация и аналитическая обработка событий в historian SCADA с фиксацией времени восстановления Tв;", bold_prefix="   • ")
    add_body_p(doc, "2) Регистрационный метод – ведение электронного паспорта оборудования и журнала инцидентов дежурного инженера АСУ ТП;", bold_prefix="   • ")
    add_body_p(doc, "3) Измерительный метод – контроль дрейфа метрологических характеристик датчиков давления и температуры по протоколу HART без демонтажа с установки;", bold_prefix="   • ")
    add_body_p(doc, "4) Экспертный метод – периодическая аттестация операторов на тренажере и аудит промышленной безопасности.", bold_prefix="   • ")

    add_body_p(doc, "Расчет интегрального показателя качества этапа Эксплуатации:")
    add_body_p(doc, "Для Варианта 1 (REGUL): q~1 = 68/50=1,36; q~2 = 0,99985/0,9995=1,00035; q~3 = 0,5/0,1=5,0; q~4 = 30/15=2,0; q~5 = 0,99/0,98=1,01; q~6 = 5,0/2,0=2,50; q~7 = 9,0/8,0=1,125. Итого:")
    add_body_p(doc, "I_экспл(Вар.1) = 0,25*1,36 + 0,20*1,00035 + 0,15*5,0 + 0,10*2,0 + 0,10*1,01 + 0,10*2,50 + 0,10*1,125 = 1,954.", bold_prefix="   • ")
    add_body_p(doc, "Для Варианта 2 (Yokogawa): q~1 = 88/50=1,76; q~2 = 0,99995/0,9995=1,00045; q~3 = 0,5/0,05=10,0; q~4 = 30/12=2,50; q~5 = 0,995/0,98=1,015; q~6 = 5,0/1,5=3,33; q~7 = 9,4/8,0=1,175. Итого:")
    add_body_p(doc, "I_экспл(Вар.2) = 0,25*1,76 + 0,20*1,00045 + 0,15*10,0 + 0,10*2,50 + 0,10*1,015 + 0,10*3,33 + 0,10*1,175 = 2,942.", bold_prefix="   • ")
    add_body_p(doc, "Решение на шлюзе Gate 4: Go (Система допущена к постоянной промышленной эксплуатации на опасном объекте).")

    # ------------------ СВОДНОЕ СРАВНЕНИЕ И ШЛЮЗЫ ------------------
    add_heading_1(doc, "4. СВОДНЫЙ АНАЛИЗ И СРАВНЕНИЕ ВАРИАНТОВ НА ШЛЮЗАХ")
    add_body_p(doc, "В таблице 5 представлены сводные значения интегральных показателей качества I_stage и удельных показателей эффективности I_уд на всех четырех контрольных шлюзах жизненного цикла.")

    t5 = doc.add_table(rows=1, cols=6)
    t5.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t5)

    t5_headers = ["Шлюз (Gate)", "Контролируемый этап", "Порог шлюза", "Вариант 1 (REGUL)", "Вариант 2 (Yokogawa)", "Решение комиссии"]
    t5_widths = [Mm(25), Mm(45), Mm(25), Mm(28), Mm(28), Mm(25)]
    for j, cell in enumerate(t5.rows[0].cells):
        cell.width = t5_widths[j]
        set_cell_margins(cell, top=80, bottom=80, left=80, right=80)
        set_cell_shading(cell, "EAEAEA")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        format_paragraph(p, space_before=2, space_after=2, line_spacing=1.0, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run(t5_headers[j])
        r.font.name = 'Times New Roman'
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)

    gate_summary = [
        ("Gate 1", "Маркетинг (ТЗ и концепция)", ">= 0,85", "1,215 (I_уд=0,0289)", "1,301 (I_уд=0,0167)", "Go (Одобрено)"),
        ("Gate 2", "Разработка (проект и РД)", ">= 0,85", "1,299", "1,448", "Go (Одобрено)"),
        ("Gate 3", "Производство (заводской FAT)", ">= 0,85", "2,094", "2,620", "Go (Одобрено)"),
        ("Gate 4", "Эксплуатация (ввод в действие)", ">= 0,85", "1,954", "2,942", "Go (Одобрено)")
    ]

    for row_data in gate_summary:
        row = t5.add_row()
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            cell.width = t5_widths[j]
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            format_paragraph(p, space_before=1, space_after=1, line_spacing=1.0, first_indent=0, align=WD_ALIGN_PARAGRAPH.CENTER)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(0, 0, 0)

    add_body_p(doc, "Анализ сводных результатов показывает, что хотя зарубежный комплекс (Вариант 2) имеет несколько более высокие абсолютные значения интегрального критерия за счет тройного мажоритирования (TMR), отечественный комплекс (Вариант 1 на базе ПЛК REGUL R500):")
    add_body_p(doc, "1) Полностью удовлетворяет всем пороговым критериям надежности и безопасности (SIL 3, Kг = 0,9998, T0 = 68 000 ч) на всех четырех шлюзах гейтовой системы;", bold_prefix="   1. ")
    add_body_p(doc, "2) Обладает превосходством по удельному критерию экономической эффективности I_уд в 1,73 раза (стоимость внедрения 42 млн руб. против 78 млн руб.);", bold_prefix="   2. ")
    add_body_p(doc, "3) Имеет критически важные преимущества по доступности оперативной сервисной поддержки 24/7 в РФ, отсутствию санкционных рисков и сокращенному сроку поставки (3,5 месяца против 9 месяцев).", bold_prefix="   3. ")
    add_body_p(doc, "В связи с этим для реализации на объекте ректификации углеводородов окончательно рекомендуется Вариант 1.")

    # ------------------ ЗАКЛЮЧЕНИЕ ------------------
    add_heading_1(doc, "5. ЗАКЛЮЧЕНИЕ")
    add_body_p(doc, "В ходе выполнения лабораторной работы № 5 получены следующие основные результаты:")
    add_body_p(doc, "1. Изучена концепция и принципы функционирования гейтовой системы управления качеством и надежностью (Stage-Gate) сложных АСОИУ на стадиях жизненного цикла;", bold_prefix="   1) ")
    add_body_p(doc, "2. Разработана сквозная структурно-функциональная схема гейтовой системы с выделением четырех стадий (Маркетинг, Разработка, Производство, Эксплуатация) и соответствующих контрольных шлюзов (Gate 1 – Gate 4), регламентированы критерии принятия решений (Go, Kill, Hold, Recycle);", bold_prefix="   2) ")
    add_body_p(doc, "3. Сформирована научно обоснованная номенклатура единичных показателей надежности, эргономики, технологичности и безопасности для непрерывного технологического объекта – АСУ ТП ректификационной колонны разделения углеводородов;", bold_prefix="   3) ")
    add_body_p(doc, "4. Произведен расчет нормированных комплексных (интегральных) критериев качества этапов I_stage и удельных критериев эффективности I_уд для двух альтернативных аппаратных платформ;", bold_prefix="   4) ")
    add_body_p(doc, "5. Обоснованы методы контроля и обеспечения показателей (социологический, аналитический расчетный, имитационный на HIL-симуляторе, инструментальный измерительный, статистический сбор SOE в SCADA);", bold_prefix="   5) ")
    add_body_p(doc, "6. Доказано, что отечественный аппаратно-программный комплекс РСУ/ПАЗ обеспечивает надежное прохождение всех шлюзов системы качества и является оптимальным по совокупности параметров надежности, безопасности и импортонезависимости.", bold_prefix="   6) ")

    # ------------------ СПИСОК ИСТОЧНИКОВ ------------------
    add_heading_1(doc, "6. СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    sources = [
        "ГОСТ 15467-79. Управление качеством продукции. Основные понятия. Термины и определения. – М.: Издательство стандартов, 2002. – 19 с.",
        "ГОСТ 24.701-86. Надежность автоматизированных систем управления. Основные положения. – М.: Стандартинформ, 2006. – 18 с.",
        "ГОСТ 27.002-2015. Надежность в технике. Термины и определения. – М.: Стандартинформ, 2016. – 28 с.",
        "ГОСТ Р МЭК 61508-1-2012. Функциональная безопасность систем электрических, электронных, программируемых электронных, связанных с безопасностью. Часть 1. Общие требования. – М.: Стандартинформ, 2013. – 42 с.",
        "ГОСТ 34.601-90. Автоматизированные системы. Стадии создания. – М.: Стандартинформ, 2009. – 8 с.",
        "РД 50-680-89. Методические указания. Автоматизированные системы. Основные положения. – М.: Издательство стандартов, 1990. – 14 с.",
        "Купер Р.Д. Модель Stage-Gate: новые правила создания успешных продуктов. – М.: Манн, Иванов и Фербер, 2017. – 368 с.",
        "Методические указания к лабораторной работе № 5 «Разработка показателей качества АСОИУ для гейтовой системы управления качеством на стадиях ЖЦ» / Сост.: Ризванов К.А. – Уфа: УГАТУ, 2021. – 4 с."
    ]
    for s in sources:
        add_body_p(doc, s, bold_prefix="• ")

    doc.save(REPORT_PATH)
    print(f"[OK] Успешно сформирован отчет по ЛР5: {REPORT_PATH}")

if __name__ == '__main__':
    build_report_lr5()

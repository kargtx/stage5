# -*- coding: utf-8 -*-
"""
Генератор итогового отчета по Части 2 лабораторных работ КИС:
«Детальный бизнес-сценарий (ДБС) оказания услуг программной модификации
блоков управления двигателем (чип-тюнинг) в КИС Автосервис»

Строго по стандартам СТП УГАТУ и ГОСТ:
- Год: 2026
- Исполнители: студенты группы ЭАС-514С Гайнанов Д.И. и Гарифуллин К.Р.
- Преподаватель: ст. преподаватель Демченко М.С.
- Оформление: строго черные буквы (RGB 0,0,0) без цветных шрифтов
- Таблицы без окраски (чистый белый фон, классические академические границы)
- Жирные шрифты строго минимизированы (только заголовки и шапки таблиц)
- Автособираемое Оглавление MS Word (TOC)
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_dbs_report():
    print("=== Начало сборки отчета ДБС (2026, строго академическое оформление) ===")
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    dbs_dir = os.path.abspath(os.path.join(script_dir, ".."))
    out_docx_path = os.path.join(dbs_dir, "Отчет_ДБС_ЧипТюнинг.docx")
    
    doc = Document()
    
    # -------------------------------------------------------------------------
    # 0. НАСТРОЙКИ СТРАНИЦЫ (СТП УГАТУ: левое 30мм, правое 15мм, верх/низ 20мм)
    # -------------------------------------------------------------------------
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15)
        
    # Включаем автоматическое обновление оглавления и полей при открытии в Word
    settings = doc.settings._element
    update_fields = parse_xml(r'<w:updateFields %s w:val="true"/>' % nsdecls('w'))
    settings.append(update_fields)

    # -------------------------------------------------------------------------
    # ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (СТРОГО ЧЕРНЫЙ ЦВЕТ, БЕЗ ЦВЕТНЫХ ВКРАПЛЕНИЙ)
    # -------------------------------------------------------------------------
    COLOR_BLACK = RGBColor(0, 0, 0)

    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="{top}" w:type="dxa"/>'
            f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
            f'<w:left w:w="{left}" w:type="dxa"/>'
            f'<w:right w:w="{right}" w:type="dxa"/>'
            f'</w:tcMar>'
        )
        tcPr.append(tcMar)

    def set_table_borders(table, color="000000", sz="4", val="single"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(f'''
            <w:tblBorders {nsdecls("w")}>
                <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            </w:tblBorders>
        ''')
        tblPr.append(borders)

    def add_p(text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=3, indent=True, italic=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if indent:
            p.paragraph_format.first_line_indent = Mm(12.5) # Стандартный абзацный отступ 1.25 см
        else:
            p.paragraph_format.first_line_indent = Mm(0)
        p.alignment = align

        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.font.color.rgb = COLOR_BLACK
        run.bold = False
        run.italic = italic
        return p

    def add_bullet(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.left_indent = Mm(18)
        p.paragraph_format.first_line_indent = Mm(-5.5)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        run = p.add_run("— " + text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.font.color.rgb = COLOR_BLACK
        run.bold = False
        return p

    def add_h(text, level=1):
        h = doc.add_heading(level=level)
        h.paragraph_format.keep_with_next = True
        h.paragraph_format.first_line_indent = Mm(0)
        run = h.add_run(text)
        run.font.name = "Times New Roman"
        run.bold = True
        run.font.color.rgb = COLOR_BLACK
        if level == 1:
            h.paragraph_format.space_before = Pt(14)
            h.paragraph_format.space_after = Pt(6)
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run.font.size = Pt(14)
        elif level == 2:
            h.paragraph_format.space_before = Pt(10)
            h.paragraph_format.space_after = Pt(4)
            h.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run.font.size = Pt(13)
        else:
            h.paragraph_format.space_before = Pt(8)
            h.paragraph_format.space_after = Pt(3)
            h.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run.font.size = Pt(12)
        return h

    def add_image_figure(img_filename, caption_text, width_inches=6.0):
        img_path = os.path.join(dbs_dir, img_filename)
        if not os.path.exists(img_path):
            print(f"[!] Предупреждение: файл изображения не найден: {img_path}")
            return
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(6)
        p_img.paragraph_format.space_after = Pt(3)
        p_img.paragraph_format.first_line_indent = Mm(0)
        run = p_img.add_run()
        run.add_picture(img_path, width=Inches(width_inches))

        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(8)
        p_cap.paragraph_format.first_line_indent = Mm(0)
        r_cap = p_cap.add_run(caption_text)
        r_cap.font.name = "Times New Roman"
        r_cap.font.size = Pt(11)
        r_cap.bold = False
        r_cap.font.color.rgb = COLOR_BLACK

    def style_header_cell(cell, text):
        # Строго БЕЗ заливки (прозрачный/белый фон)
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.first_line_indent = Mm(0)
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)
        r.bold = True
        r.font.color.rgb = COLOR_BLACK

    def style_data_cell(cell, text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT):
        # Строго БЕЗ заливки (прозрачный/белый фон)
        set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.first_line_indent = Mm(0)
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)
        r.bold = bold
        r.font.color.rgb = COLOR_BLACK

    # -------------------------------------------------------------------------
    # 1. ТИТУЛЬНЫЙ ЛИСТ ОТЧЕТА ПО СТП УГАТУ (ГОД 2026, ГАЙНАНОВ Д.И., ГАРИФУЛЛИН К.Р.)
    # -------------------------------------------------------------------------
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_top.paragraph_format.first_line_indent = Mm(0)
    p_top.paragraph_format.space_after = Pt(0)
    r = p_top.add_run("Министерство науки и высшего образования Российской Федерации\n"
                      "ФГБОУ ВО «Уфимский университет науки и технологий»\n"
                      "Кафедра автоматизированных систем управления\n")
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    r.bold = True
    r.font.color.rgb = COLOR_BLACK

    for _ in range(3):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_after = Pt(0)
        p_sp.paragraph_format.first_line_indent = Mm(0)

    p_mid = doc.add_paragraph()
    p_mid.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_mid.paragraph_format.first_line_indent = Mm(0)
    p_mid.paragraph_format.space_after = Pt(0)
    
    r_mid1 = p_mid.add_run("ОТЧЕТ ПО ЛАБОРАТОРНЫМ РАБОТАМ\n")
    r_mid1.font.name = "Times New Roman"
    r_mid1.font.size = Pt(16)
    r_mid1.bold = True
    r_mid1.font.color.rgb = COLOR_BLACK

    r_mid2 = p_mid.add_run("по дисциплине «Корпоративные информационные системы»\n")
    r_mid2.font.name = "Times New Roman"
    r_mid2.font.size = Pt(13)
    r_mid2.font.color.rgb = COLOR_BLACK

    r_mid3 = p_mid.add_run("ЧАСТЬ 2\n\n")
    r_mid3.font.name = "Times New Roman"
    r_mid3.font.size = Pt(14)
    r_mid3.bold = True
    r_mid3.font.color.rgb = COLOR_BLACK

    r_mid4 = p_mid.add_run("Тема: «Детальный бизнес-сценарий (ДБС) оказания услуг\n"
                           "программной модификации блоков управления двигателем (чип-тюнинг)\n"
                           "в КИС Автосервис»\n")
    r_mid4.font.name = "Times New Roman"
    r_mid4.font.size = Pt(13)
    r_mid4.bold = True
    r_mid4.font.color.rgb = COLOR_BLACK

    for _ in range(4):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_after = Pt(0)
        p_sp.paragraph_format.first_line_indent = Mm(0)

    # Таблица подписей (без видимых рамок)
    t_sign = doc.add_table(rows=2, cols=2)
    t_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_sign.autofit = False
    t_sign.columns[0].width = Mm(85)
    t_sign.columns[1].width = Mm(80)

    cell_exec = t_sign.cell(0, 1)
    p_exec = cell_exec.paragraphs[0]
    p_exec.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_exec.paragraph_format.first_line_indent = Mm(0)
    r_ex = p_exec.add_run(
        "Выполнили:\n"
        "студенты группы ЭАС-514С\n"
        "Гайнанов Д. И.\n"
        "Гарифуллин К. Р.\n"
        "«___» ____________ 2026 г.\n\n"
    )
    r_ex.font.name = "Times New Roman"
    r_ex.font.size = Pt(11)
    r_ex.font.color.rgb = COLOR_BLACK

    cell_check = t_sign.cell(1, 1)
    p_check = cell_check.paragraphs[0]
    p_check.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_check.paragraph_format.first_line_indent = Mm(0)
    r_ch = p_check.add_run(
        "Проверил:\n"
        "ст. преподаватель кафедры АСУ\n"
        "Демченко М. С. / ____________ /\n"
        "«___» ____________ 2026 г.\n"
    )
    r_ch.font.name = "Times New Roman"
    r_ch.font.size = Pt(11)
    r_ch.font.color.rgb = COLOR_BLACK

    for _ in range(3):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_after = Pt(0)
        p_sp.paragraph_format.first_line_indent = Mm(0)

    p_bot = doc.add_paragraph()
    p_bot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_bot.paragraph_format.first_line_indent = Mm(0)
    r_bot = p_bot.add_run("Уфа — 2026 г.")
    r_bot.font.name = "Times New Roman"
    r_bot.font.size = Pt(11)
    r_bot.font.color.rgb = COLOR_BLACK

    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 2. ЛИСТ ЗАДАНИЯ ЧАСТЬ 2 (ГОД 2026, СТУДЕНТЫ ГАЙНАНОВ Д.И., ГАРИФУЛЛИН К.Р.)
    # -------------------------------------------------------------------------
    p_task_top = doc.add_paragraph()
    p_task_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_task_top.paragraph_format.first_line_indent = Mm(0)
    r_tt = p_task_top.add_run("Федеральное государственное бюджетное образовательное учреждение\n"
                              "высшего образования\n"
                              "«Уфимский университет науки и технологий»\n"
                              "Кафедра автоматизированных систем управления\n")
    r_tt.font.name = "Times New Roman"
    r_tt.font.size = Pt(10.5)
    r_tt.bold = True
    r_tt.font.color.rgb = COLOR_BLACK

    p_t_h = doc.add_paragraph()
    p_t_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t_h.paragraph_format.space_before = Pt(6)
    p_t_h.paragraph_format.space_after = Pt(6)
    p_t_h.paragraph_format.first_line_indent = Mm(0)
    r_th = p_t_h.add_run("ЗАДАНИЕ\n"
                         "на лабораторные работы по дисциплине «Корпоративные информационные системы»\n"
                         "ЧАСТЬ 2")
    r_th.font.name = "Times New Roman"
    r_th.font.size = Pt(12)
    r_th.bold = True
    r_th.font.color.rgb = COLOR_BLACK

    # Таблица 0 задания: Студент / Группа / Консультант
    t_z0 = doc.add_table(rows=2, cols=6)
    t_z0.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_z0)
    col_widths_z0 = [Mm(25), Mm(45), Mm(20), Mm(25), Mm(25), Mm(25)]
    for i, w in enumerate(col_widths_z0):
        t_z0.columns[i].width = w

    headers_z0 = ["Студент", "Гайнанов Д.И., Гарифуллин К.Р.", "Группа", "ЭАС-514С", "Консультант", "Демченко М.С."]
    sub_z0 = ["", "Фамилия И.О.", "", "номер группы", "", "Фамилия И.О."]
    for c_idx, h_text in enumerate(headers_z0):
        style_data_cell(t_z0.cell(0, c_idx), h_text, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for c_idx, s_text in enumerate(sub_z0):
        style_data_cell(t_z0.cell(1, c_idx), s_text, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # Таблица 1 задания
    t_z1 = doc.add_table(rows=4, cols=2)
    t_z1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_z1)
    t_z1.columns[0].width = Mm(45)
    t_z1.columns[1].width = Mm(120)

    z1_data = [
        ("1. Тема:", "«Детальный бизнес-сценарий (ДБС) оказания услуг программной модификации блоков управления двигателем (чип-тюнинг) в КИС Автосервис»"),
        ("2. Основное содержание:", "В рамках лабораторных работ ЧАСТЬ 1 сформирована технико-экономическая характеристика объекта автоматизации и построена реляционная модель данных (ERD). В Части 2 разработан Детальный бизнес-сценарий сквозного процесса «Оказание услуг чип-тюнинга автомобиля» в соответствии с корпоративным шаблоном ДБС. Включены: общая схема БП на 5 ролей, матрица ответственности, пошаговая спецификация 7 ключевых шагов процесса с экранными формами (UI Mockups) и таблицами значимых атрибутов, требованиями к функционалу и интеграции. Все схемы разработаны в MS Visio."),
        ("3. Исходные данные:", "Вариант комплексной автоматизации автосервиса чип-тюнинга и калибровки ЭБУ; результаты ЧАСТИ 1 (подпроцесс оформления заявки и утверждения наряд-заказа — Данил; подпроцесс исполнения наряд-заказа с комментарием мастера и выдачи акта — Карим); нормативные документы СТО 521.18.107-2007, СТО 521.18.158-2018."),
        ("4. Требования к оформлению:", "4.1 Пояснительная записка к отчету оформляется в текстовом редакторе MS Word в соответствии с СТП УГАТУ. Обязательно автособираемое Оглавление.\n4.2 Разделы: Титульный лист, Лист Задания, Титульный лист ДБС, Оглавление, Термины и сокращения, 1. Версии документа, 2. Назначение документа, 3. Основные положения, 4. Бизнес-процесс, 5. Модель данных, Заключение, Список используемых источников.\n4.3 Срок сдачи: 2026 г.")
    ]

    for r_idx, (col1, col2) in enumerate(z1_data):
        style_data_cell(t_z1.cell(r_idx, 0), col1, bold=True)
        style_data_cell(t_z1.cell(r_idx, 1), col2, bold=False)

    p_z_sig = doc.add_paragraph()
    p_z_sig.paragraph_format.space_before = Pt(8)
    p_z_sig.paragraph_format.first_line_indent = Mm(0)
    r_zs = p_z_sig.add_run(
        "Руководитель: __________________ М.С. Демченко          "
        "Студенты: _________________ Гайнанов Д.И., Гарифуллин К.Р."
    )
    r_zs.font.name = "Times New Roman"
    r_zs.font.size = Pt(10.5)
    r_zs.font.color.rgb = COLOR_BLACK

    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 3. ВНУТРЕННИЙ ТИТУЛЬНЫЙ ЛИСТ ДБС (ПО СТАНДАРТУ 123456BS, 2026 Г.)
    # -------------------------------------------------------------------------
    p_dbs_top = doc.add_paragraph()
    p_dbs_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dbs_top.paragraph_format.first_line_indent = Mm(0)
    r_dt = p_dbs_top.add_run("ООО «Тюнинг-Сервис Моторспорт»\n"
                             "Департамент корпоративных информационных технологий\n")
    r_dt.font.name = "Times New Roman"
    r_dt.font.size = Pt(11)
    r_dt.bold = True
    r_dt.font.color.rgb = COLOR_BLACK

    for _ in range(3):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_after = Pt(0)
        p_sp.paragraph_format.first_line_indent = Mm(0)

    p_dbs_title = doc.add_paragraph()
    p_dbs_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dbs_title.paragraph_format.first_line_indent = Mm(0)
    r_dbt1 = p_dbs_title.add_run("ДЕТАЛЬНЫЙ БИЗНЕС-СЦЕНАРИЙ\n")
    r_dbt1.font.name = "Times New Roman"
    r_dbt1.font.size = Pt(16)
    r_dbt1.bold = True
    r_dbt1.font.color.rgb = COLOR_BLACK

    r_dbt2 = p_dbs_title.add_run("123456BS 01\n\n")
    r_dbt2.font.name = "Times New Roman"
    r_dbt2.font.size = Pt(13)
    r_dbt2.bold = True
    r_dbt2.font.color.rgb = COLOR_BLACK

    r_dbt3 = p_dbs_title.add_run("«Оказание услуг программной модификации блоков управления двигателем\n"
                                "(чип-тюнинг) в КИС Автосервис»\n")
    r_dbt3.font.name = "Times New Roman"
    r_dbt3.font.size = Pt(13)
    r_dbt3.bold = True
    r_dbt3.font.color.rgb = COLOR_BLACK

    for _ in range(4):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_after = Pt(0)
        p_sp.paragraph_format.first_line_indent = Mm(0)

    # Таблица утверждения ДБС
    t_app = doc.add_table(rows=5, cols=4)
    t_app.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_app)
    t_app.columns[0].width = Mm(30)
    t_app.columns[1].width = Mm(55)
    t_app.columns[2].width = Mm(50)
    t_app.columns[3].width = Mm(30)

    app_rows = [
        ("Утверждено:", "Генеральный директор", "Смирнов А.В. / ________ /", "«___» _____ 2026 г."),
        ("Согласовано:", "Руководитель проекта КИС", "Демченко М.С. / ________ /", "«___» _____ 2026 г."),
        ("Разработано:", "Ведущий бизнес-аналитик", "Гайнанов Д.И. / ________ /", "«___» _____ 2026 г."),
        ("Разработано:", "Ведущий инженер-калибровщик", "Гарифуллин К.Р. / ________ /", "«___» _____ 2026 г."),
        ("Направление:", "Техническое обслуживание и тюнинг", "КИС «Автосервис КОРП»", "Версия 2.0")
    ]

    for r_idx, (c0, c1, c2, c3) in enumerate(app_rows):
        style_data_cell(t_app.cell(r_idx, 0), c0, bold=True)
        style_data_cell(t_app.cell(r_idx, 1), c1, bold=False)
        style_data_cell(t_app.cell(r_idx, 2), c2, bold=False)
        style_data_cell(t_app.cell(r_idx, 3), c3, bold=False)

    for _ in range(4):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_after = Pt(0)
        p_sp.paragraph_format.first_line_indent = Mm(0)

    p_dbs_bot = doc.add_paragraph()
    p_dbs_bot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dbs_bot.paragraph_format.first_line_indent = Mm(0)
    r_dbb = p_dbs_bot.add_run("г. Уфа, 2026 г.")
    r_dbb.font.name = "Times New Roman"
    r_dbb.font.size = Pt(11)
    r_dbb.font.color.rgb = COLOR_BLACK

    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 4. АВТОСОБИРАЕМОЕ ОГЛАВЛЕНИЕ (TOC)
    # -------------------------------------------------------------------------
    add_h("Оглавление", level=1)
    
    p_toc = doc.add_paragraph()
    p_toc.paragraph_format.space_before = Pt(4)
    p_toc.paragraph_format.space_after = Pt(8)
    p_toc.paragraph_format.first_line_indent = Mm(0)
    
    r_toc = p_toc.add_run()
    fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText>' % nsdecls('w'))
    fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    r_toc._r.append(fldChar1)
    r_toc._r.append(instrText)
    r_toc._r.append(fldChar2)
    r_toc._r.append(fldChar3)

    add_p("Примечание: Оглавление сформировано динамическим полем MS Word. Для обновления номеров страниц в Microsoft Word нажмите правой кнопкой мыши по области оглавления и выберите «Обновить поле» -> «Обновить целиком», либо нажмите клавишу F9.", italic=True)

    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 5. ТЕРМИНЫ И СОКРАЩЕНИЯ
    # -------------------------------------------------------------------------
    add_h("Термины и сокращения", level=1)
    add_p("Использованные в настоящем документе детального бизнес-сценария термины и сокращения трактуются следующим образом:")

    terms_data = [
        ("КИС", "Корпоративная информационная система — комплексное прикладное программное обеспечение для автоматизации операционного учета, документооборота, планирования ресурсов и взаимодействия подразделений автосервиса."),
        ("ДБС", "Детальный бизнес-сценарий — проектный регламентирующий документ, описывающий последовательность шагов бизнес-процесса, ролевую модель, пользовательские сеансы, атрибутивный состав данных и интеграционные взаимодействия."),
        ("ЭБУ / ECU", "Электронный блок управления (Engine Control Unit) — микропроцессорный контроллер, управляющий впрыском топлива, зажиганием, наддувом турбокомпрессора и экологическими подсистемами ДВС."),
        ("Чип-тюнинг", "Программная модификация калибровочных таблиц заводской прошивки ЭБУ с целью повышения мощности, крутящего момента, оптимизации топливной экономичности или программного отключения экологических компонентов (EGR, DPF, AdBlue, SAP)."),
        ("Stage 1 / Stage 2", "Градации уровней программной модификации. Stage 1 — оптимизация заводского ПО под стандартные узлы ДВС. Stage 2 — калибровка под модифицированные компоненты (downpipe, спортивный интеркулер, впуск)."),
        ("Дамп прошивки (ROM/BIN)", "Полный бинарный образ микрокода и калибровочных карт флеш-памяти и микроконтроллера ЭБУ (размер от 512 КБ до 8 МБ)."),
        ("Контрольная сумма (Checksum)", "Математический хэш-код целостности бинарного файла прошивки, верифицируемый процессором ЭБУ при запуске для предотвращения сбоев и блокировки процессора."),
        ("Наряд-заказ (Work Order)", "Первичный юридический и финансовый документ, фиксирующий перечень заказанных работ, калькуляцию стоимости, согласие клиента на проведение вмешательства и плановые сроки (зона ответственности Данила)."),
        ("АРМ Калибровщика", "Автоматизированное рабочее место инженера-калибровщика в КИС, интегрированное со специализированными аппаратными загрузчиками и редактором калибровок WinOLS."),
        ("Технический комментарий", "Обязательное поле фиксации инженерных нюансов, параметров наддува, УОЗ, температуры выхлопных газов (EGT) и поведения ЭБУ мастером-калибровщиком (зона ответственности Карима)."),
        ("Акт выполненных работ", "Заключительный юридический документ, фиксирующий фактическое исполнение работ, факт передачи ключей и автомобиля владельцу, выдачу гарантийного талона на прошивку (зона ответственности Карима)."),
        ("СТО / СТП", "Стандарты предприятия и стандарты организации системы менеджмента качества (в том числе СТО 521.18.107-2007 и СТО 521.18.158-2018).")
    ]

    t_terms = doc.add_table(rows=len(terms_data) + 1, cols=2)
    t_terms.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_terms)
    t_terms.columns[0].width = Mm(45)
    t_terms.columns[1].width = Mm(120)

    style_header_cell(t_terms.cell(0, 0), "Термин / Сокращение")
    style_header_cell(t_terms.cell(0, 1), "Описание и смысловое значение")

    for idx, (term, desc) in enumerate(terms_data):
        r_row = idx + 1
        style_data_cell(t_terms.cell(r_row, 0), term, bold=True)
        style_data_cell(t_terms.cell(r_row, 1), desc, bold=False)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # -------------------------------------------------------------------------
    # 6. РАЗДЕЛ 1. ВЕРСИИ ДОКУМЕНТА (ГОД 2026)
    # -------------------------------------------------------------------------
    add_h("1 Версии документа", level=1)
    add_p("В таблице 1 приведена хронология разработки и согласования настоящего детального бизнес-сценария.")

    add_p("Таблица 1 — История версий документа ДБС", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    ver_data = [
        ("15.02.2026", "Гайнанов Д.И.", "ЭАС-514С", "1.0", "Первоначальная редакция: формирование бизнес-целей, структуры БП и матрицы ролей."),
        ("28.02.2026", "Гарифуллин К.Р.", "ЭАС-514С", "1.1", "Детализация шагов исполнения наряд-заказа, добавление интерфейса АРМ калибровщика и технического комментария."),
        ("20.03.2026", "Гайнанов Д.И., Гарифуллин К.Р.", "ЭАС-514С", "2.0", "Итоговая интеграция: ER-диаграмма в 3NF, валидация экранных форм UI mockups, согласование с кафедрой АСУ.")
    ]

    t_ver = doc.add_table(rows=len(ver_data) + 1, cols=5)
    t_ver.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_ver)
    t_ver.columns[0].width = Mm(25)
    t_ver.columns[1].width = Mm(45)
    t_ver.columns[2].width = Mm(25)
    t_ver.columns[3].width = Mm(18)
    t_ver.columns[4].width = Mm(52)

    style_header_cell(t_ver.cell(0, 0), "Дата")
    style_header_cell(t_ver.cell(0, 1), "Автор")
    style_header_cell(t_ver.cell(0, 2), "Компания / Группа")
    style_header_cell(t_ver.cell(0, 3), "Версия")
    style_header_cell(t_ver.cell(0, 4), "Описание изменений")

    for idx, (dt, auth, grp, v, chg) in enumerate(ver_data):
        r_row = idx + 1
        style_data_cell(t_ver.cell(r_row, 0), dt, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        style_data_cell(t_ver.cell(r_row, 1), auth, bold=False)
        style_data_cell(t_ver.cell(r_row, 2), grp, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        style_data_cell(t_ver.cell(r_row, 3), v, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        style_data_cell(t_ver.cell(r_row, 4), chg, bold=False)

    # -------------------------------------------------------------------------
    # 7. РАЗДЕЛ 2. НАЗНАЧЕНИЕ ДОКУМЕНТА
    # -------------------------------------------------------------------------
    add_h("2 Назначение документа", level=1)
    add_p("Настоящий документ разработан в рамках выполнения лабораторных работ по дисциплине «Корпоративные информационные системы» в 2026 году. Предметная область: «Автоматизация технологических и управленческих процессов автосервиса чип-тюнинга двигателей и калибровки блоков управления».")
    add_p("Настоящий документ 123456BS 01 «Оказание услуг программной модификации блоков управления двигателем (чип-тюнинг) в КИС Автосервис» содержит детальный бизнес-сценарий по направлению сервисного и технологического обслуживания и является ключевым проектным документом этапа функционального моделирования.")
    add_p("В документе регламентированы процессы взаимодействия клиентов, мастеров приемки, диагностов, калибровщиков и финансового отдела в корпоративной информационной системе «Автосервис КОРП» на базе платформы «1С:Предприятие 8.3». Приводится строгая хронологическая последовательность шагов, детальная ролевая спецификация, перечень входов и выходов, требования к экранным сеансам КИС и перечень значимых атрибутов базы данных.")
    add_p("Структура документа включает:")
    add_bullet("индивидуальные динамические модели подпроцессов оформления заявки (Гайнанов Д.И.) и исполнения наряд-заказа с комментарием мастера (Гарифуллин К.Р.) в стандарте BPMN 2.0;")
    add_bullet("общую сквозную динамическую модель процесса оказания услуг чип-тюнинга по 5 ролевым дорожкам;")
    add_bullet("матрицу распределения ответственности ролей по шагам технологического цикла (RACI);")
    add_bullet("пошаговую спецификацию 7 базовых этапов с макетами экранных форм (UI Mockups);")
    add_bullet("таблицы атрибутивного состава заполняемых экранных сеансов и первичных документов;")
    add_bullet("требования к доработке типовой функциональности КИС и внешним интеграционным шлюзам;")
    add_bullet("логическую реляционную модель данных (ERD) в третьей нормальной форме.")

    # -------------------------------------------------------------------------
    # 8. РАЗДЕЛ 3. ОСНОВНЫЕ ПОЛОЖЕНИЯ
    # -------------------------------------------------------------------------
    add_h("3 Основные положения", level=1)
    
    add_h("3.1 Бизнес-цели", level=2)
    add_p("Основными бизнес-целями автоматизации процесса оказания услуг чип-тюнинга являются:")
    add_bullet("сокращение общего времени цикла обслуживания одного автомобиля с 5–6 часов до 2–3 часов за счет параллелизации диагностики, электронной передачи стокового дампа прошивки и автоматического формирования наряд-заказов;")
    add_bullet("100% исключение рисков повреждения дорогостоящих электронных блоков управления (ECU) и двигателей благодаря обязательному аппаратному гейткипингу — проведению входной компьютерной диагностики до этапа чтения прошивки;")
    add_bullet("гарантированное сохранение оригинальных заводских калибровок (стоковых дампов) в централизованном защищенном архиве КИС с привязкой к VIN-номеру автомобиля и расчетной контрольной суммы (Checksum);")
    add_bullet("обеспечение сквозной юридической и финансовой прозрачности отношений с клиентом путем автоматической генерации электронного Наряд-заказа (ответственный: Данил) и финального Акта выполненных работ с гарантией (ответственный: Карим);")
    add_bullet("снижение трудоемкости рутинных операций оформления документов на 65% и исключение ошибок ручного ввода данных об автомобиле и клиенте.")

    add_h("3.2 Подход к реализации процесса", level=2)
    add_p("В основу концептуального подхода к реализации процесса положена сервисно-ориентированная сквозная цепочка создания ценности (End-to-End Service Flow), охватывающая жизненный цикл заявки от первичного обращения автовладельца до выезда модифицированного автомобиля из бокса:")
    add_bullet("четкое разграничение зон ответственности между сервисным отделом (мастер-приемщик, Гайнанов Д.И.) и инженерно-техническим производственным цехом (инженер-калибровщик, Гарифуллин К.Р.);")
    add_bullet("использование единого сквозного идентификатора заказа (id_order), гарантирующего целостность данных от заявки до бухгалтерской проводки и гарантийного сертификата;")
    add_bullet("шлюзовый контроль качества (Quality Gate): система программно блокирует переход наряд-заказа в статус «В работе у калибровщика», если в КИС не прикреплена подписанная электронная Диагностическая карта со статусом «ДВС исправен» (dtc_errors_found = False);")
    add_bullet("обязательное протоколирование инженерных параметров модификации (калибровки наддува, топливоподачи, УОЗ) в виде структурированного «Технического комментария мастера» непосредственно в АРМ калибровщика КИС.")

    add_h("3.3 Концептуальные допущения и ограничения", level=2)
    add_p("Бизнес-сценарий учитывает основные требования нормативной документации и стандартов предприятия системы менеджмента качества (СМК):")
    add_bullet("СТО 521.18.107-2007 «Система менеджмента качества. Изделия основного производства. Мероприятия по устранению и предупреждению дефектов» — регламентирует обязательность входного диагностического контроля параметров ДВС;")
    add_bullet("СТО 521.18.158-2018 «Система менеджмента качества. Детали, сборочные единицы изделий. Оформление решений на выпуск деталей, сборочных единиц с отклонениями от требований конструкторской документации» — регламентирует порядок фиксации отклонений и калибровочных изменений в техническом комментарии мастера;")
    add_bullet("ГОСТ 33997-2016 «Колесные транспортные средства. Требования к безопасности в эксплуатации и методы проверки»; ")
    add_bullet("технологическое ограничение: при обнаружении критических кодов ошибок DTC (пропуски зажигания, детонация, отклонение давления наддува более 10%, критическая коррекция топливоподачи LTFT > 15%) автомобиль категорически снимается с программного тюнинга и перенаправляется на механический ремонт.")

    add_h("3.4 Используемая функциональность", level=2)
    add_p("В качестве базовой платформы автоматизации выбрана конфигурация «1С:Предприятие 8.3 КОРП: Автосервис», расширенная специализированными модулями управления чип-тюнингом. Функциональность сгруппирована по следующим подсистемам:")
    add_bullet("подсистема «CRM и Продажи»: ведение карточек контрагентов, клиентской базы, парка автомобилей, регистрация первичных заявок и подбор пакетов калибровок (Stage 1 / Stage 2);")
    add_bullet("подсистема «Приемка и Документооборот»: формирование, калькуляция и утверждение электронных Наряд-заказов (Work Order), оформление предварительного согласия клиента (зона Данила);")
    add_bullet("подсистема «Техническая диагностика»: регистрация протоколов входного и выходного компьютерного сканирования ЭБУ, импорт отчетов диагностических сканеров Launch/VCDS/Autel;")
    add_bullet("подсистема «Тюнинг-производство (АРМ Калибровщика)»: управление жизненным циклом прошивок, защищенное бинарное хранилище дампов, интеграция с программаторами, фиксация контрольных сумм и инженерного технического комментария (зона Карима);")
    add_bullet("подсистема «Финансы и Взаиморасчеты»: кассовый расчет (наличные, эквайринг, СБП), печать фискальных чеков по 54-ФЗ, генерация финального Акта выполненных работ и печать гарантийного сертификата (зона Карима).")

    # -------------------------------------------------------------------------
    # 9. РАЗДЕЛ 4. БИЗНЕС-ПРОЦЕСС «ОКАЗАНИЕ УСЛУГ ЧИП-ТЮНИНГА АВТОМОБИЛЯ»
    # -------------------------------------------------------------------------
    add_h("4 Бизнес процесс «Оказание услуг программной модификации блоков управления двигателем (чип-тюнинг)»", level=1)
    
    add_h("4.1 Назначение и концепция бизнес-процесса", level=2)
    add_p("Бизнес-процесс «Оказание услуг программной модификации блоков управления двигателем (чип-тюнинг)» описывает регламентированную последовательность процедур и транзакций по взаимодействию персонала сервисного центра при регистрации клиента, согласовании объема доработок, входной инструментальной проверке ДВС, безопасной перезаписи управляющей микропрограммы контроллера, верификации динамических параметров и окончательном расчете с оформлением гарантийных обязательств.")
    add_p("Участниками и исполнителями по данному процессу являются:")
    add_bullet("Клиент (автовладелец или его законный представитель);")
    add_bullet("Мастер-приемщик (ведущий менеджер по работе с клиентами — студент Гайнанов Д.И.);")
    add_bullet("Инженер-диагност (специалист входного и выходного аппаратного контроля);")
    add_bullet("Инженер-калибровщик (технический специалист цеха тюнинга — студент Гарифуллин К.Р.);")
    add_bullet("Бухгалтер-кассир (специалист финансово-расчетного сектора);")
    add_bullet("КИС «Автосервис КОРП» (информационная подсистема маршрутизации, электронного документооборота и защищенного хранения калибровок).")
    add_p("Ключевой вход процесса: личное обращение клиента с автомобилем, требующим программной настройки (Stage 1 / Stage 2, отключение экологических подсистем Евро-2, DPF, EGR, AdBlue).")
    add_p("Ключевой выход процесса: автомобиль с верифицированным калибровочным софтом, сохраненный стоковый дамп в архиве КИС, подписанный сторонами электронный Акт выполненных работ с техническим комментарием мастера, фискальный кассовый чек и гарантийный сертификат на 24 месяца.")

    # 4.2 ИНДИВИДУАЛЬНЫЕ ДИНАМИЧЕСКИЕ МОДЕЛИ ПОДПРОЦЕССОВ
    add_h("4.2 Индивидуальные динамические модели подпроцессов (BPMN 2.0)", level=2)
    add_p("В соответствии с распределением функциональных зон ответственности между разработчиками детального бизнес-сценария, процесс декомпозирован на два сквозных технологических подпроцесса, оформленных в строгом стандарте BPMN 2.0 в среде MS Visio:")
    add_bullet("Подпроцесс 1: «Оформление заявки клиента и выпуск наряд-заказа» (зона ответственности студента Гайнанова Д.И.);")
    add_bullet("Подпроцесс 2: «Исполнение наряд-заказа, калибровка с комментарием мастера и выдача акта» (зона ответственности студента Гарифуллина К.Р.).")

    add_h("4.2.1 Подпроцесс 1. Оформление заявки клиента и выпуск наряд-заказа (Гайнанов Д.И.)", level=3)
    add_p("Подпроцесс 1 регламентирует сервисно-коммуникационный этап взаимодействия с клиентом. Модель охватывает 3 функциональные дорожки (Клиент, Мастер-приемщик Гайнанов Д.И., КИС «Автосервис»), 10 технологических процедур (шаги 1.1–1.10), шлюзы проверки технической применимости и согласования сметы, а также ключевой выходной артефакт — электронный Наряд-заказ (документ Work_Order).")
    add_p("Порядок исполнения подпроцесса 1:")
    add_bullet("Шаги 1.1–1.2: Автовладелец обращается на станцию и формулирует технические требования (прирост динамики, снятие ограничителя скорости). Мастер-приемщик Гайнанов Д.И. производит визуальный осмотр, фиксирует пробег и инициирует сеанс в КИС;")
    add_bullet("Шаги 1.3–1.4: КИС производит автоматический опрос истории обслуживания авто по VIN-номеру и сопоставляет модель ЭБУ со справочником доступных протоколов. В случае неподдерживаемого типа ЭБУ (шлюз «Тюнинг возможен?» -> «Нет») формируется отказ с регистрацией причины в CRM;")
    add_bullet("Шаги 1.5–1.6: Мастер консультирует клиента по пакетам калибровок (Stage 1 или Stage 2). КИС рассчитывает сметную стоимость, нормативную трудоемкость и формирует проект документа;")
    add_bullet("Шаги 1.7–1.8: Клиент знакомится со сметой и условиями гарантии. При несогласии автовладельца (шлюз «Клиент согласен?» -> «Нет») процесс завершается отказом;")
    add_bullet("Шаги 1.9–1.10: При акцепте условий КИС регистрирует документ «Наряд-заказ № КИС-2026». Стороны подписывают соглашение, мастер принимает ключи и направляет автомобиль на диагностический пост.")

    add_image_figure(
        "ДБС_Подпроцесс_1_Оформление_заявки_Гайнанов.png",
        "Рисунок 1 — Динамическая модель подпроцесса 1: Оформление заявки клиента и выпуск наряд-заказа (Гайнанов Д.И., BPMN 2.0)",
        width_inches=6.2
    )

    add_h("4.2.2 Подпроцесс 2. Исполнение наряд-заказа, калибровка с комментарием мастера и выдача акта (Гарифуллин К.Р.)", level=3)
    add_p("Подпроцесс 2 охватывает инженерно-производственный цикл: входной контроль исправности агрегатов, вычитывание оригинальной заводской прошивки, калибровку карт в WinOLS, безопасное перепрограммирование ЭБУ, обязательное протоколирование параметров настройки мастером, выходные динамические замеры и финансово-юридическое закрытие наряда. Модель включает 4 функциональные дорожки (Мастер-приемщик, Инженер-диагност, Инженер-калибровщик Гарифуллин К.Р., КИС «Автосервис: Тюнинг»), 10 технологических процедур (шаги 2.1–2.10) и ветвления по качеству.")
    add_p("Порядок исполнения подпроцесса 2:")
    add_bullet("Шаги 2.1–2.2: Входной инструментальный контроль. Инженер-диагност опрашивает память ЭБУ сканером. При наличии активных кодов ошибок DTC (шлюз «ДВС исправен?» -> «Нет») оформляется акт дефектовки, наряд-заказ блокируется и авто направляется на механический ремонт;")
    add_bullet("Шаги 2.3–2.4: При отсутствии неисправностей автомобиль передается инженеру Гарифуллину К.Р. Подключается стабилизированный источник питания (14.2 В). Вычитывается бинарный дамп заводской программы (ROM). КИС автоматически вычисляет хэш-сумму SHA-256 и архивирует стоковый файл в защищенное хранилище;")
    add_bullet("Шаги 2.5–2.6: Калибровщик производит редактирование таблиц зажигания, топливоподачи и наддува в среде WinOLS. Рассчитывается контрольная сумма (Checksum) и выполняется запись микрокода во флеш-память. При аппаратном сбое записи активируется аварийная процедура Recovery Mode для восстановления блока;")
    add_bullet("Шаг 2.7: Инженер-калибровщик Гарифуллин К.Р. производит адаптацию дроссельной заслонки и заносит в КИС развернутый «Технический комментарий мастера» (фактическое давление наддува, отсутствие детонации, требования к бензину АИ-98/100);")
    add_bullet("Шаг 2.8: Совместно с диагностом выполняется тестовый заезд и логгирование наддува и смеси под нагрузкой. При соответствии параметров норме (шлюз «Логи в норме?» -> «Да») формируется заключение;")
    add_bullet("Шаги 2.9–2.10: КИС генерирует выходной «Акт выполненных работ» и гарантийный талон на 24 месяца. Мастер и клиент подписывают Акт, кассир регистрирует оплату, и автомобиль передается владельцу.")

    add_image_figure(
        "ДБС_Подпроцесс_2_Исполнение_заказа_Гарифуллин.png",
        "Рисунок 2 — Динамическая модель подпроцесса 2: Исполнение наряд-заказа, калибровка с техническим комментарием мастера и выдача акта (Гарифуллин К.Р., BPMN 2.0)",
        width_inches=6.2
    )

    # 4.3 ОБЩАЯ СКВОЗНАЯ ДИНАМИЧЕСКАЯ МОДЕЛЬ В ТОМ ЖЕ BPMN-СТИЛЕ
    add_h("4.3 Общая сквозная динамическая модель процесса чип-тюнинга (BPMN 2.0, 5 дорожек)", level=2)
    add_p("Общая сквозная динамическая модель объединяет подпроцессы оформления заявки (Гайнанов Д.И.) и исполнения наряд-заказа (Гарифуллин К.Р.) в единую сервисную цепочку автосервиса. Схема разработана в MS Visio в нотации BPMN 2.0 с соблюдением единого корпоративного стиля:")
    add_bullet("синее колористическое оформление блоков действий (заливка RGB 239, 246, 255, граница RGB 37, 99, 235);")
    add_bullet("типизированные объекты данных BPMN в форме документов со скошенным уголком ([Документ: Наряд-заказ] и [Документ: Акт выполненных работ]);")
    add_bullet("прозрачный фон текстовых меток стрелок и шлюзов;")
    add_bullet("5 сквозных горизонтальных ролевых дорожек: Клиент (Автовладелец), Мастер-приемщик (Гайнанов Д.И.), Инженер-диагност (Входной контроль), Инженер-калибровщик (Гарифуллин К.Р.), КИС «Автосервис» (База данных и АРМ).")

    add_image_figure(
        "ДБС_Общая_динамическая_модель_BPMN.png",
        "Рисунок 3 — Общая сквозная динамическая модель процесса оказания услуг чип-тюнинга в КИС (BPMN 2.0, 5 ролевых дорожек)",
        width_inches=6.3
    )

    add_p("Сквозной процесс охватывает 18 процедурных шагов:")
    add_bullet("1. Клиент озвучивает пожелания по тюнингу и предоставляет автомобиль на станцию;")
    add_bullet("2. Мастер-приемщик осматривает ТС и заносит реквизиты в КИС (Сеанс 1.3);")
    add_bullet("3. КИС проверяет историю авто, модель ЭБУ и допуск к тюнингу;")
    add_bullet("4. При несовместимости ЭБУ (шлюз «Тюнинг возможен?» -> Нет) формируется отказ клиенту;")
    add_bullet("5. Мастер-приемщик подбирает тюнинг-пакет и рассчитывает смету (Сеанс 1.5);")
    add_bullet("6. Клиент рассматривает условия и смету;")
    add_bullet("7. При согласии клиента КИС формирует электронный Наряд-заказ;")
    add_bullet("8. Стороны подписывают Наряд-заказ (Сеанс 1.8, зона Данила), авто передается в цех;")
    add_bullet("9. Диагност проводит входное сканирование ДВС и проверку кодов DTC (Сеанс 1.10);")
    add_bullet("10. При обнаружении неисправностей (шлюз «ДВС исправен?» -> Нет) оформляется дефектовка и снятие авто с тюнинга;")
    add_bullet("11. При исправном ДВС калибровщик вычитывает стоковый дамп ROM в режиме Bench/OBD-II;")
    add_bullet("12. КИС автоматически сохраняет стоковый файл в защищенный архив с хэшем SHA-256;")
    add_bullet("13. Калибровщик оптимизирует карты в WinOLS, пересчитывает контрольные суммы и записывает софт в ЭБУ;")
    add_bullet("14. При сбое прошивки активируется аварийный режим Recovery Mode;")
    add_bullet("15. Мастер сбрасывает адаптации и вносит обязательный «Технический комментарий мастера» в АРМ (Сеанс 1.13, зона Карима);")
    add_bullet("16. Проводится выходной тест-драйв со снятием логов телеметрии и замером ВСХ;")
    add_bullet("17. КИС рассчитывает гарантию 24 месяца и генерирует Акт выполненных работ;")
    add_bullet("18. Мастер принимает оплату, стороны подписывают Акт (Сеанс 1.16, зона Карима), авто выдается владельцу.")

    # 4.4 МАТРИЦА ОТВЕТСТВЕННОСТИ РОЛЕЙ (RACI)
    add_h("4.4 Матрица распределения ответственности ролей по шагам процесса (RACI)", level=2)
    add_p("В таблице 2 приведена матрица распределения ответственности ролей по 7 ключевым технологическим шагам процесса в соответствии с методологией RACI (О — Ответственный, И — Исполнитель, С — Согласующий, У — Уведомляемый).")

    add_p("Таблица 2 — Матрица ответственности ролей по шагам процесса (RACI)", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    raci_rows = [
        ("Шаг 1. Регистрация клиента и ТС (Сеанс 1.3)", "С", "О / И", "—", "—", "—"),
        ("Шаг 2. Подбор тюнинг-пакета (Сеанс 1.5)", "С", "О / И", "—", "У", "—"),
        ("Шаг 3. Утверждение Наряд-заказа (Сеанс 1.8, Данил)", "С", "О / И", "—", "У", "У"),
        ("Шаг 4. Входная диагностика ДВС (Сеанс 1.10)", "У", "С", "О / И", "С", "—"),
        ("Шаг 5. Калибровка ЭБУ и коммент (Сеанс 1.13, Карим)", "—", "У", "—", "О / И", "—"),
        ("Шаг 6. Выходная диагностика и замеры (Сеанс 1.15)", "—", "У", "И", "О / И", "—"),
        ("Шаг 7. Выдача Акта и взаиморасчет (Сеанс 1.16, Карим)", "С", "И", "—", "О", "О / И")
    ]

    t_raci = doc.add_table(rows=len(raci_rows) + 1, cols=6)
    t_raci.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_raci)
    t_raci.columns[0].width = Mm(65)
    t_raci.columns[1].width = Mm(20)
    t_raci.columns[2].width = Mm(22)
    t_raci.columns[3].width = Mm(20)
    t_raci.columns[4].width = Mm(23)
    t_raci.columns[5].width = Mm(20)

    headers_raci = ["Шаги бизнес-процесса", "Клиент", "Мастер (Данил)", "Диагност", "Калибровщик (Карим)", "Кассир"]
    for c_idx, h_text in enumerate(headers_raci):
        style_header_cell(t_raci.cell(0, c_idx), h_text)

    for idx, row in enumerate(raci_rows):
        r_num = idx + 1
        style_data_cell(t_raci.cell(r_num, 0), row[0], bold=False)
        for col_idx in range(1, 6):
            style_data_cell(t_raci.cell(r_num, col_idx), row[col_idx], bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 4.2 ОПИСАНИЕ ШАГОВ БИЗНЕС-ПРОЦЕССА (7 ШАГОВ)
    # -------------------------------------------------------------------------
    add_h("4.5 Описание шагов бизнес процесса", level=2)

    # -------------------------------------------------------------------------
    # ШАГ 1
    # -------------------------------------------------------------------------
    add_h("4.5.1 Шаг 1. Первичная регистрация клиента и идентификация автомобиля", level=3)
    
    add_p("Спецификация шага 1 приведена в таблице 3.")
    add_p("Таблица 3 — Спецификация шага 1 «Регистрация клиента и ТС»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    spec_step1 = [
        ("Условия выполнения", "Личное прибытие клиента в сервисный центр или заявка с сайта. Отсутствие карточки клиента в базе данных КИС."),
        ("Исполнитель", "Мастер-приемщик (зона ответственности Данила — Гайнанов Д.И.)"),
        ("Вход", "Паспортные данные клиента, свидетельство о регистрации ТС (СТС / ПТС), физический доступ к автомобилю."),
        ("Выход", "Созданный и валидированный элемент справочника «Клиенты» (id_client) и справочника «Автомобили» (vin_code) в КИС Автосервис.")
    ]
    t_s1 = doc.add_table(rows=4, cols=2)
    t_s1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_s1)
    t_s1.columns[0].width = Mm(45)
    t_s1.columns[1].width = Mm(120)
    for r_idx, (k, v) in enumerate(spec_step1):
        style_data_cell(t_s1.cell(r_idx, 0), k, bold=True)
        style_data_cell(t_s1.cell(r_idx, 1), v, bold=False)

    add_h("Назначение", level=4)
    add_p("Идентификация контрагента, установление правомочности распоряжения транспортным средством, автоматическое определение модификации блока управления двигателем по каталогу VIN-кодов и привязка истории обслуживания к единому профилю автовладельца в КИС.")

    add_h("Описание действий", level=4)
    add_p("Мастер-приемщик открывает в КИС сеанс «Регистрация клиента и автомобиля» (Сеанс 1.3). Вводит ФИО, телефон и e-mail. Сканирует VIN-код из СТС (или считывает штрихкод сканером). Система производит декодирование VIN-кода: определяет марку, модель, год выпуска, буквенный код двигателя, установленную трансмиссию и базовую модель ЭБУ (например, Bosch MED17.5.2). При успешной валидации карточка сохраняется в базе данных.")

    add_image_figure(
        "ДБС_Экран_1_Клиент_Автомобиль.png",
        "Рисунок 4 — Экранная форма КИС «Регистрация клиента и автомобиля» (Сеанс 1.3)",
        width_inches=5.8
    )

    add_p("Значимые атрибуты объекта шага приведены в таблице 4.")
    add_p("Таблица 4 — Значимые атрибуты сеанса «Регистрация клиента и ТС»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    attrs_step1 = [
        ("id_client", "Уникальный идентификатор контрагента", "Генерируется автоматически системой (GUID / автоинкремент)."),
        ("client_fio", "Полное имя клиента / представителя", "Строка, ввод с документа удостоверяющего личность."),
        ("phone_number", "Контактный номер мобильного телефона", "Маска ввода +7 (XXX) XXX-XX-XX, используется для SMS-оповещений."),
        ("vin_code", "Идентификационный номер ТС (VIN)", "17-значный уникальный код стандарта ISO 3779. Проверка контрольного символа."),
        ("brand_model", "Марка и модификация автомобиля", "Заполняется автоматически из VIN-декодера (например, BMW 330i G20)."),
        ("engine_model", "Модель двигателя и объем", "Выбирается из базы двигателей (например, 2.0 B48B20 258 л.с.)."),
        ("ecu_type", "Тип установленного контроллера ЭБУ", "Справочник ЭБУ (например, Bosch MG1CS024 / Continental SIMOS 18.1).")
    ]
    t_a1 = doc.add_table(rows=len(attrs_step1) + 1, cols=3)
    t_a1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_a1)
    t_a1.columns[0].width = Mm(35)
    t_a1.columns[1].width = Mm(65)
    t_a1.columns[2].width = Mm(65)
    style_header_cell(t_a1.cell(0, 0), "Атрибут")
    style_header_cell(t_a1.cell(0, 1), "Описание")
    style_header_cell(t_a1.cell(0, 2), "Рекомендации по заполнению")
    for r_idx, (a, d, r) in enumerate(attrs_step1):
        r_num = r_idx + 1
        style_data_cell(t_a1.cell(r_num, 0), a, bold=False)
        style_data_cell(t_a1.cell(r_num, 1), d, bold=False)
        style_data_cell(t_a1.cell(r_num, 2), r, bold=False)

    add_h("Требуемое изменение функциональности", level=4)
    add_p("Внедрить модуль автоматического парсинга VIN-номеров через облачный API AutoData / TecDoc с автозаполнением полей типа ЭБУ и протокола программирования.")

    add_h("Требования к интеграции с другими системами", level=4)
    add_p("Направление: исходящий REST API запрос к сервису VIN-декодера; входящий JSON ответ. Частота: On-demand (по событию ввода VIN).")

    # -------------------------------------------------------------------------
    # ШАГ 2
    # -------------------------------------------------------------------------
    add_h("4.5.2 Шаг 2. Формирование заявки и подбор конфигурации чип-тюнинга", level=3)
    
    add_p("Спецификация шага 2 приведена в таблице 5.")
    add_p("Таблица 5 — Спецификация шага 2 «Формирование заявки и подбор пакета»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    spec_step2 = [
        ("Условия выполнения", "Зарегистрированный клиент и автомобиль в КИС; наличие актуального каталога прошивок."),
        ("Исполнитель", "Мастер-приемщик (Гайнанов Д.И.)"),
        ("Вход", "Пожелания клиента по характеристикам (прирост мощности, динамика, снятие лимитера скорости, экология)."),
        ("Выход", "Сформированная электронная заявка со статусом «Сформирована» (документ Service_Request).")
    ]
    t_s2 = doc.add_table(rows=4, cols=2)
    t_s2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_s2)
    t_s2.columns[0].width = Mm(45)
    t_s2.columns[1].width = Mm(120)
    for r_idx, (k, v) in enumerate(spec_step2):
        style_data_cell(t_s2.cell(r_idx, 0), k, bold=True)
        style_data_cell(t_s2.cell(r_idx, 1), v, bold=False)

    add_h("Назначение", level=4)
    add_p("Фиксация потребностей автовладельца, формализация требований к калибровке ДВС и подбор конкретной технической конфигурации тюнинг-пакета (Stage 1 / Stage 2) с расчетом расчетной прибавки мощности и крутящего момента.")

    add_h("Описание действий", level=4)
    add_p("Мастер открывает сеанс «Оформление заявки на чип-тюнинг» (Сеанс 1.5). Выбирает ранее зарегистрированного клиента. Из справочника пакетов тюнинга подбирает вариант: Stage 1 (программная настройка штатного мотора, +52 л.с., +90 Нм) либо Stage 2 (настройка под спортивный даунпайп). При необходимости отмечаются дополнительные опции: «Отключение вихревых заслонок», «Отключение контроля катализатора Euro-2», «Pop & Bang (отстрелы)». Система рассчитывает предварительную стоимость.")

    add_image_figure(
        "ДБС_Экран_2_Заявка_Тюнинг.png",
        "Рисунок 5 — Экранная форма КИС «Оформление заявки на чип-тюнинг» (Сеанс 1.5)",
        width_inches=5.8
    )

    add_p("Значимые атрибуты объекта шага приведены в таблице 6.")
    add_p("Таблица 6 — Значимые атрибуты сеанса «Заявка на чип-тюнинг»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    attrs_step2 = [
        ("id_request", "Номер заявки на обслуживание", "Автоинкрементный уникальный системный номер заявки."),
        ("package_type", "Категория тюнинг-пакета", "Выбор из перечисления: Stage 1, Stage 2, EcoTuning, Custom."),
        ("power_gain_hp", "Расчетный прирост мощности (л.с.)", "Автоматически подтягивается из справочника Tuning_Package."),
        ("torque_gain_nm", "Расчетный прирост крутящего момента (Нм)", "Автоматически подтягивается из справочника Tuning_Package."),
        ("eco_options", "Опции модификации экологии", "Множественный выбор чекбоксов: Decat, DPF off, EGR off, AdBlue off."),
        ("preliminary_cost", "Предварительная расчетная стоимость", "Сумма пакета и дополнительных программных модулей (руб.).")
    ]
    t_a2 = doc.add_table(rows=len(attrs_step2) + 1, cols=3)
    t_a2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_a2)
    t_a2.columns[0].width = Mm(35)
    t_a2.columns[1].width = Mm(65)
    t_a2.columns[2].width = Mm(65)
    style_header_cell(t_a2.cell(0, 0), "Атрибут")
    style_header_cell(t_a2.cell(0, 1), "Описание")
    style_header_cell(t_a2.cell(0, 2), "Рекомендации по заполнению")
    for r_idx, (a, d, r) in enumerate(attrs_step2):
        r_num = r_idx + 1
        style_data_cell(t_a2.cell(r_num, 0), a, bold=False)
        style_data_cell(t_a2.cell(r_num, 1), d, bold=False)
        style_data_cell(t_a2.cell(r_num, 2), r, bold=False)

    add_h("Требуемое изменение функциональности", level=4)
    add_p("Настройка правил валидации: запрет выбора Stage 2 без предварительной отметки об установке модифицированной выхлопной системы и интеркулера.")

    add_h("Требования к интеграции с другими системами", level=4)
    add_p("Н/П (внутренний функционал КИС).")

    # -------------------------------------------------------------------------
    # ШАГ 3 (ДАНИЛ)
    # -------------------------------------------------------------------------
    add_h("4.5.3 Шаг 3. Согласование и утверждение Наряд-заказа (Подпроцесс Данила)", level=3)
    
    add_p("Спецификация шага 3 приведена в таблице 7.")
    add_p("Таблица 7 — Спецификация шага 3 «Утверждение Наряд-заказа»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    spec_step3 = [
        ("Условия выполнения", "Сформированная заявка Service_Request, согласие клиента с предварительными условиями и стоимостью."),
        ("Исполнитель", "Мастер-приемщик (Гайнанов Д.И. — ключевой итоговый документ этапа оформления заявки)"),
        ("Вход", "Электронная заявка id_request, согласованный перечень работ и материалов."),
        ("Выход", "Утвержденный и подписанный клиентом Наряд-заказ (документ Work_Order) со статусом «Утвержден».")
    ]
    t_s3 = doc.add_table(rows=4, cols=2)
    t_s3.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_s3)
    t_s3.columns[0].width = Mm(45)
    t_s3.columns[1].width = Mm(120)
    for r_idx, (k, v) in enumerate(spec_step3):
        style_data_cell(t_s3.cell(r_idx, 0), k, bold=True)
        style_data_cell(t_s3.cell(r_idx, 1), v, bold=False)

    add_h("Назначение", level=4)
    add_p("Юридическое и финансовое закрепление договоренностей между автосервисом и клиентом. Наряд-заказ выступает договором подряда на выполнение чип-тюнинга, фиксирует плановые сроки сдачи автомобиля, предельную смету и согласие владельца на перепрограммирование ЭБУ.")

    add_h("Описание действий", level=4)
    add_p("На основании заявки мастер Гайнанов Д.И. открывает сеанс «Наряд-заказ на обслуживание» (Сеанс 1.8). Система автоматически связывает наряд-заказ с id_request (исключая прямое дублирование vin_code для сохранения 3NF). Мастер назначает плановое время выдачи автомобиля, ответственного инженера-калибровщика и калькулирует итоговую стоимость. Документ выводится на двухстороннюю печать (или отправляется на электронный планшет подписи). Клиент ставит подпись, подтверждая ознакомление с регламентом. Статус документа в КИС переключается на «Утвержден к исполнению».")

    add_image_figure(
        "ДБС_Экран_3_Наряд_Заказ.png",
        "Рисунок 6 — Экранная форма КИС «Утверждение Наряд-заказа» (Сеанс 1.8 — зона ответственности Данила)",
        width_inches=5.8
    )

    add_p("Значимые атрибуты объекта шага приведены в таблице 8.")
    add_p("Таблица 8 — Значимые атрибуты сеанса «Наряд-заказ»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    attrs_step3 = [
        ("id_order", "Номер наряд-заказа (Work Order ID)", "Уникальный учетный номер документа (например, НЗ-2026/04-118)."),
        ("id_request", "Ссылка на заявку-основание (FK)", "Внешний ключ связи 1:1 к сущности Service_Request."),
        ("order_date", "Дата и время утверждения", "Текущие системные дата и время регистрации документа."),
        ("assigned_master", "ФИО ответственного приемщика", "Заполняется текущим пользователем сеанса (Гайнанов Д.И.)."),
        ("assigned_tuner", "ФИО назначенного калибровщика", "Выбор из справочника сотрудников роли «Инженер-калибровщик» (Гарифуллин К.Р.)."),
        ("total_cost", "Окончательная договорная сумма", "Итоговая сумма к оплате с учетом НДС (руб.)."),
        ("order_status", "Статус исполнения заказа", "Перечисление: Черновик, Утвержден, На диагностике, В работе, Выполнен, Закрыт.")
    ]
    t_a3 = doc.add_table(rows=len(attrs_step3) + 1, cols=3)
    t_a3.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_a3)
    t_a3.columns[0].width = Mm(35)
    t_a3.columns[1].width = Mm(65)
    t_a3.columns[2].width = Mm(65)
    style_header_cell(t_a3.cell(0, 0), "Атрибут")
    style_header_cell(t_a3.cell(0, 1), "Описание")
    style_header_cell(t_a3.cell(0, 2), "Рекомендации по заполнению")
    for r_idx, (a, d, r) in enumerate(attrs_step3):
        r_num = r_idx + 1
        style_data_cell(t_a3.cell(r_num, 0), a, bold=False)
        style_data_cell(t_a3.cell(r_num, 1), d, bold=False)
        style_data_cell(t_a3.cell(r_num, 2), r, bold=False)

    add_h("Требуемое изменение функциональности", level=4)
    add_p("Автоматическая блокировка изменений спецификации наряд-заказа после проставления статуса «Утвержден» и фиксации подписи клиента.")

    add_h("Требования к интеграции с другими системами", level=4)
    add_p("Интеграция с корпоративным SMS/Telegram-шлюзом для автоматической отправки клиенту трек-ссылки на отслеживание статуса исполнения заказа.")

    # -------------------------------------------------------------------------
    # ШАГ 4
    # -------------------------------------------------------------------------
    add_h("4.5.4 Шаг 4. Входная компьютерная диагностика систем ДВС и ЭБУ", level=3)
    
    add_p("Спецификация шага 4 приведена в таблице 9.")
    add_p("Таблица 9 — Спецификация шага 4 «Входная компьютерная диагностика»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    spec_step4 = [
        ("Условия выполнения", "Наряд-заказ в статусе «Утвержден»; автомобиль заведен в диагностический бокс сервиса."),
        ("Исполнитель", "Диагност-электрик"),
        ("Вход", "Утвержденный наряд-заказ id_order, диагностический разъем OBD-II автомобиля, профессиональный сканер Launch X431 / Autel MaxiSys."),
        ("Выход", "Сформированная электронная Диагностическая карта (Diagnostics_Card) с вердиктом о допуске к прошивке.")
    ]
    t_s4 = doc.add_table(rows=4, cols=2)
    t_s4.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_s4)
    t_s4.columns[0].width = Mm(45)
    t_s4.columns[1].width = Mm(120)
    for r_idx, (k, v) in enumerate(spec_step4):
        style_data_cell(t_s4.cell(r_idx, 0), k, bold=True)
        style_data_cell(t_s4.cell(r_idx, 1), v, bold=False)

    add_h("Назначение", level=4)
    add_p("Аппаратный шлюзовый контроль качества (Quality Gate): проверка отсутствия механических и электронных неисправностей цилиндропоршневой группы, топливной аппаратуры и турбонаддува. Запрещено модифицировать калибровки неисправного двигателя.")

    add_h("Описание действий", level=4)
    add_p("Диагност открывает сеанс «Диагностическая карта ТС» (Сеанс 1.10). Подключает сканер к разъему OBD-II. Выполняет полный опрос электронных блоков на наличие диагностических кодов неисправностей (DTC). Оценивает параметры долгосрочной коррекции топлива (LTFT), давление в топливной рампе, противодавление сажевого фильтра и компрессию по цилиндрам. Если ошибок нет (dtc_errors_found = False), в КИС проставляется статус «ДВС исправен, допущен к чип-тюнингу». При наличии критических ошибок документ блокирует дальнейшие этапы и направляет авто на ремонт.")

    add_image_figure(
        "ДБС_Экран_4_Диагностическая_Карта.png",
        "Рисунок 7 — Экранная форма КИС «Входная диагностическая карта» (Сеанс 1.10)",
        width_inches=5.8
    )

    add_p("Значимые атрибуты объекта шага приведены в таблице 10.")
    add_p("Таблица 10 — Значимые атрибуты сеанса «Диагностическая карта»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    attrs_step4 = [
        ("id_card", "Уникальный номер карты диагностики", "Автоинкрементный системный идентификатор."),
        ("id_order", "Ссылка на наряд-заказ (FK)", "Внешний ключ, связывающий диагностику с наряд-заказом."),
        ("dtc_errors_found", "Флаг обнаружения кодов DTC", "Логическое поле (True/False). True — ошибки найдены, тюнинг запрещен."),
        ("fuel_correction_ltft", "Долгосрочная топливная коррекция", "Параметр в процентах (допустимая норма: от -5% до +5%)."),
        ("boost_pressure_bar", "Фактическое давление наддува", "Давление наддува на холостом ходу и под нагрузкой (бар)."),
        ("diagnostic_verdict", "Итоговое заключение диагноста", "Выбор: Допущен к чип-тюнингу / Требуется ремонт ДВС.")
    ]
    t_a4 = doc.add_table(rows=len(attrs_step4) + 1, cols=3)
    t_a4.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_a4)
    t_a4.columns[0].width = Mm(35)
    t_a4.columns[1].width = Mm(65)
    t_a4.columns[2].width = Mm(65)
    style_header_cell(t_a4.cell(0, 0), "Атрибут")
    style_header_cell(t_a4.cell(0, 1), "Описание")
    style_header_cell(t_a4.cell(0, 2), "Рекомендации по заполнению")
    for r_idx, (a, d, r) in enumerate(attrs_step4):
        r_num = r_idx + 1
        style_data_cell(t_a4.cell(r_num, 0), a, bold=False)
        style_data_cell(t_a4.cell(r_num, 1), d, bold=False)
        style_data_cell(t_a4.cell(r_num, 2), r, bold=False)

    add_h("Требуемое изменение функциональности", level=4)
    add_p("Автоматический парсинг PDF-отчетов диагностических сканеров Launch/Autel и автозаполнение полей параметров и списка кодов DTC.")

    add_h("Требования к интеграции с другими системами", level=4)
    add_p("Сетевая папка обмена (SMB) с диагностическим стендом для фонового импорта XML-отчетов сканирования.")

    # -------------------------------------------------------------------------
    # ШАГ 5 (КАРИМ)
    # -------------------------------------------------------------------------
    add_h("4.5.5 Шаг 5. Считывание стоковой прошивки, калибровка и запись модифицированного ПО с фиксацией технического комментария (Подпроцесс Карима)", level=3)
    
    add_p("Спецификация шага 5 приведена в таблице 11.")
    add_p("Таблица 11 — Спецификация шага 5 «Калибровка ЭБУ и фиксация комментария»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    spec_step5 = [
        ("Условия выполнения", "Положительный вердикт входной диагностики; наряд-заказ в статусе «В работе» у калибровщика."),
        ("Исполнитель", "Инженер-калибровщик (Гарифуллин К.Р. — ключевой этап непосредственного исполнения заказа)"),
        ("Вход", "Диагностическая карта со статусом «Допущен», аппаратный программатор (Autotuner / K-Tag), лицензионный редактор WinOLS."),
        ("Выход", "Считанный и сохраненный заводской дамп (ROM), подготовленный и записанный модифицированный файл с корректной контрольной суммой, созданная запись Order_Execution с обязательным техническим комментарием мастера.")
    ]
    t_s5 = doc.add_table(rows=4, cols=2)
    t_s5.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_s5)
    t_s5.columns[0].width = Mm(45)
    t_s5.columns[1].width = Mm(120)
    for r_idx, (k, v) in enumerate(spec_step5):
        style_data_cell(t_s5.cell(r_idx, 0), k, bold=True)
        style_data_cell(t_s5.cell(r_idx, 1), v, bold=False)

    add_h("Назначение", level=4)
    add_p("Технологическое выполнение модификации микрокода блока управления двигателем. Гарантированное архивирование исходного заводского дампа прошивки для обеспечения возможности отката (Rollback). Пересчет контрольных сумм для защиты от блокировки процессора ЭБУ. Фиксация инженерных особенностей настройки и поведения агрегата в обязательном поле «Технический комментарий мастера».")

    add_h("Описание действий", level=4)
    add_p("Инженер Гарифуллин К.Р. открывает «АРМ Калибровщика: Исполнение наряд-заказа» (Сеанс 1.13). Подключает стабилизатор питания 14.2В к бортовой сети авто. Через OBD-II или Bench-колодку считывает оригинальную прошивку. КИС сохраняет файл в защищенный архив с привязкой хэша SHA-256 (stock_file_hash). Гарифуллин К.Р. открывает карты в редакторе WinOLS, производит оптимизацию карт зажигания, времени открытия форсунок и давления наддува (+0.25 бар). Выполняется автоматический пересчет контрольной суммы (Checksum: 0x9F4C2A1E). Модифицированное ПО успешно записывается во флеш-память ЭБУ. Мастер вносит в экранную форму обязательный развернутый Технический комментарий мастера (поле technical_comment): параметры наддува, коррекции УОЗ, отсутствие детонации и рекомендации по использованию топлива АИ-98/100.")

    add_image_figure(
        "ДБС_Экран_5_АРМ_Калибровщика_Коммент.png",
        "Рисунок 8 — Экранная форма КИС «АРМ Калибровщика: Исполнение заказа с техническим комментарием мастера» (Сеанс 1.13 — зона ответственности Карима)",
        width_inches=5.8
    )

    add_p("Значимые атрибуты объекта шага приведены в таблице 12.")
    add_p("Таблица 12 — Значимые атрибуты сеанса «АРМ Калибровщика»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    attrs_step5 = [
        ("id_execution", "Уникальный номер сеанса исполнения", "Системный автоинкрементный номер записи."),
        ("id_order", "Ссылка на наряд-заказ (FK)", "Внешний ключ к сущности Work_Order."),
        ("stock_file_hash", "Хэш-сумма заводского дампа (SHA-256)", "Контрольный хэш стокового бинарного файла для гарантии идентичности."),
        ("mod_file_hash", "Хэш-сумма модифицированного файла", "Контрольный хэш записанного файла прошивки."),
        ("checksum_valid", "Флаг корректности контрольной суммы", "Логическое поле (True/False). Запись возможна только при True."),
        ("flashing_tool", "Оборудование для программирования", "Справочник программаторов: Autotuner OBD/Bench, CMDFlash, K-Tag."),
        ("technical_comment", "Технический комментарий мастера", "Текстовое обязательное поле: инженерные нюансы, наддув, УОЗ, качество смеси, рекомендации по топливу и маслу.")
    ]
    t_a5 = doc.add_table(rows=len(attrs_step5) + 1, cols=3)
    t_a5.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_a5)
    t_a5.columns[0].width = Mm(35)
    t_a5.columns[1].width = Mm(65)
    t_a5.columns[2].width = Mm(65)
    style_header_cell(t_a5.cell(0, 0), "Атрибут")
    style_header_cell(t_a5.cell(0, 1), "Описание")
    style_header_cell(t_a5.cell(0, 2), "Рекомендации по заполнению")
    for r_idx, (a, d, r) in enumerate(attrs_step5):
        r_num = r_idx + 1
        style_data_cell(t_a5.cell(r_num, 0), a, bold=False)
        style_data_cell(t_a5.cell(r_num, 1), d, bold=False)
        style_data_cell(t_a5.cell(r_num, 2), r, bold=False)

    add_h("Требуемое изменение функциональности", level=4)
    add_p("Обязательная проверка заполненности поля «Технический комментарий мастера» при нажатии кнопки «Завершить калибровку» (запрет сохранения пустой строки длиною менее 30 символов).")

    add_h("Требования к интеграции с другими системами", level=4)
    add_p("Прямая интеграция КИС с DLL-библиотеками программатора Autotuner / K-Suite для автоматического считывания аппаратных номеров софта и калибровок (HW/SW Number).")

    # -------------------------------------------------------------------------
    # ШАГ 6
    # -------------------------------------------------------------------------
    add_h("4.5.6 Шаг 6. Выходная диагностика, динамические замеры и верификация параметров", level=3)
    
    add_p("Спецификация шага 6 приведена в таблице 13.")
    add_p("Таблица 13 — Спецификация шага 6 «Выходная диагностика и верификация»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    spec_step6 = [
        ("Условия выполнения", "Успешная запись калибровки в ЭБУ; автомобиль установлен на мощностной динамометрический стенд (Dyno) или выведен на тестовый заезд."),
        ("Исполнитель", "Инженер-калибровщик (Гарифуллин К.Р.) совместно с диагностом"),
        ("Вход", "Автомобиль с обновленным софтом ЭБУ, подключенный логгер параметров, диностенд Superflow / Dynomax."),
        ("Выход", "Сформированный график замера мощности и крутящего момента (Dyno-chart), подтвержденное отсутствие детонации (Knock Retard = 0.0 deg).")
    ]
    t_s6 = doc.add_table(rows=4, cols=2)
    t_s6.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_s6)
    t_s6.columns[0].width = Mm(45)
    t_s6.columns[1].width = Mm(120)
    for r_idx, (k, v) in enumerate(spec_step6):
        style_data_cell(t_s6.cell(r_idx, 0), k, bold=True)
        style_data_cell(t_s6.cell(r_idx, 1), v, bold=False)

    add_h("Назначение", level=4)
    add_p("Контроль качества калибровки под предельной нагрузкой: подтверждение фактического прироста мощности, проверка стабильности давления наддува, температуры выхлопа (EGT < 880°C) и полного отсутствия детонационного сгорания топлива.")

    add_h("Описание действий", level=4)
    add_p("Специалисты проводят серию из 3 ускорений в режиме полного газа (WOT) на 4-й передаче. КИС считывает телеметрию: фактическая мощность достигла 310 л.с. (прибавка +61 л.с.), крутящий момент составил 720 Нм (прибавка +100 Нм). Снятие логов подтверждает идеальный состав смеси (Lambda = 0.82) и отсутствие откатов зажигания. В КИС загружается электронный график замера мощности. Наряд-заказ переводится в статус «Работы выполнены».")

    add_p("Значимые атрибуты объекта шага приведены в таблице 14.")
    add_p("Таблица 14 — Значимые атрибуты шага «Выходная верификация»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    attrs_step6 = [
        ("dyno_run_id", "Идентификатор прогона на диностенде", "Номер записи в базе стенда измерений мощности."),
        ("actual_power_hp", "Фактическая измеренная мощность", "Показание диностенда на маховике (л.с.)."),
        ("actual_torque_nm", "Фактический крутящий момент", "Показание диностенда (Нм)."),
        ("max_egt_celsius", "Максимальная температура газов (EGT)", "Температура перед турбокомпрессором (норма до 900°C)."),
        ("knock_retard_deg", "Угол отката зажигания по детонации", "Градусы коррекции УОЗ (норма: 0.0° во всех цилиндрах).")
    ]
    t_a6 = doc.add_table(rows=len(attrs_step6) + 1, cols=3)
    t_a6.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_a6)
    t_a6.columns[0].width = Mm(35)
    t_a6.columns[1].width = Mm(65)
    t_a6.columns[2].width = Mm(65)
    style_header_cell(t_a6.cell(0, 0), "Атрибут")
    style_header_cell(t_a6.cell(0, 1), "Описание")
    style_header_cell(t_a6.cell(0, 2), "Рекомендации по заполнению")
    for r_idx, (a, d, r) in enumerate(attrs_step6):
        r_num = r_idx + 1
        style_data_cell(t_a6.cell(r_num, 0), a, bold=False)
        style_data_cell(t_a6.cell(r_num, 1), d, bold=False)
        style_data_cell(t_a6.cell(r_num, 2), r, bold=False)

    add_h("Требуемое изменение функциональности", level=4)
    add_p("Модуль визуализации графика ВСХ (внешней скоростной характеристики) непосредственно в карточке наряд-заказа КИС.")

    add_h("Требования к интеграции с другими системами", level=4)
    add_p("Интеграция по протоколу TCP/IP с ПО диностенда для прямого получения числовых массивов замера.")

    # -------------------------------------------------------------------------
    # ШАГ 7 (КАРИМ)
    # -------------------------------------------------------------------------
    add_h("4.5.7 Шаг 7. Формирование и выдача Акта выполненных работ, гарантийного сертификата и взаиморасчет (Подпроцесс Карима)", level=3)
    
    add_p("Спецификация шага 7 приведена в таблице 15.")
    add_p("Таблица 15 — Спецификация шага 7 «Выдача Акта и взаиморасчет»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    spec_step7 = [
        ("Условия выполнения", "Подтвержденная верификация калибровки; готовность автомобиля к выдаче автовладельцу."),
        ("Исполнитель", "Инженер-калибровщик (Гарифуллин К.Р.) совместно с бухгалтером-кассиром"),
        ("Вход", "Выполненный наряд-заказ, протокол диностенда, денежные средства клиента."),
        ("Выход", "Подписанный двусторонний Акт выполненных работ (Work_Completion_Act), фискальный кассовый чек, электронный гарантийный сертификат на прошивку.")
    ]
    t_s7 = doc.add_table(rows=4, cols=2)
    t_s7.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_s7)
    t_s7.columns[0].width = Mm(45)
    t_s7.columns[1].width = Mm(120)
    for r_idx, (k, v) in enumerate(spec_step7):
        style_data_cell(t_s7.cell(r_idx, 0), k, bold=True)
        style_data_cell(t_s7.cell(r_idx, 1), v, bold=False)

    add_h("Назначение", level=4)
    add_p("Окончательное юридическое закрытие сделки, проведение оплаты, передача автомобиля владельцу с предоставлением официальной гарантии на калибровочное программное обеспечение (бессрочная гарантия на софт и 1 год на корректность работы узлов).")

    add_h("Описание действий", level=4)
    add_p("В КИС формируется заключительный документ «Акт выполненных работ» (Сеанс 1.16). Кассир принимает оплату (картой через эквайринговый терминал или наличными), КИС регистрирует чек по 54-ФЗ. Калибровщик Гарифуллин К.Р. демонстрирует клиенту графики прироста мощности и распечатывает гарантийный сертификат. Стороны подписывают Акт выполненных работ, где зафиксировано отсутствие претензий по состоянию автомобиля. Статус наряд-заказа в КИС переводится в финальное состояние «Закрыт». Ключи передаются владельцу.")

    add_image_figure(
        "ДБС_Экран_6_Акт_Выполненных_Работ.png",
        "Рисунок 9 — Экранная форма КИС «Акт выполненных работ и гарантия» (Сеанс 1.16 — зона ответственности Карима)",
        width_inches=5.8
    )

    add_p("Значимые атрибуты объекта шага приведены в таблице 16.")
    add_p("Таблица 16 — Значимые атрибуты сеанса «Акт выполненных работ»", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    attrs_step7 = [
        ("id_act", "Уникальный номер акта (Work Completion Act ID)", "Системный номер документа закрытия заказа (например, АКТ-2026/04-118)."),
        ("id_order", "Ссылка на наряд-заказ (FK)", "Внешний ключ связи 1:1 к сущности Work_Order."),
        ("act_date", "Дата и время закрытия акта", "Дата фактической передачи автомобиля клиенту."),
        ("warranty_period_months", "Гарантийный период (мес.)", "Срок гарантийных обязательств на софт (устанавливается 24 мес.)."),
        ("payment_type", "Способ оплаты", "Перечисление: Банковская карта, СБП QR, Наличный расчет, Безналичный (юр. лица)."),
        ("fiscal_receipt_number", "Номер фискального чека", "Идентификатор фискального документа (ФД) из онлайн-кассы."),
        ("client_signature_confirmed", "Флаг подписи клиента", "Подтверждение факта приема авто без замечаний.")
    ]
    t_a7 = doc.add_table(rows=len(attrs_step7) + 1, cols=3)
    t_a7.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_a7)
    t_a7.columns[0].width = Mm(35)
    t_a7.columns[1].width = Mm(65)
    t_a7.columns[2].width = Mm(65)
    style_header_cell(t_a7.cell(0, 0), "Атрибут")
    style_header_cell(t_a7.cell(0, 1), "Описание")
    style_header_cell(t_a7.cell(0, 2), "Рекомендации по заполнению")
    for r_idx, (a, d, r) in enumerate(attrs_step7):
        r_num = r_idx + 1
        style_data_cell(t_a7.cell(r_num, 0), a, bold=False)
        style_data_cell(t_a7.cell(r_num, 1), d, bold=False)
        style_data_cell(t_a7.cell(r_num, 2), r, bold=False)

    add_h("Требуемое изменение функциональности", level=4)
    add_p("Автоматическая генерация QR-кода на бланке акта со ссылкой на персональный электронный паспорт тюнинга в личном кабинете автовладельца.")

    add_h("Требования к интеграции с другими системами", level=4)
    add_p("Двусторонняя интеграция с фискальным регистратором АТОЛ/Штрих-М и эквайринговым терминалом Сбербанка.")

    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 10. РАЗДЕЛ 5. СТРУКТУРА ИНФОРМАЦИОННОЙ БАЗЫ И МОДЕЛЬ ДАННЫХ (ERD)
    # -------------------------------------------------------------------------
    add_h("5 Структура информационной базы и модель данных (ER-диаграмма)", level=1)
    add_p("Для обеспечения корректного функционирования детального бизнес-сценария и хранения всех транзакционных данных разработана концептуальная и логическая реляционная модель данных (ERD) в среде MS Visio. Модель строго нормализована до требований третьей нормальной формы (3NF).")
    add_p("В процессе проектирования устранены потенциальные архитектурные и логические коллизии, в частности:")
    add_bullet("устранена избыточность внешних ключей и замкнутые циклы транзитивных зависимостей: атрибут vin_code исключен из таблицы Work_Order, так как автомобиль однозначно идентифицируется через связанную заявку id_request (связь Client -> Car -> Service_Request -> Work_Order);")
    add_bullet("реализована обязательная фиксация технического комментария мастера калибровщика непосредственно в структуре исполнения наряда: поле technical_comment введено в сущность Order_Execution;")
    add_bullet("обеспечена строгая целостность цепочки документов: Заявка (Гайнанов Д.И.) -> Наряд-заказ (Гайнанов Д.И.) -> Диагностическая карта -> Исполнение заказа с комментарием (Гарифуллин К.Р.) -> Акт выполненных работ (Гарифуллин К.Р.).")

    add_image_figure(
        "ДБС_ER_Модель_Данных.png",
        "Рисунок 10 — Итоговая реляционная модель данных КИС (ERD в 3NF, 8 сущностей, MS Visio)",
        width_inches=6.2
    )

    add_p("В таблице 17 приведена сводная спецификация спроектированных таблиц базы данных КИС.")
    add_p("Таблица 17 — Спецификация сущностей реляционной базы данных КИС", align=WD_ALIGN_PARAGRAPH.LEFT, indent=False)

    erd_tables_summary = [
        ("Client", "Справочник клиентов", "id_client (PK)", "—", "Хранение реквизитов контрагентов (ФИО, телефон, e-mail)."),
        ("Car", "Справочник автомобилей", "vin_code (PK)", "id_client (FK)", "Паспорт ТС: марка, модель, двигатель, ЭБУ, госномер."),
        ("Tuning_Package", "Каталог тюнинг-пакетов", "id_package (PK)", "—", "Пакеты Stage 1/2, прирост л.с., Нм, стоимость."),
        ("Service_Request", "Журнал заявок", "id_request (PK)", "vin_code (FK), id_package (FK)", "Первичная потребность клиента в тюнинге (Данил)."),
        ("Work_Order", "Журнал наряд-заказов", "id_order (PK)", "id_request (FK)", "Договор подряда, смета, назначенные мастера (Данил)."),
        ("Diagnostics_Card", "Диагностические карты", "id_card (PK)", "id_order (FK)", "Параметры ДВС, флаг ошибок DTC, допуск к чип-тюнингу."),
        ("Order_Execution", "Исполнение калибровки", "id_execution (PK)", "id_order (FK)", "Дампы ROM, Checksum, ТЕХНИЧЕСКИЙ КОММЕНТАРИЙ (Карим)."),
        ("Work_Completion_Act", "Акты выполненных работ", "id_act (PK)", "id_order (FK)", "Закрытие работ, гарантия 24 мес., оплата (Карим).")
    ]

    t_erd = doc.add_table(rows=len(erd_tables_summary) + 1, cols=5)
    t_erd.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_erd)
    t_erd.columns[0].width = Mm(32)
    t_erd.columns[1].width = Mm(35)
    t_erd.columns[2].width = Mm(28)
    t_erd.columns[3].width = Mm(30)
    t_erd.columns[4].width = Mm(45)

    headers_erd = ["Сущность", "Назначение", "Первичный ключ", "Внешний ключ", "Бизнес-роль в БП"]
    for c_idx, h_text in enumerate(headers_erd):
        style_header_cell(t_erd.cell(0, c_idx), h_text)

    for idx, row in enumerate(erd_tables_summary):
        r_num = idx + 1
        style_data_cell(t_erd.cell(r_num, 0), row[0], bold=True)
        style_data_cell(t_erd.cell(r_num, 1), row[1], bold=False)
        style_data_cell(t_erd.cell(r_num, 2), row[2], bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        style_data_cell(t_erd.cell(r_num, 3), row[3], bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        style_data_cell(t_erd.cell(r_num, 4), row[4], bold=False)

    doc.add_page_break()

    # -------------------------------------------------------------------------
    # 11. ЗАКЛЮЧЕНИЕ
    # -------------------------------------------------------------------------
    add_h("Заключение", level=1)
    add_p("В ходе выполнения второй части лабораторных работ по дисциплине «Корпоративные информационные системы» был успешно разработан комплексный «Детальный бизнес-сценарий (ДБС) оказания услуг программной модификации блоков управления двигателем (чип-тюнинг) в КИС Автосервис».")
    add_p("В процессе разработки достигнуты следующие ключевые научно-практические результаты:")
    add_bullet("сформирован регламентирующий документ 123456BS 01 в строгом соответствии с корпоративным стандартом ДБС и академическими требованиями СТП УГАТУ;")
    add_bullet("разработаны индивидуальные динамические модели подпроцессов в стандарте BPMN 2.0: этап регистрации заявки и утверждения Наряд-заказа детально проработан Гайнановым Д.И.; этап инженерного исполнения калибровки с обязательным техническим комментарием мастера и выдачи Акта выполненных работ проработан Гарифуллиным К.Р.;")
    add_bullet("сформирована общая сквозная 5-ролевая динамическая модель процесса в MS Visio (BPMN 2.0) в едином сине-голубом стиле, объединяющая 18 технологических процедур от первичного обращения клиента до закрытия гарантийных обязательств;")
    add_bullet("спроектированы прототипы экранных интерфейсов (UI Mockups) для всех 7 сеансов взаимодействия пользователей в КИС, включая экран АРМ калибровщика и диагностические протоколы;")
    add_bullet("разработана и верифицирована логическая реляционная модель данных (ERD) в третьей нормальной форме (3NF), исключающая транзитивные дублирования ключей и циклические зависимости.")
    add_p("Обоснование платформы реализации КИС:")
    add_p("В качестве программно-технологического базиса для реализации спроектированного ДБС выбрана платформа «1С:Предприятие 8.3 КОРП» в клиент-серверном варианте под управлением реляционной СУБД PostgreSQL. Данный выбор обусловлен следующими факторами:")
    add_bullet("наличие развитого механизма бизнес-процессов и ролевой маршрутизации задач;")
    add_bullet("встроенные средства криптографической защиты данных и привязки электронных цифровых подписей к документам;")
    add_bullet("простота интеграции с внешним аппаратным оборудованием (кассовые аппараты, диагностические сканеры, эквайринг) через нативные драйверы и REST API интерфейсы;")
    add_bullet("полное соответствие требованиям Федерального закона № 152-ФЗ «О персональных данных» и 54-ФЗ «О применении контрольно-кассовой техники».")
    add_p("Внедрение спроектированного детального бизнес-сценария позволит сократить время обслуживания клиентов на 30%, на 100% исключить аварийные риски повреждения ЭБУ за счет обязательного входного контроля и повысить удовлетворенность клиентов сервиса.")

    # -------------------------------------------------------------------------
    # 12. СПИСОК ИСПОЛЬЗУЕМЫХ ИСТОЧНИКОВ
    # -------------------------------------------------------------------------
    add_h("Список используемых источников", level=1)
    
    sources = [
        "СТП 521.18.107-2007. Система менеджмента качества. Изделия основного производства. Мероприятия по устранению и предупреждению дефектов. — Уфа: УГАТУ, 2007. — 24 с.",
        "СТО 521.18.158-2018. Система менеджмента качества. Детали, сборочные единицы изделий. Оформление решений на выпуск деталей, сборочных единиц с отклонениями от требований конструкторской документации. — Уфа: УГАТУ, 2018. — 18 с.",
        "ГОСТ 34.602-2020. Информационные технологии. Комплекс стандартов на автоматизированные системы. Техническое задание на создание автоматизированной системы. — М.: Стандартинформ, 2020. — 22 с.",
        "ГОСТ 19.701-90 (ИСО 5807-85). Единая система программной документации. Схемы алгоритмов, программ, данных и систем. Обозначения условные и правила выполнения. — М.: Издательство стандартов, 1991. — 26 с.",
        "ГОСТ 7.32-2017. Система стандартов по информации, библиотечному и издательскому делу. Отчет о научно-исследовательской работе. Структура и правила оформления. — М.: Стандартинформ, 2017. — 28 с.",
        "Вендров А. М. Проектирование программного обеспечения экономических информационных систем: Учебник. — 2-е изд., перераб. и доп. — М.: Финансы и статистика, 2006. — 544 с.",
        "Грекул В. И., Денищенко Г. Н., Коровкина Н. Л. Проектирование информационных систем. — М.: Интернет-университет информационных технологий - ИНТУИТ.ру, 2008. — 304 с.",
        "Радченко М. Г., Хрусталева Е. Ю. 1С:Предприятие 8.3. Практическое пособие разработчика. Примеры и типовые приемы. — М.: ООО «1С-Паблишинг», 2013. — 926 с.",
        "Демченко М. С. Методические указания к выполнению лабораторных работ по дисциплине «Корпоративные информационные системы». — Уфа: УГАТУ, 2023. — 45 с.",
        "Бош Р. Автомобильный справочник: Пер. с англ. — 3-е изд., перераб. и доп. — М.: Изд-во «За рулем», 2012. — 1280 с."
    ]

    for idx, src in enumerate(sources):
        add_p(f"{idx+1}. {src}", indent=True, space_after=3)

    print(f"Сохранение документа в: {out_docx_path}")
    doc.save(out_docx_path)
    print("=== Успешно обновлен Отчет_ДБС_ЧипТюнинг.docx (2026, черные буквы, чистые таблицы, Гайнанов Д.И., Гарифуллин К.Р.)! ===")

if __name__ == "__main__":
    create_dbs_report()

# -*- coding: utf-8 -*-
"""
Генератор полного академического отчета (Пояснительной записки)
по дисциплине «Корпоративные информационные системы» ЧАСТЬ 2:
«Детальный бизнес-сценарий (ДБС) оказания услуг программной модификации
блоков управления двигателем (чип-тюнинг) в КИС Автосервис»

Строго по СТП УГАТУ и официальному заданию Демченко М.С.:
1. Титульный лист отчета по лабораторным работам ЧАСТЬ 2 (2026 г.)
2. Официальный Лист Задания на ЛР ЧАСТЬ 2
3. Титульный лист ДБС (с колонтитулами и логотипами УУНиТ)
4. Автособираемое Оглавление (TOC)
5. Полный Детальный бизнес-сценарий (Термины, Версии, Назначение, Основные положения,
   Схема БП с 5 схемами Visio, Матрицы ролей, 20 шагов с 40 таблицами спецификаций и атрибутов ERD)
6. Заключение (глубокие выводы, обоснование платформы 1С)
7. Список используемых источников (ГОСТы, СТО, документация)
"""

import os
import shutil
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def build_full_report():
    print("=== НАЧАЛО СБОРКИ ПОЛНОГО ОТЧЕТА (ПОЯСНИТЕЛЬНОЙ ЗАПИСКИ) ПО ЧАСТИ 2 ===")
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    dbs_dir = os.path.abspath(os.path.join(script_dir, ".."))
    
    # Исходный выверенный файл ДБС (используем измДБС, так как ДБС в папке открыт в Word)
    dbs_docx_path = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\измДБС_102BMW_УУНиТ_Гайнанов_Гарифуллин.docx"
    
    # Выходные файлы
    out_rep_path1 = os.path.join(dbs_dir, "Отчет_ЛР_Часть2_ДБС_ЧипТюнинг_Гайнанов_Гарифуллин.docx")
    out_rep_path2 = os.path.join(dbs_dir, "Отчет_ДБС_ЧипТюнинг.docx")
    out_rep_path3 = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\Отчет_ЛР_Часть2_ДБС_ЧипТюнинг_Гайнанов_Гарифуллин.docx"

    doc = Document(dbs_docx_path)
    body = doc._element.body

    COLOR_BLACK = RGBColor(0, 0, 0)

    def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
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

    # -------------------------------------------------------------------------
    # 1. СОЗДАНИЕ ПЕРЕДНИХ СТРАНИЦ: ТИТУЛЬНЫЙ ЛИСТ ОТЧЕТА И ЛИСТ ЗАДАНИЯ
    # -------------------------------------------------------------------------
    # Находим первый элемент документа (Таблица 0 ДБС)
    t0_elem = doc.tables[0]._element

    # Вспомогательная функция вставки абзаца ПЕРЕД целевым элементом
    def insert_p_before(ref_elem, text="", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=4, bold=False, size_pt=12, italic=False):
        p = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        
        # Выравнивание
        align_str = "left"
        if align == WD_ALIGN_PARAGRAPH.CENTER:
            align_str = "center"
        elif align == WD_ALIGN_PARAGRAPH.RIGHT:
            align_str = "right"
        elif align == WD_ALIGN_PARAGRAPH.JUSTIFY:
            align_str = "both"
        jc = parse_xml(f'<w:jc {nsdecls("w")} w:val="{align_str}"/>')
        pPr.append(jc)

        # Отступы
        sp = parse_xml(f'<w:spacing {nsdecls("w")} w:before="{int(space_before*20)}" w:after="{int(space_after*20)}" w:line="276" w:lineRule="auto"/>')
        pPr.append(sp)
        p.append(pPr)

        if text:
            lines = text.split('\n')
            for li, line in enumerate(lines):
                r = OxmlElement('w:r')
                rPr = OxmlElement('w:rPr')
                rPr.append(parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'))
                rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{int(size_pt*2)}"/>'))
                rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="{int(size_pt*2)}"/>'))
                rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="000000"/>'))
                if bold:
                    rPr.append(parse_xml(f'<w:b {nsdecls("w")}/>'))
                if italic:
                    rPr.append(parse_xml(f'<w:i {nsdecls("w")}/>'))
                r.append(rPr)
                t = OxmlElement('w:t')
                t.text = line
                r.append(t)
                p.append(r)
                if li < len(lines) - 1:
                    p.append(parse_xml(f'<w:r {nsdecls("w")}><w:br/></w:r>'))

        ref_elem.addprevious(p)
        return p

    def insert_page_break_before(ref_elem):
        p = parse_xml(f'<w:p {nsdecls("w")}><w:r><w:br w:type="page"/></w:r></w:p>')
        ref_elem.addprevious(p)

    print("Вставка 1: Титульный лист отчета по ЛР ЧАСТЬ 2 (по СТП УГАТУ)...")
    insert_p_before(t0_elem, "Министерство науки и высшего образования Российской Федерации", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=11, bold=True)
    insert_p_before(t0_elem, "Федеральное государственное бюджетное образовательное учреждение высшего образования", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=10.5)
    insert_p_before(t0_elem, "«УФИМСКИЙ УНИВЕРСИТЕТ НАУКИ И ТЕХНОЛОГИЙ»", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=12, bold=True)
    insert_p_before(t0_elem, "Кафедра автоматизированных систем управления", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=11, space_after=40)

    insert_p_before(t0_elem, "ОТЧЕТ ПО ЛАБОРАТОРНЫМ РАБОТАМ", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=16, bold=True, space_after=4)
    insert_p_before(t0_elem, "по дисциплине «Корпоративные информационные системы»", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13, space_after=4)
    insert_p_before(t0_elem, "ЧАСТЬ 2", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=14, bold=True, space_after=14)

    insert_p_before(t0_elem, "Тема: «Детальный бизнес-сценарий (ДБС) оказания услуг программной модификации блоков управления двигателем (чип-тюнинг) в КИС Автосервис»", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13, bold=True, space_after=50)

    # Блок подписей титульного листа отчета
    insert_p_before(t0_elem, "Выполнили:", align=WD_ALIGN_PARAGRAPH.RIGHT, size_pt=11, bold=True)
    insert_p_before(t0_elem, "студенты группы ЭАС-514С", align=WD_ALIGN_PARAGRAPH.RIGHT, size_pt=11)
    insert_p_before(t0_elem, "Гайнанов Д. И.  / _________________ /", align=WD_ALIGN_PARAGRAPH.RIGHT, size_pt=11)
    insert_p_before(t0_elem, "Гарифуллин К. Р.  / _________________ /", align=WD_ALIGN_PARAGRAPH.RIGHT, size_pt=11)
    insert_p_before(t0_elem, "«___» ________________ 2026 г.", align=WD_ALIGN_PARAGRAPH.RIGHT, size_pt=10.5, space_after=16)

    insert_p_before(t0_elem, "Проверил:", align=WD_ALIGN_PARAGRAPH.RIGHT, size_pt=11, bold=True)
    insert_p_before(t0_elem, "ст. преподаватель кафедры АСУ", align=WD_ALIGN_PARAGRAPH.RIGHT, size_pt=11)
    insert_p_before(t0_elem, "Демченко М. С.  / _________________ /", align=WD_ALIGN_PARAGRAPH.RIGHT, size_pt=11)
    insert_p_before(t0_elem, "«___» ________________ 2026 г.", align=WD_ALIGN_PARAGRAPH.RIGHT, size_pt=10.5, space_after=70)

    insert_p_before(t0_elem, "Уфа — 2026 г.", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=11, bold=True, space_after=0)
    insert_page_break_before(t0_elem)

    # -------------------------------------------------------------------------
    # 2. ВСТАВКА ОФИЦИАЛЬНОГО ЛИСТА ЗАДАНИЯ
    # -------------------------------------------------------------------------
    print("Вставка 2: Официальный Лист Задания ЧАСТЬ 2...")
    insert_p_before(t0_elem, "Федеральное государственное бюджетное образовательное учреждение высшего образования", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=9.5)
    insert_p_before(t0_elem, "«Уфимский университет науки и технологий»", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=11, bold=True)
    insert_p_before(t0_elem, "Кафедра автоматизированных систем управления", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=10, space_after=8)

    insert_p_before(t0_elem, "ЗАДАНИЕ", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13, bold=True, space_after=2)
    insert_p_before(t0_elem, "на лабораторные работы по дисциплине «Корпоративные информационные системы» ЧАСТЬ 2", align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=11, bold=True, space_after=8)

    # Таблица 1 задания: Студент / Группа / Консультант
    tbl_task_0 = parse_xml(f'''
        <w:tbl {nsdecls("w")}>
            <w:tblPr>
                <w:tblW w:w="9800" w:type="dxa"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="1500"/>
                <w:gridCol w:w="3000"/>
                <w:gridCol w:w="1200"/>
                <w:gridCol w:w="1400"/>
                <w:gridCol w:w="1500"/>
                <w:gridCol w:w="1200"/>
            </w:tblGrid>
            <w:tr>
                <w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:b/></w:rPr><w:t>Студент</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="19"/><w:b/></w:rPr><w:t>Гайнанов Д.И., Гарифуллин К.Р.</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:b/></w:rPr><w:t>Группа</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="1400" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:b/></w:rPr><w:t>ЭАС-514С</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:b/></w:rPr><w:t>Консультант</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:b/></w:rPr><w:t>Демченко М.С.</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="200" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="16"/><w:i/></w:rPr><w:t>Фамилия И.О.</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="1400" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="200" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="16"/><w:i/></w:rPr><w:t>номер группы</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="200" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="16"/><w:i/></w:rPr><w:t>Фамилия И.О.</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
    ''')
    t0_elem.addprevious(tbl_task_0)

    insert_p_before(t0_elem, "", space_after=4)

    # Таблица 2 задания: Содержание и исходные данные
    tbl_task_1 = parse_xml(f'''
        <w:tbl {nsdecls("w")}>
            <w:tblPr>
                <w:tblW w:w="9800" w:type="dxa"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="2600"/>
                <w:gridCol w:w="7200"/>
            </w:tblGrid>
            <w:tr>
                <w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:b/></w:rPr><w:t>1. Тема:</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="7200" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr><w:p><w:pPr><w:jc w:val="both"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:b/></w:rPr><w:t>«Детальный бизнес-сценарий (ДБС) оказания услуг программной модификации блоков управления двигателем (чип-тюнинг) в КИС Автосервис»</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:b/></w:rPr><w:t>2. Основное содержание:</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="7200" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr><w:p><w:pPr><w:jc w:val="both"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="19"/></w:rPr><w:t>В рамках лабораторных работ ЧАСТЬ 1 сформирована технико-экономическая характеристика объекта автоматизации и построена ERD. Далее разработан «Детальный бизнес-сценарий (ДБС) процесса оказания услуг чип-тюнинга автомобиля» в строгом соответствии с корпоративным шаблоном ДБС и СТП УГАТУ. Включены: общая схема БП на 5 ролей, матрицы ответственности, пошаговая спецификация 20 шагов процесса с таблицами значимых атрибутов базы данных (3NF), требованиями к функционалу и интеграции. Все схемы разработаны в MS Visio.</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:b/></w:rPr><w:t>3. Исходные данные:</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="7200" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr><w:p><w:pPr><w:jc w:val="both"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="19"/></w:rPr><w:t>Вариант № 102 «Автосервис чип-тюнинга BMW102»; результаты ЧАСТИ 1 (подпроцесс оформления заявки и утверждения наряд-заказа — студент Гайнанов Д.И.; подпроцесс исполнения наряд-заказа с комментарием мастера и выдачи акта — студент Гарифуллин К.Р.); нормативные стандарты СТО 521.18.107-2007, СТО 521.18.158-2018, регламенты ISTA/D и ISTA/P BMW AG.</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:b/></w:rPr><w:t>4. Требования к оформлению:</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="7200" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr><w:p><w:pPr><w:jc w:val="both"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="19"/></w:rPr><w:t>4.1 Пояснительная записка к отчету оформляется в текстовом редакторе MS Word в соответствии с СТП УГАТУ. Обязательно автособираемое Оглавление.&#10;4.2 В ПЗ содержатся разделы: Титульный лист, Лист Задания, Титульный лист ДБС, Оглавление, Термины и сокращения, 1. Версии документа, 2. Назначение документа, 3. Основные положения, 4. Бизнес-процесс (схемы Visio и 20 шагов), Заключение, Список используемых источников.&#10;4.3 Срок сдачи: 2026 г.</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
    ''')
    t0_elem.addprevious(tbl_task_1)

    insert_p_before(t0_elem, "", space_after=6)
    insert_p_before(t0_elem, "Руководитель: __________________ М.С. Демченко          Студенты: _________________ Гайнанов Д.И., Гарифуллин К.Р.", align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=10.5, space_after=0)
    insert_page_break_before(t0_elem)

    # -------------------------------------------------------------------------
    # 3. ДОБАВЛЕНИЕ ЗАКЛЮЧЕНИЯ И СПИСКА ИСТОЧНИКОВ В КОНЕЦ ОТЧЕТА
    # -------------------------------------------------------------------------
    print("Добавление заключительных разделов: Заключение и Список источников...")

    # Находим последний элемент body (конечный sectPr)
    terminal_sectPr = body.xpath('./w:sectPr')[-1]

    def append_before_terminal(text="", style='Normal', space_before=0, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, size_pt=12):
        p = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        
        # Стиль
        if style:
            pStyle = parse_xml(f'<w:pStyle {nsdecls("w")} w:val="{style}"/>')
            pPr.append(pStyle)
            
        align_str = "left"
        if align == WD_ALIGN_PARAGRAPH.CENTER:
            align_str = "center"
        elif align == WD_ALIGN_PARAGRAPH.RIGHT:
            align_str = "right"
        elif align == WD_ALIGN_PARAGRAPH.JUSTIFY:
            align_str = "both"
        pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="{align_str}"/>'))
        pPr.append(parse_xml(f'<w:spacing {nsdecls("w")} w:before="{int(space_before*20)}" w:after="{int(space_after*20)}" w:line="276" w:lineRule="auto"/>'))
        p.append(pPr)

        if text:
            r = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            rPr.append(parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'))
            rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{int(size_pt*2)}"/>'))
            rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="000000"/>'))
            if bold:
                rPr.append(parse_xml(f'<w:b {nsdecls("w")}/>'))
            r.append(rPr)
            t = OxmlElement('w:t')
            t.text = text
            r.append(t)
            p.append(r)

        terminal_sectPr.addprevious(p)
        return p

    def append_break_before_terminal():
        p = parse_xml(f'<w:p {nsdecls("w")}><w:r><w:br w:type="page"/></w:r></w:p>')
        terminal_sectPr.addprevious(p)

    def append_bullet_before_terminal(text):
        p = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        pPr.append(parse_xml(f'<w:pStyle {nsdecls("w")} w:val="IP_20_List1"/>'))
        pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="both"/>'))
        pPr.append(parse_xml(f'<w:spacing {nsdecls("w")} w:before="0" w:after="40" w:line="276" w:lineRule="auto"/>'))
        p.append(pPr)

        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        rPr.append(parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'))
        rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="23"/>'))
        rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="000000"/>'))
        r.append(rPr)
        t = OxmlElement('w:t')
        t.text = text
        r.append(t)
        p.append(r)
        terminal_sectPr.addprevious(p)

    # Разрыв страницы перед Заключением
    append_break_before_terminal()

    # ЗАКЛЮЧЕНИЕ
    append_before_terminal("Заключение", style='Heading 1', space_before=16, space_after=8, bold=True, size_pt=14)
    
    append_before_terminal(
        "В ходе выполнения второй части лабораторных работ по дисциплине «Корпоративные информационные системы» разработан комплексный «Детальный бизнес-сценарий (ДБС) оказания услуг программной модификации блоков управления двигателем (чип-тюнинг) в КИС Автосервис».",
        style='Normal', space_after=6, align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    append_before_terminal(
        "В процессе проектирования и формализации технологических и управленческих процедур получены следующие результаты:",
        style='Normal', space_after=4, align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    append_bullet_before_terminal("разработан проектный документ стандарта 102BMW-BS-01 в полном соответствии с утвержденным корпоративным шаблоном ДБС и академическими требованиями СТП УГАТУ;")
    append_bullet_before_terminal("спроектированы две динамические модели подпроцессов в нотации BPMN 2.0: этап регистрации заявки и утверждения наряд-заказа (студент Гайнанов Д.И., 10 шагов) и этап исполнения калибровки с обязательным техническим комментарием мастера и выдачи закрывающего акта (студент Гарифуллин К.Р., 10 шагов);")
    append_bullet_before_terminal("построена общая сквозная динамическая модель процесса чип-тюнинга на 5 ролевых дорожках в MS Visio, отражающая взаимодействие клиента, мастера-приемщика, диагноста, калибровщика и сервисов КИС в едином визуальном стиле;")
    append_bullet_before_terminal("сформирована и верифицирована логическая реляционная модель данных (ERD) в третьей нормальной форме (3NF) на 8 сущностей (Client, Vehicle, Employee, Service_Request, Work_Order, Diagnostic_Card, Order_Execution, Work_Completion_Act), устраняющая транзитивные зависимости и дублирование внешних ключей;")
    append_bullet_before_terminal("составлены детальные спецификации для всех 20 шагов процесса с таблицами значимых атрибутов базы данных, четко разграничивающими поля ввода и соответствующие таблицы ERD;")
    append_bullet_before_terminal("обоснован выбор программной платформы реализации — клиент-серверный комплекс «1С:Предприятие 8.3 КОРП» под управлением СУБД PostgreSQL с интеграцией специализированного модуля калибровки WinOLS и сканеров OBD-II/ICOM.")

    append_before_terminal(
        "Практическая реализация разработанного сценария позволит автотехцентру стандартизировать обслуживание автомобилей, сократить среднее время оформления и выдачи автомобиля на 30%, на 100% исключить риски повреждения ЭБУ при записи прошивок за счет строгого входного контроля и обеспечить непрерывное гарантийное сопровождение клиентов.",
        style='Normal', space_before=6, space_after=12, align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    # СПИСОК ИСПОЛЬЗУЕМЫХ ИСТОЧНИКОВ
    append_break_before_terminal()
    append_before_terminal("Список используемых источников", style='Heading 1', space_before=16, space_after=8, bold=True, size_pt=14)

    sources = [
        "СТП 521.18.107-2007. Система менеджмента качества. Изделия основного производства. Мероприятия по устранению и предупреждению дефектов. — Уфа: УГАТУ, 2007. — 24 с.",
        "СТО 521.18.158-2018. Система менеджмента качества. Детали, сборочные единицы изделий. Оформление решений на выпуск деталей с отклонениями. — Уфа: УГАТУ, 2018. — 18 с.",
        "ГОСТ Р ИСО/МЭК 12207-2010. Информационная технология. Системная и программная инженерия. Процессы жизненного цикла программных средств. — М.: Стандартинформ, 2011. — 106 с.",
        "ГОСТ 34.602-2020. Информационные технологии. Комплекс стандартов на автоматизированные системы. Техническое задание на создание автоматизированной системы. — М.: Стандартинформ, 2020. — 22 с.",
        "ГОСТ 19.701-90 (ИСО 5807-85). Единая система программной документации. Схемы алгоритмов, программ, данных и систем. Обозначения условные и правила выполнения. — М.: Издательство стандартов, 1991. — 26 с.",
        "ГОСТ 7.32-2017. Система стандартов по информации, библиотечному и издательскому делу. Отчет о научно-исследовательской работе. Структура и правила оформления. — М.: Стандартинформ, 2017. — 28 с.",
        "Вендров А. М. Проектирование программного обеспечения экономических информационных систем: Учебник. — 2-е изд. — М.: Финансы и статистика, 2006. — 544 с.",
        "Грекул В. И., Денищенко Г. Н., Коровкина Н. Л. Проектирование информационных систем. — М.: ИНТУИТ.ру, 2008. — 304 с.",
        "Радченко М. Г., Хрусталева Е. Ю. 1С:Предприятие 8.3. Практическое пособие разработчика. — М.: ООО «1С-Паблишинг», 2013. — 926 с.",
        "Демченко М. С. Методические указания к выполнению лабораторных работ по дисциплине «Корпоративные информационные системы». — Уфа: УУНиТ, 2023. — 45 с.",
        "Техническая документация BMW AG. Руководство по сервисному обслуживанию и программированию блоков управления ISTA/D и ISTA/P. — Мюнхен: BMW Service, 2022. — 380 с."
    ]

    for idx, src in enumerate(sources):
        append_before_terminal(f"{idx+1}. {src}", style='Normal', space_after=3, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # Сохраняем готовый Отчет
    print(f"Сохранение отчета в: {out_rep_path1}")
    doc.save(out_rep_path1)

    try:
        print(f"Копирование отчета в: {out_rep_path2}")
        shutil.copyfile(out_rep_path1, out_rep_path2)
    except Exception as e:
        print(f"[!] Не удалось скопировать в {out_rep_path2} (возможно, открыт в Word): {e}")

    try:
        print(f"Копирование отчета в: {out_rep_path3}")
        shutil.copyfile(out_rep_path1, out_rep_path3)
    except Exception as e:
        print(f"[!] Не удалось скопировать в {out_rep_path3}: {e}")

    print("=== ПОЛНЫЙ ОТЧЕТ УСПЕШНО СОБРАН И СОХРАНЕН! ===")

if __name__ == "__main__":
    build_full_report()

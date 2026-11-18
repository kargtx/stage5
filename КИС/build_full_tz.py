import os
import re
import zipfile
import shutil
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

base_dir = r"C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС"
tpl_tz = os.path.join(base_dir, "ШАБЛОН ТЗ.docx")
tpl_dbs = os.path.join(base_dir, "ШАБЛОН ДБС.docx")

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('shd'):
            tcPr.remove(child)
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_text(cell, text, bold=False, italic=False, font_name="Times New Roman", size_pt=9.0, align='left', color_rgb=(0,0,0)):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.05
    if align == 'center':
        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right':
        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)
    return p

def ensure_table_rows(table, count):
    while len(table.rows) < count:
        table.add_row()
    while len(table.rows) > count:
        tr = table.rows[-1]._tr
        tr.getparent().remove(tr)

def inject_clean_headers(docx_path, student_name, report_title):
    temp_dir = docx_path + "_unpacked"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    
    with zipfile.ZipFile(docx_path, 'r') as z:
        z.extractall(temp_dir)
        
    with zipfile.ZipFile(tpl_dbs, 'r') as zd:
        h1_xml = zd.read('word/header1.xml').decode('utf-8')
        h1_rels = zd.read('word/_rels/header1.xml.rels').decode('utf-8')
        h2_xml = zd.read('word/header2.xml').decode('utf-8')
        h2_rels = zd.read('word/_rels/header2.xml.rels').decode('utf-8')
        img1_bytes = zd.read('word/media/image1.png')
        img2_bytes = zd.read('word/media/image2.png')

    def customize_header(xml_text):
        res = xml_text.replace("Детальный бизнес-сценарий", "Частное техническое задание")
        res = re.sub(r'102BMW-BS-01\s*«[^»]+»', f"на разработку отчетной формы «{report_title}»", res)
        return res

    h1_xml_custom = customize_header(h1_xml)
    h2_xml_custom = customize_header(h2_xml)

    with open(os.path.join(temp_dir, 'word', 'header1.xml'), 'w', encoding='utf-8') as f:
        f.write(h1_xml_custom)
    with open(os.path.join(temp_dir, 'word', '_rels', 'header1.xml.rels'), 'w', encoding='utf-8') as f:
        f.write(h1_rels)
    with open(os.path.join(temp_dir, 'word', 'header2.xml'), 'w', encoding='utf-8') as f:
        f.write(h2_xml_custom)
    with open(os.path.join(temp_dir, 'word', '_rels', 'header2.xml.rels'), 'w', encoding='utf-8') as f:
        f.write(h2_rels)
    
    media_dir = os.path.join(temp_dir, 'word', 'media')
    os.makedirs(media_dir, exist_ok=True)
    with open(os.path.join(media_dir, 'image1.png'), 'wb') as f:
        f.write(img1_bytes)
    with open(os.path.join(media_dir, 'image2.png'), 'wb') as f:
        f.write(img2_bytes)

    if os.path.exists(docx_path):
        os.remove(docx_path)
    with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, temp_dir)
                zout.write(full_path, rel_path)
                
    shutil.rmtree(temp_dir)
    print(f"Injected clean headers & logos into {os.path.basename(docx_path)}")


def build_garifullin_tz():
    print("--- Building TZ for Garifullin K.R. ---")
    doc = docx.Document(tpl_tz)
    
    # Title Page
    doc.paragraphs[0].text = "ФГБОУ ВО «Уфимский университет науки и технологий»"
    doc.paragraphs[1].text = "Частное техническое задание"
    doc.paragraphs[2].text = "на разработку отчетной формы «Отчет по количеству исполненных заявок за период»"
    
    # Section 0 footer
    for p in doc.sections[0].first_page_footer.paragraphs:
        if '2023' in p.text:
            p.text = p.text.replace('2023', '2026')
            
    # Table 0: Approval
    t0 = doc.tables[0]
    set_cell_text(t0.rows[0].cells[0], "УТВЕРЖДАЮ:", bold=True, size_pt=9.5)
    set_cell_text(t0.rows[1].cells[0], "Генеральный директор ООО «BMW102»", size_pt=9)
    set_cell_text(t0.rows[1].cells[1], "Гарифуллин К.Р.", size_pt=9)
    set_cell_text(t0.rows[1].cells[2], "/ ____________ /", size_pt=9, align='center')
    set_cell_text(t0.rows[1].cells[3], "« 25 » марта 2026 г.", size_pt=9, align='right')
    set_cell_text(t0.rows[2].cells[0], "СОГЛАСОВАНО:", bold=True, size_pt=9.5)
    set_cell_text(t0.rows[3].cells[0], "Ст. преподаватель кафедры АСУ УУНиТ", size_pt=9)
    set_cell_text(t0.rows[3].cells[1], "Демченко М.С.", size_pt=9)
    set_cell_text(t0.rows[3].cells[2], "/ ____________ /", size_pt=9, align='center')
    set_cell_text(t0.rows[3].cells[3], "« 26 » марта 2026 г.", size_pt=9, align='right')
    while len(t0.rows) > 4:
        tr = t0.rows[-1]._tr
        tr.getparent().remove(tr)
        
    # Table 1: Terms
    t1 = doc.tables[1]
    terms_g = [
        ("КИС", "Корпоративная информационная система; программный комплекс автоматизации автотехцентра (КИС «BMW102-ERP»)."),
        ("ОФ", "Отчетная форма; регламентированный выходной электронный или печатный документ аналитической отчетности."),
        ("ЗДФ", "Запрос на доработку функциональности системы; документ, определяющий требования к расширению функциональности."),
        ("ДБС", "Детальный бизнес-сценарий (102BMW-BS-01 «Организация и исполнение услуг чип-тюнинга и ТО автопарка УУНиТ»)."),
        ("Наряд-заказ", "Первичный производственный документ, регламентирующий перечень работ, калькуляцию и назначенного мастера."),
        ("Акт выполненных работ", "Закрывающий юридический документ, фиксирующий факт надлежащего исполнения наряд-заказа."),
        ("ЭБУ (ECU)", "Электронный блок управления двигателем внутреннего сгорания (Engine Control Unit)."),
        ("Чип-тюнинг (Stage 1/2)", "Программная модификация калибровочных таблиц ЭБУ для оптимизации мощности и крутящего момента ДВС."),
        ("СУБД", "Система управления реляционными базами данных (PostgreSQL 16), обеспечивающая хранение и целостность данных."),
        ("СКД", "Система компоновки данных платформы 1С для интерактивного формирования сложных сводных аналитических отчетов.")
    ]
    ensure_table_rows(t1, len(terms_g) + 1)
    set_cell_text(t1.rows[0].cells[0], "Термин", bold=True, size_pt=9.5, align='center')
    set_cell_text(t1.rows[0].cells[1], "Описание", bold=True, size_pt=9.5, align='center')
    set_cell_background(t1.rows[0].cells[0], "E2E8F0")
    set_cell_background(t1.rows[0].cells[1], "E2E8F0")
    for idx, (term, desc) in enumerate(terms_g):
        row = t1.rows[idx + 1]
        set_cell_text(row.cells[0], term, bold=True, size_pt=8.5)
        set_cell_text(row.cells[1], desc, size_pt=8.5)
        if idx % 2 == 1:
            set_cell_background(row.cells[0], "F8FAFC")
            set_cell_background(row.cells[1], "F8FAFC")

    # Table 2: Version History
    t2 = doc.tables[2]
    v_headers = ["Автор", "Дата", "Компания / Подразделение", "Версия", "Описание изменений"]
    for ci, h in enumerate(v_headers):
        set_cell_text(t2.rows[0].cells[ci], h, bold=True, size_pt=9, align='center')
        set_cell_background(t2.rows[0].cells[ci], "E2E8F0")
    v_rows = [
        ("Гарифуллин К.Р.", "25.03.2026", "ЭАС-514С / ООО «BMW102»", "1.0", "Разработка первой редакции частного технического задания на разработку формы «Отчет по количеству исполненных заявок за период»."),
        ("Демченко М.С.", "26.03.2026", "Кафедра АСУ ФГБОУ ВО «УУНиТ»", "1.1", "Рецензирование документа, согласование атрибутивного состава формы и методики испытаний."),
        ("", "", "", "", ""),
        ("", "", "", "", ""),
        ("", "", "", "", ""),
        ("", "", "", "", "")
    ]
    ensure_table_rows(t2, len(v_rows) + 1)
    for ri, rvals in enumerate(v_rows):
        row = t2.rows[ri + 1]
        for ci, val in enumerate(rvals):
            set_cell_text(row.cells[ci], val, size_pt=8.5, align='center' if ci in [1, 3] else 'left')

    # Paragraphs 22-24: Section 1
    doc.paragraphs[22].text = (
        "Настоящий документ разработан в рамках выполнения лабораторного практикума и курсового проектирования "
        "по дисциплине «Корпоративные информационные системы» в весенний семестр 2026 учебного года. "
        "Проект выполнен студентом кафедры АСУ УУНиТ учебной группы ЭАС-514С Гарифуллиным Каримом Рамилевичем (вариант №102). "
        "Предметная область: «Сервисное обслуживание, диагностика ЭСУД и программная оптимизация (чип-тюнинг) автомобильного парка Заказчика»."
    )
    doc.paragraphs[23].text = (
        "Подпроцесс: «Исполнение и закрытие наряд-заказов цеха чип-тюнинга». Разрабатываемый программный модуль реализует "
        "требования Заказчика — ФГБОУ ВО «Уфимский университет науки и технологий» (УУНиТ) — к аналитической отчетности автотехцентра ООО «BMW102»."
    )
    doc.paragraphs[24].text = (
        "Настоящее частное техническое задание регламентирует разработку выходной аналитической формы в среде КИС «BMW102-ERP» "
        "и является неотъемлемой составной частью сквозного проектного решения 102BMW-BS-01 «Организация и исполнение услуг чип-тюнинга и ТО автопарка»."
    )

    # Paragraphs 34, 36: Section 2
    doc.paragraphs[34].text = (
        "В настоящем ЗДФ описаны требования и вариант реализации доработки в области функционального подпроцесса "
        "«Исполнение и закрытие наряд-заказов цеха чип-тюнинга» в части добавления функционала интерактивного формирования и вывода на печать "
        "из КИС «BMW102-ERP» регламентированной аналитической отчетной формы «Отчет по количеству исполненных заявок за период». "
        "Действие ЗДФ подчинено рамкам сквозного детального бизнес-сценария 102BMW-BS-01."
    )
    doc.paragraphs[36].text = (
        "Реализовать вывод из КИС «BMW102-ERP» регламентированной формы аналитического отчета «Отчет по количеству исполненных заявок за период» "
        "следующего вида (рисунок 1). Полный перечень атрибутов данного аналитического отчета представлен в Таблице 2."
    )

    # Figure 1: Report Mockup
    doc.paragraphs[38].text = ""
    doc.paragraphs[38].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[38].paragraph_format.space_before = Pt(6)
    doc.paragraphs[38].paragraph_format.space_after = Pt(4)
    doc.paragraphs[38].add_run().add_picture(os.path.join(base_dir, 'tz_garifullin_fig1_report.png'), width=Inches(6.0))
    doc.paragraphs[39].text = "Рисунок 1 – Макет формы аналитического отчета «Отчет по количеству исполненных заявок за период»"
    doc.paragraphs[40].text = ""

    # Paragraph 42, 44, 45: Use Case
    doc.paragraphs[42].text = (
        "Последовательность взаимодействия оператора (старшего мастера или мастера-приемщика) и системы КИС «BMW102-ERP» "
        "при формировании отчетной формы регламентируется сценарием использования, представленным в Таблице 1."
    )
    doc.paragraphs[44].text = "3.2.1 Сценарий использования «Формирование отчета по исполненным заявкам за период»"
    doc.paragraphs[45].text = "Таблица 1 – Спецификация сценария использования ОФ"

    # Table 3: Use Case Table
    t3 = doc.tables[3]
    uc_g = [
        ("Краткое описание", "Формирование сводной ведомости исполненных и закрытых наряд-заказов автотехцентра за указанный временной интервал с фильтрацией по мастеру-диагносту и расчетными итогами."),
        ("Действующие лица", "Старший мастер цеха (Мастер-приемщик), Руководитель сервиса, КИС «BMW102-ERP»."),
        ("Триггер (Событие)", "Необходимость анализа загрузки цеха, расчета сдельной оплаты труда мастеров и контроля объемов выполненных работ по чип-тюнингу автопарка Заказчика."),
        ("Предусловия", "1. Пользователь успешно авторизован в КИС с правами «Старший мастер» или «Руководитель».\n2. В БД зафиксированы наряд-заказы в статусах «Выполнена» / «Закрыта» за анализируемый период."),
        ("Основной поток событий", "1. Пользователь переходит в раздел «Аналитика и отчеты -> Цех чип-тюнинга -> Отчет по исполненным заявкам».\n2. Система открывает сеанс-конфигуратор отчета.\n3. Пользователь задает диапазон дат (дата начала, дата окончания) и при необходимости выбирает мастера-исполнителя.\n4. Пользователь нажимает кнопку «Сформировать» (F5).\n5. Система выполняет агрегирующий SQL-запрос к БД PostgreSQL и генерирует печатный макет.\n6. Система отображает готовую форму на экране с возможностью постраничного просмотра, печати или экспорта."),
        ("Альтернативный поток А1", "За выбранный период исполненные заявки не найдены. Система выводит сообщение «За выбранный период исполненных заявок не обнаружено» и очищает табличную часть."),
        ("Альтернативный поток А2", "Пользователь указал некорректный интервал дат (дата окончания меньше даты начала). Система подсвечивает поле ошибки и блокирует формирование до исправления."),
        ("Постусловия", "Сформированная печатная форма отображена на экране, зафиксирована дата и время построения, отчет готов к выводу на печать или выгрузке в Excel/PDF."),
        ("Частота использования", "Ежедневно в конце рабочей смены, еженедельно при подведении итогов, ежемесячно при расчете сдельной оплаты."),
        ("Требования к интерфейсу", "Открытие конфигуратора не более 1 сек, построение выборки объемом до 5 000 строк не более 2.5 сек, поддержка горячих клавиш Ctrl+P, Esc."),
        ("Ограничения безопасности", "Доступ к суммам оплат и расчету стоимости нормо-часа предоставляется только пользователям с подтвержденной ролью доступа."),
        ("Особые требования", "Формат выгрузки в Excel должен сохранять числовые типы данных для ячеек стоимости без потери точности.")
    ]
    ensure_table_rows(t3, len(uc_g) + 1)
    set_cell_text(t3.rows[0].cells[0], "Разделы сценария", bold=True, size_pt=9, align='center')
    set_cell_text(t3.rows[0].cells[1], "Описание разделов сценария", bold=True, size_pt=9, align='center')
    set_cell_background(t3.rows[0].cells[0], "E2E8F0")
    set_cell_background(t3.rows[0].cells[1], "E2E8F0")
    for idx, (sec_name, sec_val) in enumerate(uc_g):
        row = t3.rows[idx + 1]
        set_cell_text(row.cells[0], sec_name, bold=True, size_pt=8.5)
        set_cell_text(row.cells[1], sec_val, size_pt=8.5)
        if idx % 2 == 1:
            set_cell_background(row.cells[0], "F8FAFC")
            set_cell_background(row.cells[1], "F8FAFC")

    # Paragraphs 48-53
    doc.paragraphs[48].text = (
        "Архитектура доработки базируется на типовых механизмах платформы КИС «BMW102-ERP» с использованием СУБД PostgreSQL 16 "
        "и Системы компоновки данных (СКД). Форма отчета генерируется на основе параметризованного SQL-представления v_executed_orders_report, "
        "обеспечивающего выборку и расчет агрегатов по закрытым наряд-заказам за выбранный период дат."
    )
    doc.paragraphs[49].text = ""
    doc.paragraphs[49].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[49].paragraph_format.space_before = Pt(6)
    doc.paragraphs[49].paragraph_format.space_after = Pt(4)
    doc.paragraphs[49].add_run().add_picture(os.path.join(base_dir, 'tz_garifullin_fig2_dialog.png'), width=Inches(5.8))
    doc.paragraphs[50].text = "Рисунок 2 – Макет экранной формы сеанса-конфигуратора параметров отчета"
    doc.paragraphs[51].text = (
        "Сеанс-конфигуратор отчета представляет собой модальное диалоговое окно, позволяющее оператору гибко настраивать параметры выборки: "
        "диапазон дат, фильтр по мастеру-диагносту, а также включать детализацию по выполненным технологическим операциям чип-тюнинга."
    )
    doc.paragraphs[52].text = "Атрибутивный состав выходной формы отчета приведен в Таблице 2."
    doc.paragraphs[53].text = "Таблица 2 – Атрибутивный состав формы отчета"

    # Table 4: Report Attributes
    t4 = doc.tables[4]
    attrs_g = [
        ("1", "Номер наряд-заказа", "Строка (15)", "ЗН-2026-00102", "Уникальный регистрационный номер наряд-заказа цеха."),
        ("2", "Дата закрытия", "Дата (ДД.ММ.ГГГГ)", "24.03.2026", "Фактическая дата завершения работ и подписания акта."),
        ("3", "Автомобиль / Госномер", "Строка (40)", "BMW 530d xDrive (р102уу102)", "Модель транспортного средства и государственный регистрационный номер."),
        ("4", "Услуга / ПО чип-тюнинга", "Строка (60)", "Чип-тюнинг Stage 1 + Отключение EGR", "Наименование выполненного пакета программной калибровки ЭБУ."),
        ("5", "Мастер-исполнитель", "Строка (50)", "Гарифуллин К.Р. (Калибровщик)", "ФИО мастера цеха, производившего перепрограммирование ЭБУ."),
        ("6", "Стоимость работ, руб.", "Число (12, 2)", "35 000,00", "Итоговая расчетная стоимость работ по закрытому наряд-заказу.")
    ]
    ensure_table_rows(t4, len(attrs_g) + 1)
    set_cell_text(t4.rows[0].cells[0], "№", bold=True, size_pt=9, align='center')
    set_cell_text(t4.rows[0].cells[1], "Атрибут", bold=True, size_pt=9, align='center')
    set_cell_text(t4.rows[0].cells[2], "Тип данных", bold=True, size_pt=9, align='center')
    set_cell_text(t4.rows[0].cells[3], "Пример заполнения", bold=True, size_pt=9, align='center')
    set_cell_text(t4.rows[0].cells[4], "Примечание", bold=True, size_pt=9, align='center')
    for ci in range(5): set_cell_background(t4.rows[0].cells[ci], "E2E8F0")
    for idx, (num, name, dt, ex, note) in enumerate(attrs_g):
        row = t4.rows[idx + 1]
        set_cell_text(row.cells[0], num, size_pt=8.5, align='center')
        set_cell_text(row.cells[1], name, bold=True, size_pt=8.5)
        set_cell_text(row.cells[2], dt, italic=True, size_pt=8, align='center')
        set_cell_text(row.cells[3], ex, size_pt=8.5)
        set_cell_text(row.cells[4], note, size_pt=8.5)
        if idx % 2 == 1:
            for ci in range(5): set_cell_background(row.cells[ci], "F8FAFC")

    # Paragraphs 54-63
    doc.paragraphs[54].text = "Связи таблиц базы данных КИС, из которых формируется отчет, приведены на схеме (рисунок 3):"
    p_img3 = doc.paragraphs[55].insert_paragraph_before()
    p_img3.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    p_img3.paragraph_format.space_before = Pt(6)
    p_img3.paragraph_format.space_after = Pt(4)
    p_img3.add_run().add_picture(os.path.join(base_dir, 'tz_garifullin_fig3_erd.png'), width=Inches(6.0))
    doc.paragraphs[56].text = "Рисунок 3 – Реляционная модель данных (ERD) подпроцесса исполнения заявок"
    doc.paragraphs[58].text = (
        "Запуск формирования формы выполняется старшим мастером посредством выбора отчета через меню "
        "«Аналитика и отчеты -> Цех чип-тюнинга -> Отчет по исполненным заявкам» на панели команд КИС «BMW102-ERP». "
        "При этом открывается сеанс-конфигуратор со следующим набором атрибутов (Таблица 3):"
    )
    doc.paragraphs[59].text = "Таблица 3 – Атрибуты сеанса-конфигуратора ОФ"

    # Table 5: Configurator Attributes
    t5 = doc.tables[5]
    cfg_g = [
        ("1", "Период формирования (С ... По ...)", "Период дат (ДД.ММ.ГГГГ)", "Обязательный фильтр диапазона дат закрытия наряд-заказов (по умолчанию — текущий месяц)."),
        ("2", "Мастер-исполнитель", "Ссылка на справочник «Сотрудники»", "Необязательный фильтр. При незаполненном значении выводятся заявки всех мастеров."),
        ("3", "Статус заявки", "Список значений («Выполнена», «Закрыта»)", "Фильтр по производственным статусам. По умолчанию включены все закрытые наряд-заказы.")
    ]
    ensure_table_rows(t5, len(cfg_g) + 1)
    set_cell_text(t5.rows[0].cells[0], "№", bold=True, size_pt=9, align='center')
    set_cell_text(t5.rows[0].cells[1], "Наименование", bold=True, size_pt=9, align='center')
    set_cell_text(t5.rows[0].cells[2], "Тип", bold=True, size_pt=9, align='center')
    set_cell_text(t5.rows[0].cells[3], "Примечание", bold=True, size_pt=9, align='center')
    for ci in range(4): set_cell_background(t5.rows[0].cells[ci], "E2E8F0")
    for idx, (num, name, dt, note) in enumerate(cfg_g):
        row = t5.rows[idx + 1]
        set_cell_text(row.cells[0], num, size_pt=8.5, align='center')
        set_cell_text(row.cells[1], name, bold=True, size_pt=8.5)
        set_cell_text(row.cells[2], dt, size_pt=8.5, align='center')
        set_cell_text(row.cells[3], note, size_pt=8.5)
        if idx % 2 == 1:
            for ci in range(4): set_cell_background(row.cells[ci], "F8FAFC")

    # Paragraphs 61, 63, 64
    doc.paragraphs[61].text = (
        "На распечатанном бланке отчета автоматически выводится заголовок формы отчета «Отчет по количеству исполненных заявок за период», "
        "наименование автотехцентра ООО «BMW102», интервал дат выборки, текущая дата печати и ФИО оператора."
    )
    doc.paragraphs[63].text = "Источники атрибутов формы с логикой их формирования указаны в Таблице 4."
    doc.paragraphs[64].text = "Таблица 4 – Источники атрибутов формы с логикой формирования"

    # Table 6: Sources
    t6 = doc.tables[6]
    sources_g = [
        ("1", "Номер наряд-заказа", "ЗН-2026-00102", "Строка (15)", "Таблица work_orders, поле order_number", "Прямая выборка из таблицы наряд-заказов."),
        ("2", "Дата закрытия", "24.03.2026", "Дата", "Таблица work_orders, поле completion_date", "Дата фактического подписания акта сдачи-приемки работ."),
        ("3", "Автомобиль / Госномер", "BMW 530d (р102уу102)", "Строка (40)", "Таблица cars, поля model, plate_number", "Конкатенация: cars.model || ' (' || cars.plate_number || ')'."),
        ("4", "Услуга / ПО", "Stage 1 + EGR off", "Строка (60)", "Таблица services, поле service_name", "Связка work_order_items.service_id = services.service_id."),
        ("5", "Мастер-исполнитель", "Гарифуллин К.Р.", "Строка (50)", "Таблица employees, поле full_name", "Связка work_orders.master_id = employees.employee_id."),
        ("6", "Стоимость, руб.", "35 000,00", "Число (12, 2)", "Таблица work_orders, поле total_cost", "Сумма по строкам спецификации заказа: SUM(price * qty).")
    ]
    ensure_table_rows(t6, len(sources_g) + 1)
    set_cell_text(t6.rows[0].cells[0], "№", bold=True, size_pt=8.5, align='center')
    set_cell_text(t6.rows[0].cells[1], "Наименование", bold=True, size_pt=8.5, align='center')
    set_cell_text(t6.rows[0].cells[2], "Пример", bold=True, size_pt=8.5, align='center')
    set_cell_text(t6.rows[0].cells[3], "Тип данных", bold=True, size_pt=8.5, align='center')
    set_cell_text(t6.rows[0].cells[4], "Источник (ERD)", bold=True, size_pt=8.5, align='center')
    set_cell_text(t6.rows[0].cells[5], "Реализация / SQL", bold=True, size_pt=8.5, align='center')
    for ci in range(6): set_cell_background(t6.rows[0].cells[ci], "E2E8F0")
    for idx, (num, name, ex, dt, src, real) in enumerate(sources_g):
        row = t6.rows[idx + 1]
        set_cell_text(row.cells[0], num, size_pt=8, align='center')
        set_cell_text(row.cells[1], name, bold=True, size_pt=8)
        set_cell_text(row.cells[2], ex, size_pt=8)
        set_cell_text(row.cells[3], dt, italic=True, size_pt=7.5, align='center')
        set_cell_text(row.cells[4], src, size_pt=8)
        set_cell_text(row.cells[5], real, size_pt=8)
        if idx % 2 == 1:
            for ci in range(6): set_cell_background(row.cells[ci], "F8FAFC")

    # Paragraphs 68, 71, 73, 74, 77, 80
    doc.paragraphs[68].text = (
        "Ограничения и допущения: формирование отчета осуществляется исключительно по закрытым наряд-заказам. "
        "В выборку не включаются предварительные заявки со статусами «Черновик» и «Отменена». "
        "При отсутствии данных за указанный интервал времени система выводит диагностическое сообщение без генерации пустых таблиц."
    )
    doc.paragraphs[71].text = (
        "Программа и методика приемочных испытаний регламентируют условия, порядок и сценарии проверки разработанной формы отчета "
        "на соответствие требованиям настоящего ТЗ перед вводом в промышленную эксплуатацию."
    )
    doc.paragraphs[73].text = (
        "Для проведения испытаний требуется функционирующий тестовый контур КИС «BMW102-ERP» на базе PostgreSQL 16 "
        "с наполненными справочниками сотрудников, услуг и тестовым пулом закрытых наряд-заказов (не менее 20 записей)."
    )
    doc.paragraphs[74].text = "Предварительных настроек клиентских рабочих мест, сверх штатных средств КИС «BMW102-ERP» и ОС Windows, не требуется."
    doc.paragraphs[77].text = "Программа испытаний включает перечень контрольных сценариев проверки (Таблица 5):"
    doc.paragraphs[80].text = "Таблица 5 – Программа приемочных испытаний"

    # Table 7: Test Scenarios
    t7 = doc.tables[7]
    tests_g = [
        ("1", "Формирование отчета за стандартный период (текущий месяц)", "ТР-01", "Проверка корректности извлечения данных и агрегации итогов."),
        ("2", "Фильтрация выборки по конкретному мастеру-диагносту", "ТР-02", "Проверка корректности наложения отбора по сотруднику."),
        ("3", "Формирование отчета за период с отсутствием исполненных заявок", "ТР-03", "Проверка реакции системы и вывода информационного сообщения."),
        ("4", "Экспорт сформированного отчета в форматы PDF и MS Excel", "ТР-04", "Проверка точности сохранения стилей разметки и числовых ячеек.")
    ]
    ensure_table_rows(t7, len(tests_g) + 1)
    set_cell_text(t7.rows[0].cells[0], "№ п/п", bold=True, size_pt=9, align='center')
    set_cell_text(t7.rows[0].cells[1], "Наименование сценария", bold=True, size_pt=9, align='center')
    set_cell_text(t7.rows[0].cells[2], "Номер требования", bold=True, size_pt=9, align='center')
    set_cell_text(t7.rows[0].cells[3], "Примечание", bold=True, size_pt=9, align='center')
    for ci in range(4): set_cell_background(t7.rows[0].cells[ci], "E2E8F0")
    for idx, (num, name, req, note) in enumerate(tests_g):
        row = t7.rows[idx + 1]
        set_cell_text(row.cells[0], num, size_pt=8.5, align='center')
        set_cell_text(row.cells[1], name, bold=True, size_pt=8.5)
        set_cell_text(row.cells[2], req, size_pt=8.5, align='center')
        set_cell_text(row.cells[3], note, size_pt=8.5)
        if idx % 2 == 1:
            for ci in range(4): set_cell_background(row.cells[ci], "F8FAFC")

    # Paragraphs 82, 84
    doc.paragraphs[82].text = "Пошаговая методика проведения испытаний приведена в Таблице 6:"
    doc.paragraphs[84].text = "Таблица 6 – Методика проведения испытаний"

    # Table 8: Test Steps
    t8 = doc.tables[8]
    steps_g = [
        ("Сценарий 1", "Запустить конфигуратор, указать период 01.03.2026 - 31.03.2026, нажать «Сформировать».", "Отчет сформирован, выведены все 5 закрытых заявок, итоговая сумма 185 000,00 руб. рассчитана верно."),
        ("Сценарий 2", "Выбрать в поле «Мастер» сотрудника «Гарифуллин К.Р.», нажать «Сформировать».", "В таблице отображаются только наряд-заказы мастера Гарифуллина (2 заявки), итог 80 000,00 руб."),
        ("Сценарий 3", "Указать интервал дат в будущем (01.01.2027 - 10.01.2027), нажать «Сформировать».", "Выведено предупреждение «За выбранный период исполненных заявок не обнаружено», форма не падает с ошибкой."),
        ("Сценарий 4", "Нажать кнопку «Экспорт в PDF» на командной панели отчета.", "Файл PDF успешно сохранен, форматирование бланка и колонтитулов соответствует экранному виду."),
        ("Сценарий 5", "Нажать кнопку «Экспорт в Excel» и открыть результирующий xlsx-файл.", "Файл открывается в табличном процессоре, значения в столбце сумм имеют тип Number.")
    ]
    ensure_table_rows(t8, len(steps_g) + 1)
    set_cell_text(t8.rows[0].cells[0], "Сценарий", bold=True, size_pt=9, align='center')
    set_cell_text(t8.rows[0].cells[1], "Действие оператора", bold=True, size_pt=9, align='center')
    set_cell_text(t8.rows[0].cells[2], "Ожидаемый результат", bold=True, size_pt=9, align='center')
    for ci in range(3): set_cell_background(t8.rows[0].cells[ci], "E2E8F0")
    for idx, (sc, act, exp) in enumerate(steps_g):
        row = t8.rows[idx + 1]
        set_cell_text(row.cells[0], sc, bold=True, size_pt=8.5, align='center')
        set_cell_text(row.cells[1], act, size_pt=8.5)
        set_cell_text(row.cells[2], exp, size_pt=8.5)
        if idx % 2 == 1:
            for ci in range(3): set_cell_background(row.cells[ci], "F8FAFC")

    # Paragraph 86: Test conclusion
    doc.paragraphs[86].text = (
        "Испытания проводятся на тестовом сервере автотехцентра ООО «BMW102» (srv-bmw102-test). "
        "По результатам испытаний оформляется протокол приемо-сдаточных испытаний, подписываемый разработчиком Гарифуллиным К.Р. "
        "и преподавателем кафедры АСУ УУНиТ Демченко М.С."
    )

    out_file = os.path.join(base_dir, "ШАБЛОН_ТЗ_Гарифуллин_К_Р.docx")
    doc.save(out_file)
    print(f"Saved initial docx: {out_file}")
    inject_clean_headers(out_file, "Гарифуллин К.Р.", "Отчет по количеству исполненных заявок за период")
    print("Garifullin TZ built successfully!")


def build_gaynanov_tz():
    print("--- Building TZ for Gaynanov D.I. ---")
    doc = docx.Document(tpl_tz)
    
    # Title Page
    doc.paragraphs[0].text = "ФГБОУ ВО «Уфимский университет науки и технологий»"
    doc.paragraphs[1].text = "Частное техническое задание"
    doc.paragraphs[2].text = "на разработку отчетной формы «Отчет по клиентам за период»"
    
    # Section 0 footer
    for p in doc.sections[0].first_page_footer.paragraphs:
        if '2023' in p.text:
            p.text = p.text.replace('2023', '2026')
            
    # Table 0: Approval
    t0 = doc.tables[0]
    set_cell_text(t0.rows[0].cells[0], "УТВЕРЖДАЮ:", bold=True, size_pt=9.5)
    set_cell_text(t0.rows[1].cells[0], "Заместитель генерального директора ООО «BMW102»", size_pt=9)
    set_cell_text(t0.rows[1].cells[1], "Гайнанов Д.И.", size_pt=9)
    set_cell_text(t0.rows[1].cells[2], "/ ____________ /", size_pt=9, align='center')
    set_cell_text(t0.rows[1].cells[3], "« 25 » марта 2026 г.", size_pt=9, align='right')
    set_cell_text(t0.rows[2].cells[0], "СОГЛАСОВАНО:", bold=True, size_pt=9.5)
    set_cell_text(t0.rows[3].cells[0], "Ст. преподаватель кафедры АСУ УУНиТ", size_pt=9)
    set_cell_text(t0.rows[3].cells[1], "Демченко М.С.", size_pt=9)
    set_cell_text(t0.rows[3].cells[2], "/ ____________ /", size_pt=9, align='center')
    set_cell_text(t0.rows[3].cells[3], "« 26 » марта 2026 г.", size_pt=9, align='right')
    while len(t0.rows) > 4:
        tr = t0.rows[-1]._tr
        tr.getparent().remove(tr)
        
    # Table 1: Terms
    t1 = doc.tables[1]
    terms_d = [
        ("КИС", "Корпоративная информационная система автотехцентра (КИС «BMW102-ERP»)."),
        ("ОФ", "Отчетная форма; регламентированный электронный или печатный документ аналитической отчетности."),
        ("CRM", "Система управления взаимоотношениями с клиентами (Customer Relationship Management) в составе КИС."),
        ("ЗДФ", "Запрос на доработку функциональности системы; документ, определяющий требования к расширению системы."),
        ("ДБС", "Детальный бизнес-сценарий (102BMW-BS-01 «Организация и исполнение услуг чип-тюнинга и ТО автопарка УУНиТ»)."),
        ("Контрагент", "Юридическое или физическое лицо — заказчик услуг автотехцентра (ФГБОУ ВО «УУНиТ»)."),
        ("Заявка на сервис", "Первичный документ регистрации обращения клиента на проведение диагностики и чип-тюнинга."),
        ("VIN-номер", "Уникальный 17-значный идентификационный номер транспортного средства (Vehicle Identification Number)."),
        ("СУБД", "Система управления реляционными базами данных (PostgreSQL 16), обеспечивающая хранение и целостность данных."),
        ("СКД", "Система компоновки данных платформы 1С для интерактивного построения сложных сводных отчетов.")
    ]
    ensure_table_rows(t1, len(terms_d) + 1)
    set_cell_text(t1.rows[0].cells[0], "Термин", bold=True, size_pt=9.5, align='center')
    set_cell_text(t1.rows[0].cells[1], "Описание", bold=True, size_pt=9.5, align='center')
    set_cell_background(t1.rows[0].cells[0], "E2E8F0")
    set_cell_background(t1.rows[0].cells[1], "E2E8F0")
    for idx, (term, desc) in enumerate(terms_d):
        row = t1.rows[idx + 1]
        set_cell_text(row.cells[0], term, bold=True, size_pt=8.5)
        set_cell_text(row.cells[1], desc, size_pt=8.5)
        if idx % 2 == 1:
            set_cell_background(row.cells[0], "F8FAFC")
            set_cell_background(row.cells[1], "F8FAFC")

    # Table 2: Version History
    t2 = doc.tables[2]
    v_headers = ["Автор", "Дата", "Компания / Подразделение", "Версия", "Описание изменений"]
    for ci, h in enumerate(v_headers):
        set_cell_text(t2.rows[0].cells[ci], h, bold=True, size_pt=9, align='center')
        set_cell_background(t2.rows[0].cells[ci], "E2E8F0")
    v_rows_d = [
        ("Гайнанов Д.И.", "25.03.2026", "ЭАС-514С / ООО «BMW102»", "1.0", "Разработка первой редакции частного технического задания на разработку формы «Отчет по клиентам за период»."),
        ("Демченко М.С.", "26.03.2026", "Кафедра АСУ ФГБОУ ВО «УУНиТ»", "1.1", "Рецензирование документа, согласование атрибутивного состава формы и методики испытаний."),
        ("", "", "", "", ""),
        ("", "", "", "", ""),
        ("", "", "", "", ""),
        ("", "", "", "", "")
    ]
    ensure_table_rows(t2, len(v_rows_d) + 1)
    for ri, rvals in enumerate(v_rows_d):
        row = t2.rows[ri + 1]
        for ci, val in enumerate(rvals):
            set_cell_text(row.cells[ci], val, size_pt=8.5, align='center' if ci in [1, 3] else 'left')

    # Paragraphs 22-24: Section 1
    doc.paragraphs[22].text = (
        "Настоящий документ разработан в рамках выполнения лабораторного практикума и курсового проектирования "
        "по дисциплине «Корпоративные информационные системы» в весенний семестр 2026 учебного года. "
        "Проект выполнен студентом кафедры АСУ УУНиТ учебной группы ЭАС-514С Гайнановым Данилом Ильдаровичем (вариант №102). "
        "Предметная область: «Сервисное обслуживание, диагностика ЭСУД и программная оптимизация (чип-тюнинг) автомобильного парка Заказчика»."
    )
    doc.paragraphs[23].text = (
        "Подпроцесс: «Регистрация клиентов, первичный прием и диспетчеризация заявок». Разрабатываемый модуль реализует "
        "требования Заказчика — ФГБОУ ВО «Уфимский университет науки и технологий» (УУНиТ) — к системе аналитической отчетности автотехцентра ООО «BMW102»."
    )
    doc.paragraphs[24].text = (
        "Настоящее частное техническое задание регламентирует разработку выходной аналитической формы в среде КИС «BMW102-ERP» "
        "и является неотъемлемой составной частью сквозного проектного решения 102BMW-BS-01 «Организация и исполнение услуг чип-тюнинга и ТО автопарка»."
    )

    # Paragraphs 34, 36: Section 2
    doc.paragraphs[34].text = (
        "В настоящем ЗДФ описаны требования и вариант реализации доработки в области функционального подпроцесса "
        "«Регистрация клиентов, первичный прием и диспетчеризация заявок» в части добавления функционала формирования и вывода на печать "
        "из КИС «BMW102-ERP» аналитической отчетной формы «Отчет по клиентам за период». "
        "Действие ЗДФ подчинено рамкам сквозного детального бизнес-сценария 102BMW-BS-01."
    )
    doc.paragraphs[36].text = (
        "Реализовать вывод из КИС «BMW102-ERP» регламентированной формы аналитического отчета «Отчет по клиентам за период» "
        "следующего вида (рисунок 1). Полный перечень атрибутов данного аналитического отчета представлен в Таблице 2."
    )

    # Figure 1: Report Mockup
    doc.paragraphs[38].text = ""
    doc.paragraphs[38].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[38].paragraph_format.space_before = Pt(6)
    doc.paragraphs[38].paragraph_format.space_after = Pt(4)
    doc.paragraphs[38].add_run().add_picture(os.path.join(base_dir, 'tz_gaynanov_fig1_report.png'), width=Inches(6.0))
    doc.paragraphs[39].text = "Рисунок 1 – Макет формы аналитического отчета «Отчет по клиентам за период»"
    doc.paragraphs[40].text = ""

    # Paragraph 42, 44, 45: Use Case
    doc.paragraphs[42].text = (
        "Последовательность взаимодействия оператора (сервис-консультанта или диспетчера) и системы КИС «BMW102-ERP» "
        "при формировании отчетной формы регламентируется сценарием использования, представленным в Таблице 1."
    )
    doc.paragraphs[44].text = "3.2.1 Сценарий использования «Формирование отчета по клиентам и взаиморасчетам за период»"
    doc.paragraphs[45].text = "Таблица 1 – Спецификация сценария использования ОФ"

    # Table 3: Use Case Table
    t3 = doc.tables[3]
    uc_d = [
        ("Краткое описание", "Формирование сводной аналитической ведомости по клиентам автотехцентра с детализацией по прикрепленным автомобилям, количеству обращений и сумме взаиморасчетов за период."),
        ("Действующие лица", "Сервис-консультант (Менеджер по работе с клиентами), Бухгалтер, Руководитель сервиса, КИС «BMW102-ERP»."),
        ("Триггер (Событие)", "Необходимость анализа клиентской базы, оценки повторных обращений, выявления ключевых корпоративных заказчиков (УУНиТ) и контроля выручки."),
        ("Предусловия", "1. Пользователь авторизован в системе с правами доступа к клиентскому модулю («Менеджер» или «Аналитик»).\n2. В БД зарегистрированы клиенты и наряд-заказы за отчетный интервал дат."),
        ("Основной поток событий", "1. Пользователь переходит в раздел меню «Клиенты и маркетинг -> Отчеты по контрагентам -> Отчет по клиентам за период».\n2. Система открывает сеанс-конфигуратор отчета.\n3. Пользователь задает интервал дат, выбирает тип клиента (Корпоративные/Все) и минимальный объем заказов.\n4. Пользователь нажимает «Сформировать» (F5).\n5. Система выполняет агрегирующий SQL-запрос к БД PostgreSQL и генерирует макет отчета.\n6. Система отображает готовую форму на экране с возможностью постраничного просмотра, печати или экспорта."),
        ("Альтернативный поток А1", "За указанный период клиенты с заданными параметрами не найдены. Система выводит сообщение «Данные за указанный период отсутствуют» и очищает табличную часть."),
        ("Альтернативный поток А2", "Пользователь указал некорректный интервал дат (дата окончания меньше даты начала). Система подсвечивает поле ошибки и блокирует построение до ввода корректных дат."),
        ("Постусловия", "Сформированная печатная форма отображена на экране, зафиксирована дата и время построения, отчет готов к выводу на печать или выгрузке в Excel/PDF."),
        ("Частота использования", "Еженедельно для оперативного анализа обращений, ежемесячно и ежеквартально при закрытии финансовых периодов."),
        ("Требования к интерфейсу", "Открытие конфигуратора не более 1 сек, построение сводной таблицы до 10 000 контрагентов не более 3.0 сек, поддержка горячих клавиш Ctrl+P, Esc."),
        ("Ограничения безопасности", "Контактные данные клиентов и финансовые объемы доступны только пользователям с подтвержденными правами роли «Менеджер» и «Руководитель»."),
        ("Особые требования", "Форматирование телефонных номеров должно соответствовать маске +7 (XXX) XXX-XX-XX, числовые значения сумм выводятся с разделителем тысяч.")
    ]
    ensure_table_rows(t3, len(uc_d) + 1)
    set_cell_text(t3.rows[0].cells[0], "Разделы сценария", bold=True, size_pt=9, align='center')
    set_cell_text(t3.rows[0].cells[1], "Описание разделов сценария", bold=True, size_pt=9, align='center')
    set_cell_background(t3.rows[0].cells[0], "E2E8F0")
    set_cell_background(t3.rows[0].cells[1], "E2E8F0")
    for idx, (sec_name, sec_val) in enumerate(uc_d):
        row = t3.rows[idx + 1]
        set_cell_text(row.cells[0], sec_name, bold=True, size_pt=8.5)
        set_cell_text(row.cells[1], sec_val, size_pt=8.5)
        if idx % 2 == 1:
            set_cell_background(row.cells[0], "F8FAFC")
            set_cell_background(row.cells[1], "F8FAFC")

    # Paragraphs 48-53
    doc.paragraphs[48].text = (
        "Архитектура доработки базируется на типовых механизмах платформы КИС «BMW102-ERP» с использованием СУБД PostgreSQL 16 "
        "и Системы компоновки данных (СКД). Форма отчета генерируется на основе параметризованного SQL-представления v_clients_summary_report, "
        "обеспечивающего выборку и расчет агрегатов по клиентам, привязанным автомобилям и наряд-заказам за выбранный период."
    )
    doc.paragraphs[49].text = ""
    doc.paragraphs[49].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[49].paragraph_format.space_before = Pt(6)
    doc.paragraphs[49].paragraph_format.space_after = Pt(4)
    doc.paragraphs[49].add_run().add_picture(os.path.join(base_dir, 'tz_gaynanov_fig2_dialog.png'), width=Inches(5.8))
    doc.paragraphs[50].text = "Рисунок 2 – Макет экранной формы сеанса-конфигуратора параметров отчета"
    doc.paragraphs[51].text = (
        "Сеанс-конфигуратор отчета представляет собой диалоговое окно, позволяющее пользователю задавать параметры выборки: "
        "диапазон дат обращений, категорию клиентов (корпоративные/физ. лица), порог минимальной суммы заказов, а также группировку данных."
    )
    doc.paragraphs[52].text = "Атрибутивный состав выходной формы отчета приведен в Таблице 2."
    doc.paragraphs[53].text = "Таблица 2 – Атрибутивный состав формы отчета"

    # Table 4: Report Attributes
    t4 = doc.tables[4]
    attrs_d = [
        ("1", "Код клиента", "Строка (10)", "КЛ-00102", "Уникальный идентификационный код клиента в базе КИС."),
        ("2", "Наименование / ФИО", "Строка (80)", "ФГБОУ ВО «УУНиТ» (Гайнанов Д.И.)", "Полное наименование организации-заказчика или ФИО владельца."),
        ("3", "Телефон / Контакт", "Строка (25)", "+7 (917) 102-00-01", "Контактный номер телефона уполномоченного представителя заказчика."),
        ("4", "Количество авто, шт.", "Целое число", "3", "Количество прикрепленных к договору транспортных средств."),
        ("5", "Всего заявок, шт.", "Целое число", "12", "Общее число оформленных заявок и наряд-заказов за выбранный период."),
        ("6", "Сумма оплат, руб.", "Число (12, 2)", "420 000,00", "Суммарный объем оплаченных услуг сервиса и чип-тюнинга за период.")
    ]
    ensure_table_rows(t4, len(attrs_d) + 1)
    set_cell_text(t4.rows[0].cells[0], "№", bold=True, size_pt=9, align='center')
    set_cell_text(t4.rows[0].cells[1], "Атрибут", bold=True, size_pt=9, align='center')
    set_cell_text(t4.rows[0].cells[2], "Тип данных", bold=True, size_pt=9, align='center')
    set_cell_text(t4.rows[0].cells[3], "Пример заполнения", bold=True, size_pt=9, align='center')
    set_cell_text(t4.rows[0].cells[4], "Примечание", bold=True, size_pt=9, align='center')
    for ci in range(5): set_cell_background(t4.rows[0].cells[ci], "E2E8F0")
    for idx, (num, name, dt, ex, note) in enumerate(attrs_d):
        row = t4.rows[idx + 1]
        set_cell_text(row.cells[0], num, size_pt=8.5, align='center')
        set_cell_text(row.cells[1], name, bold=True, size_pt=8.5)
        set_cell_text(row.cells[2], dt, italic=True, size_pt=8, align='center')
        set_cell_text(row.cells[3], ex, size_pt=8.5)
        set_cell_text(row.cells[4], note, size_pt=8.5)
        if idx % 2 == 1:
            for ci in range(5): set_cell_background(row.cells[ci], "F8FAFC")

    # Paragraphs 54-63
    doc.paragraphs[54].text = "Связи таблиц базы данных КИС, из которых формируется отчет, приведены на схеме (рисунок 3):"
    p_img3 = doc.paragraphs[55].insert_paragraph_before()
    p_img3.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    p_img3.paragraph_format.space_before = Pt(6)
    p_img3.paragraph_format.space_after = Pt(4)
    p_img3.add_run().add_picture(os.path.join(base_dir, 'tz_gaynanov_fig3_erd.png'), width=Inches(6.0))
    doc.paragraphs[56].text = "Рисунок 3 – Реляционная модель данных (ERD) подпроцесса учета клиентов и заказов"
    doc.paragraphs[58].text = (
        "Запуск формирования формы выполняется менеджером по работе с клиентами посредством выбора отчета через меню "
        "«Клиенты и маркетинг -> Отчеты по контрагентам -> Отчет по клиентам за период» на панели команд КИС «BMW102-ERP». "
        "При этом открывается сеанс-конфигуратор со следующим набором атрибутов (Таблица 3):"
    )
    doc.paragraphs[59].text = "Таблица 3 – Атрибуты сеанса-конфигуратора ОФ"

    # Table 5: Configurator Attributes
    t5 = doc.tables[5]
    cfg_d = [
        ("1", "Период дат (С ... По ...)", "Период дат (ДД.ММ.ГГГГ)", "Обязательный фильтр интервала анализа обращений клиентов."),
        ("2", "Тип клиента", "Список («Все», «Юр. лица», «Физ. лица»)", "Фильтрация контрагентов по организационно-правовой форме."),
        ("3", "Минимальная сумма заказов", "Число (10, 2)", "Порог фильтрации клиентов по объему выручки (по умолчанию 0,00 руб.)."),
        ("4", "Группировка данных", "Список («По типам клиентов», «Без группировки»)", "Управление иерархией вывода строк в отчете.")
    ]
    ensure_table_rows(t5, len(cfg_d) + 1)
    set_cell_text(t5.rows[0].cells[0], "№", bold=True, size_pt=9, align='center')
    set_cell_text(t5.rows[0].cells[1], "Наименование", bold=True, size_pt=9, align='center')
    set_cell_text(t5.rows[0].cells[2], "Тип", bold=True, size_pt=9, align='center')
    set_cell_text(t5.rows[0].cells[3], "Примечание", bold=True, size_pt=9, align='center')
    for ci in range(4): set_cell_background(t5.rows[0].cells[ci], "E2E8F0")
    for idx, (num, name, dt, note) in enumerate(cfg_d):
        row = t5.rows[idx + 1]
        set_cell_text(row.cells[0], num, size_pt=8.5, align='center')
        set_cell_text(row.cells[1], name, bold=True, size_pt=8.5)
        set_cell_text(row.cells[2], dt, size_pt=8.5, align='center')
        set_cell_text(row.cells[3], note, size_pt=8.5)
        if idx % 2 == 1:
            for ci in range(4): set_cell_background(row.cells[ci], "F8FAFC")

    # Paragraphs 61, 63, 64
    doc.paragraphs[61].text = (
        "На распечатанном бланке отчета автоматически выводится заголовок формы отчета «Отчет по клиентам за период», "
        "наименование автотехцентра ООО «BMW102», интервал дат выборки, текущая дата печати и ФИО оператора."
    )
    doc.paragraphs[63].text = "Источники атрибутов формы с логикой их формирования указаны в Таблице 4."
    doc.paragraphs[64].text = "Таблица 4 – Источники атрибутов формы с логикой формирования"

    # Table 6: Sources
    t6 = doc.tables[6]
    sources_d = [
        ("1", "Код клиента", "КЛ-00102", "Строка (10)", "Таблица clients, поле client_code", "Уникальный идентификатор контрагента в КИС."),
        ("2", "Наименование / ФИО", "ФГБОУ ВО «УУНиТ»", "Строка (80)", "Таблица clients, поле name", "Официальное наименование юридического лица или ФИО физ. лица."),
        ("3", "Телефон / Контакт", "+7 (917) 102-00-01", "Строка (25)", "Таблица clients, поле phone", "Основной телефонный номер для оповещения и согласования."),
        ("4", "Количество авто", "3", "Целое число", "Таблица cars, поле car_id", "Агрегатная функция COUNT(cars.car_id) с группировкой по client_id."),
        ("5", "Всего заявок", "12", "Целое число", "Таблица service_requests, request_id", "Агрегатная функция COUNT(request_id) за выбранный диапазон дат."),
        ("6", "Сумма оплат, руб.", "420 000,00", "Число (12, 2)", "Таблица payments, поле amount", "Суммирование проведенных оплат: SUM(payments.amount).")
    ]
    ensure_table_rows(t6, len(sources_d) + 1)
    set_cell_text(t6.rows[0].cells[0], "№", bold=True, size_pt=8.5, align='center')
    set_cell_text(t6.rows[0].cells[1], "Наименование", bold=True, size_pt=8.5, align='center')
    set_cell_text(t6.rows[0].cells[2], "Пример", bold=True, size_pt=8.5, align='center')
    set_cell_text(t6.rows[0].cells[3], "Тип данных", bold=True, size_pt=8.5, align='center')
    set_cell_text(t6.rows[0].cells[4], "Источник (ERD)", bold=True, size_pt=8.5, align='center')
    set_cell_text(t6.rows[0].cells[5], "Реализация / SQL", bold=True, size_pt=8.5, align='center')
    for ci in range(6): set_cell_background(t6.rows[0].cells[ci], "E2E8F0")
    for idx, (num, name, ex, dt, src, real) in enumerate(sources_d):
        row = t6.rows[idx + 1]
        set_cell_text(row.cells[0], num, size_pt=8, align='center')
        set_cell_text(row.cells[1], name, bold=True, size_pt=8)
        set_cell_text(row.cells[2], ex, size_pt=8)
        set_cell_text(row.cells[3], dt, italic=True, size_pt=7.5, align='center')
        set_cell_text(row.cells[4], src, size_pt=8)
        set_cell_text(row.cells[5], real, size_pt=8)
        if idx % 2 == 1:
            for ci in range(6): set_cell_background(row.cells[ci], "F8FAFC")

    # Paragraphs 68, 71, 73, 74, 77, 80
    doc.paragraphs[68].text = (
        "Ограничения и допущения: формирование отчета осуществляется по клиентам, имеющим хотя бы одну зарегистрированную заявку. "
        "Конфиденциальные персональные данные физических лиц маскируются в соответствии с требованиями 152-ФЗ. "
        "При отсутствии данных за указанный интервал времени система выводит предупреждающее сообщение."
    )
    doc.paragraphs[71].text = (
        "Программа и методика приемочных испытаний регламентируют условия, порядок и сценарии проверки разработанной формы отчета "
        "на соответствие требованиям настоящего ТЗ перед вводом в промышленную эксплуатацию."
    )
    doc.paragraphs[73].text = (
        "Для проведения испытаний требуется функционирующий тестовый контур КИС «BMW102-ERP» на базе PostgreSQL 16 "
        "с наполненными справочниками контрагентов, автомобилей и тестовым пулом оформленных заявок (не менее 20 записей)."
    )
    doc.paragraphs[74].text = "Предварительных настроек клиентских рабочих мест, сверх штатных средств КИС «BMW102-ERP» и ОС Windows, не требуется."
    doc.paragraphs[77].text = "Программа испытаний включает перечень контрольных сценариев проверки (Таблица 5):"
    doc.paragraphs[80].text = "Таблица 5 – Программа приемочных испытаний"

    # Table 7: Test Scenarios
    t7 = doc.tables[7]
    tests_d = [
        ("1", "Формирование отчета за стандартный период (текущий квартал)", "ТР-01", "Проверка корректности извлечения клиентских данных и агрегации."),
        ("2", "Фильтрация выборки по типу контрагента («Юридические лица»)", "ТР-02", "Проверка корректности наложения отбора по категории клиента."),
        ("3", "Формирование отчета с ограничением по минимальной сумме заказов", "ТР-03", "Проверка корректности фильтрации по финансовому порогу выручки."),
        ("4", "Экспорт сформированного отчета в форматы PDF и MS Excel", "ТР-04", "Проверка сохранения визуального оформления и числовых форматов ячеек.")
    ]
    ensure_table_rows(t7, len(tests_d) + 1)
    set_cell_text(t7.rows[0].cells[0], "№ п/п", bold=True, size_pt=9, align='center')
    set_cell_text(t7.rows[0].cells[1], "Наименование сценария", bold=True, size_pt=9, align='center')
    set_cell_text(t7.rows[0].cells[2], "Номер требования", bold=True, size_pt=9, align='center')
    set_cell_text(t7.rows[0].cells[3], "Примечание", bold=True, size_pt=9, align='center')
    for ci in range(4): set_cell_background(t7.rows[0].cells[ci], "E2E8F0")
    for idx, (num, name, req, note) in enumerate(tests_d):
        row = t7.rows[idx + 1]
        set_cell_text(row.cells[0], num, size_pt=8.5, align='center')
        set_cell_text(row.cells[1], name, bold=True, size_pt=8.5)
        set_cell_text(row.cells[2], req, size_pt=8.5, align='center')
        set_cell_text(row.cells[3], note, size_pt=8.5)
        if idx % 2 == 1:
            for ci in range(4): set_cell_background(row.cells[ci], "F8FAFC")

    # Paragraphs 82, 84
    doc.paragraphs[82].text = "Пошаговая методика проведения испытаний приведена в Таблице 6:"
    doc.paragraphs[84].text = "Таблица 6 – Методика проведения испытаний"

    # Table 8: Test Steps
    t8 = doc.tables[8]
    steps_d = [
        ("Сценарий 1", "Запустить конфигуратор, указать период 01.01.2026 - 31.03.2026, нажать «Сформировать».", "Отчет сформирован, выведены все зарегистрированные клиенты за квартал, итоговые суммы подсчитаны."),
        ("Сценарий 2", "Выбрать в поле «Тип клиента» значение «Юридические лица», нажать «Сформировать».", "В таблице отображаются только корпоративные заказчики (включая ФГБОУ ВО «УУНиТ»), физ. лица скрыты."),
        ("Сценарий 3", "Указать порог «Минимальная сумма заказов» = 100 000,00 руб.", "В отчете отображаются только клиенты с суммарным объемом заказов не менее 100 000 руб."),
        ("Сценарий 4", "Нажать кнопку «Экспорт в PDF» на панели управления отчетом.", "Файл PDF успешно сохранен, форматирование бланка и колонтитулов соответствует экранному виду."),
        ("Сценарий 5", "Нажать кнопку «Экспорт в Excel» и открыть результирующий xlsx-файл.", "Файл корректно открывается в табличном редакторе, столбцы сумм и количеств имеют числовой формат.")
    ]
    ensure_table_rows(t8, len(steps_d) + 1)
    set_cell_text(t8.rows[0].cells[0], "Сценарий", bold=True, size_pt=9, align='center')
    set_cell_text(t8.rows[0].cells[1], "Действие оператора", bold=True, size_pt=9, align='center')
    set_cell_text(t8.rows[0].cells[2], "Ожидаемый результат", bold=True, size_pt=9, align='center')
    for ci in range(3): set_cell_background(t8.rows[0].cells[ci], "E2E8F0")
    for idx, (sc, act, exp) in enumerate(steps_d):
        row = t8.rows[idx + 1]
        set_cell_text(row.cells[0], sc, bold=True, size_pt=8.5, align='center')
        set_cell_text(row.cells[1], act, size_pt=8.5)
        set_cell_text(row.cells[2], exp, size_pt=8.5)
        if idx % 2 == 1:
            for ci in range(3): set_cell_background(row.cells[ci], "F8FAFC")

    # Paragraph 86: Test conclusion
    doc.paragraphs[86].text = (
        "Испытания проводятся на тестовом сервере автотехцентра ООО «BMW102» (srv-bmw102-test). "
        "По результатам испытаний оформляется протокол приемо-сдаточных испытаний, подписываемый разработчиком Гайнановым Д.И. "
        "и преподавателем кафедры АСУ УУНиТ Демченко М.С."
    )

    out_file = os.path.join(base_dir, "ШАБЛОН_ТЗ_Гайнанов_Д_И.docx")
    doc.save(out_file)
    print(f"Saved initial docx: {out_file}")
    inject_clean_headers(out_file, "Гайнанов Д.И.", "Отчет по клиентам за период")
    print("Gaynanov TZ built successfully!")

if __name__ == '__main__':
    build_garifullin_tz()
    build_gaynanov_tz()
    print("=== All TZ documents built successfully! ===")

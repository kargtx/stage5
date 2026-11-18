import os
import sys
import zipfile
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

sys.stdout.reconfigure(encoding='utf-8')

print("=== Generating Technical Specification (TZ) Documents for Garifullin and Gaynanov ===")

base_dir = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС'
tpl_tz = os.path.join(base_dir, 'ШАБЛОН ТЗ.docx')

# Helper functions
def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_background(cell, fill_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_text(cell, text, bold=False, italic=False, size_pt=9.0, align='left', font_name='Arial'):
    cell.text = ''
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
    return p

def insert_picture_after(paragraph, img_path, width_inches=6.0):
    p = paragraph._p.getparent().add_paragraph()
    # Move p to be right after paragraph
    paragraph._p.addnext(p._p)
    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run()
    run.add_picture(img_path, width=Inches(width_inches))
    return p

# Logos
uunit_logo_path = os.path.join(base_dir, 'prepared_uunit_header.png')
exec_logo_path = os.path.join(base_dir, 'prepared_exec_header.png')
with open(uunit_logo_path, 'rb') as f: uunit_bytes = f.read()
with open(exec_logo_path, 'rb') as f: exec_bytes = f.read()

# Build TZ for Garifullin
def build_garifullin_tz():
    doc = docx.Document(tpl_tz)
    
    # Title Page
    doc.paragraphs[0].text = "ФГБОУ ВО «Уфимский университет науки и технологий»"
    doc.paragraphs[1].text = "Частное техническое задание"
    doc.paragraphs[2].text = "на разработку отчетной формы «Отчет по количеству исполненных заявок за период»"
    
    # Table 0: Approval
    t0 = doc.tables[0]
    t0_data = [
        ("Утверждено:", "", "", ""),
        ("Начальник транспортного отдела ФГБОУ ВО «УУНиТ»", "", "А.В. Смирнов", "«___» марта 2026 г."),
        ("Согласовано:", "", "", ""),
        ("Главный инженер ООО «BMW102»", "", "Р.Ф. Ахметов", "«___» марта 2026 г."),
        ("Ведущий разработчик (студент гр. ЭАС-514С)", "", "К.Р. Гарифуллин", "«___» марта 2026 г."),
        ("Преподаватель кафедры АСУ УУНиТ", "", "М.С. Демченко", "«___» марта 2026 г."),
        ("", "", "", ""),
        ("", "", "", ""),
        ("", "", "", "")
    ]
    for ri, rvals in enumerate(t0_data):
        row = t0.rows[ri]
        for ci, val in enumerate(rvals):
            is_bold = (ri in [0, 2] or ci == 0)
            align = 'center' if ci == 2 else ('right' if ci == 3 else 'left')
            set_cell_text(row.cells[ci], val, bold=is_bold, size_pt=9, align=align)
            
    # Table 1: Terms
    t1 = doc.tables[1]
    terms_g = [
        ("КИС", "Корпоративная информационная система; программный комплекс автоматизации автотехцентра (КИС «BMW102-ERP»)."),
        ("ОФ", "Отчетная форма; регламентированный выходной электронный или печатный документ аналитической отчетности."),
        ("ЗДФ", "Запрос на доработку функциональности системы; документ, определяющий требования к расширению функциональности."),
        ("ДБС", "Детальный бизнес-сценарий (102BMW-BS-01 «Организация и исполнение услуг чип-тюнинга и ТО автопарка УУНиТ»)."),
        ("Наряд-заказ", "Первичный документ цеха, регламентирующий согласованный перечень работ, калькуляцию и назначенного мастера."),
        ("Акт выполненных работ", "Закрывающий документ, фиксирующий надлежащее исполнение наряд-заказа и являющийся основанием для расчетов."),
        ("ЭБУ (ECU)", "Электронный блок управления двигателем внутреннего сгорания (Engine Control Unit)."),
        ("Чип-тюнинг (Stage 1)", "Программная модификация калибровочных таблиц ЭБУ для увеличения мощности и крутящего момента ДВС."),
        ("СУБД", "Система управления реляционными базами данных (PostgreSQL), обеспечивающая транзакционную целостность."),
        ("СКД", "Система компоновки данных платформы 1С для интерактивного построения сложных сводных отчетов.")
    ]
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
    for ri, rvals in enumerate(v_rows):
        row = t2.rows[ri + 1]
        for ci, val in enumerate(rvals):
            set_cell_text(row.cells[ci], val, size_pt=8.5, align='center' if ci in [1, 3] else 'left')

    # Body paragraphs
    doc.paragraphs[22].text = (
        "Настоящий документ разработан в рамках выполнения работ по дисциплине «Корпоративные информационные системы» "
        "в весенний семестр 2026 года. Вариант № 102. Предметная область: «Автоматизация процессов чип-тюнинга и сервисного "
        "обслуживания автопарка УУНиТ в автотехцентре BMW102»."
    )
    doc.paragraphs[23].text = (
        "Подпроцесс «Исполнение наряд-заказа и закрытие заявки» и реализует требования Заказчика (ФГБОУ ВО «УУНиТ») "
        "и руководства автотехцентра ООО «BMW102» к системе оперативной и аналитической отчётности производственного цеха."
    )
    doc.paragraphs[24].text = (
        "В документе описаны бизнес-требования к доработке, сценарии использования, варианты технической реализации, "
        "источники данных в реляционной модели (ERD 3NF) и методы тестирования доработки при ее приемо-сдаточных испытаниях."
    )
    
    doc.paragraphs[34].text = (
        "В настоящем ЗДФ описаны требования и вариант реализации доработки в области бизнес-процессов исполнения заказов "
        "в части добавления функционала формирования на экране и вывода на печать из КИС аналитической отчётной формы "
        "«Отчет по количеству исполненных заявок за период». Действие ЗДФ подчинено рамкам детального бизнес-сценария "
        "102BMW-BS-01 «Организация и исполнение услуг чип-тюнинга и технического обслуживания автопарка»."
    )
    
    doc.paragraphs[36].text = (
        "Реализовать вывод из ИС «BMW102-ERP» формы аналитического отчёта «Отчет по количеству исполненных заявок за период» "
        "следующего вида (рисунок 1 – Форма аналитического отчета). Полный перечень атрибутов данного аналитического отчета "
        "представлен в Таблица 2 Атрибутивный состав формы."
    )
    doc.paragraphs[38].text = "" # remove placeholder
    doc.paragraphs[39].text = "Рисунок 1 – Макет формы аналитического отчета «Отчет по количеству исполненных заявок за период»"
    
    # Replace/Insert figure 1
    # P40 currently has the old image. Let's clear its runs and add our image!
    doc.paragraphs[40].text = ""
    run_fig1 = doc.paragraphs[40].add_run()
    run_fig1.add_picture(os.path.join(base_dir, 'tz_garifullin_fig1_report.png'), width=Inches(6.2))
    doc.paragraphs[40].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    
    doc.paragraphs[42].text = (
        "В данном подразделе описывается регламент взаимодействия пользователей с модулем построения отчета. "
        "Мастер цеха или руководитель формирует выборку исполненных заявок за отчетный месяц, контролирует объемы "
        "выработки, соблюдение технологических регламентов и итоговую сумму выручки."
    )
    doc.paragraphs[44].text = "3.2.1 Сценарий использования «Формирование отчета по исполненным заявкам за период»"
    doc.paragraphs[45].text = "Таблица 1 – Спецификация сценария использования ОФ"
    
    # Table 3: Use Case
    t3 = doc.tables[3]
    uc_g = [
        ("Наименование сценария", "Формирование и анализ отчета по исполненным заявкам за отчетный период"),
        ("Действующие лица", "Мастер цеха / Инженер-калибровщик (Гарифуллин К.Р.), Руководитель сервиса (Ахметов Р.Ф.)"),
        ("Описание сценария использования", "Пользователь запускает сеанс ОФ, задает интервал дат, применяет фильтры по мастеру и статусу, выполняет расчет показателей и производит анализ или печать"),
        ("Предварительные условия", "Пользователь успешно авторизован в КИС «BMW102-ERP» с правами «Мастер цеха», в БД имеются проведенные документы «Наряд-заказ» и «Акт выполненных работ»"),
        ("Выходные условия", "На экране отображена сформированная таблица исполненных заявок со сводными финансовыми и количественными показателями, доступен экспорт в XLSX и PDF/A"),
        ("Нормальное направление развития (happy flow)", "1. Открытие пункта меню «Отчеты цеха» -> «Исполненные заявки за период»;\n2. Автозаполнение периода текущим месяцем (01.03.2026 - 31.03.2026);\n3. Выбор мастера (Гарифуллин К.Р.) и статуса «Исполнен»;\n4. Нажатие кнопки «Сформировать»;\n5. Выполнение запроса к СУБД PostgreSQL, агрегация данных;\n6. Вывод печатной формы на экран;\n7. Отправка отчета на печать или экспорт в XLSX."),
        ("Альтернативное направление (unhappy flow)", "За выбранный интервал дат не найдено ни одной исполненной заявки. Система выводит информационное предупреждение «Данные за указанный период не обнаружены» без прерывания сеанса."),
        ("Исключения", "Сбой сетевого подключения к серверу баз данных. Система информирует пользователя «Ошибка подключения к БД. Повторите попытку позже» и регистрирует ошибку в системном журнале."),
        ("Частота использования сценария", "Ежедневно (в конце рабочей смены) и ежемесячно (для расчета сдельной оплаты труда и закрытия отчетного периода)."),
        ("Бизнес-правила", "В отчет включаются исключительно заявки, по которым полностью завершены ходовые испытания, внесен ТЕХНИЧЕСКИЙ КОММЕНТАРИЙ мастера и оформлен двусторонний Акт выполненных работ."),
        ("Специальные требования", "Время формирования отчета не должно превышать 3.0 секунд при объеме выборки до 50 000 записей.")
    ]
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

    doc.paragraphs[48].text = (
        "Реализация доработки выполняется в среде КИС «BMW102-ERP» на базе технологической платформы "
        "1С:Предприятие 8.3 КОРП с использованием встроенного языка и механизма Системы компоновки данных (СКД). "
        "Для интеграции с внешними веб-клиентами реализован REST API сервис выгрузки сводных данных."
    )
    doc.paragraphs[50].text = "Рисунок 2 – Макет экранной формы сеанса-конфигуратора параметров отчета"
    
    # Insert Figure 2 after P50
    insert_picture_after(doc.paragraphs[50], os.path.join(base_dir, 'tz_garifullin_fig2_dialog.png'), width_inches=5.8)
    
    doc.paragraphs[51].text = (
        "Отчетная форма представляет собой интерактивную таблицу, сгруппированную по мастерам-калибровщикам "
        "и датам закрытия нарядов, с выводом итоговых количественных и финансовых сумм за выбранный период."
    )
    doc.paragraphs[52].text = "Атрибутивный состав формы приведён в Таблице 2."
    doc.paragraphs[53].text = "Таблица 2 – Атрибутивный состав формы отчета"
    
    # Table 4: Attributes of Report Form
    t4 = doc.tables[4]
    attrs_g = [
        ("1", "Номер п/п", "INT", "1, 2, 3...", "Порядковый номер строки в отчете"),
        ("2", "Номер заявки", "VARCHAR(20)", "SR-2026-00102", "Идентификатор первичной заявки клиента"),
        ("3", "Номер наряд-заказа", "VARCHAR(20)", "WO-2026-00441", "Номер производственного наряд-заказа"),
        ("4", "Дата закрытия", "DATE", "25.03.2026", "Дата проведения выходного контроля и подписания Акта"),
        ("5", "Заказчик / Контрагент", "VARCHAR(100)", "ФГБОУ ВО «УУНиТ»", "Наименование юридического лица или ФИО клиента"),
        ("6", "Автомобиль (VIN / Гос. номер)", "VARCHAR(100)", "BMW 530d (У 102 УУ 102)", "Идентификационные параметры обслуженного ТС")
    ]
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

    doc.paragraphs[54].text = "Связи таблиц базы данных КИС, из которых формируется отчет, приведены на схеме:"
    doc.paragraphs[55].text = "Рисунок 3 – Реляционная модель данных (ERD) подпроцесса исполнения заявок"
    
    # Insert Figure 3 after P55
    insert_picture_after(doc.paragraphs[55], os.path.join(base_dir, 'tz_garifullin_fig3_erd.png'), width_inches=6.0)

    doc.paragraphs[57].text = (
        "Запуск формирования формы выполняется пользователем посредством выбора пункта меню "
        "«Цех / Отчеты / Исполненные заявки за период» на главной панели команд КИС «BMW102-ERP». "
        "При этом открывается диалоговый сеанс-конфигуратор со следующим набором управляющих атрибутов:"
    )
    doc.paragraphs[58].text = "Таблица 3 – Атрибуты сеанса-конфигуратора ОФ"
    
    # Table 5: Configurator Attributes
    t5 = doc.tables[5]
    cfg_g = [
        ("1", "Период (С .. ПО)", "Дата (с .. по)", "Задаёт интервал выборки закрытых нарядов (по умолчанию текущий месяц)"),
        ("2", "Мастер-калибровщик", "Ссылка на справочник", "Отбор по конкретному сотруднику цеха (или «Все мастера»)"),
        ("3", "Выгрузка в xls/pdf", "Флаг (булево)", "Если флаг активен, выполняется экспорт в файл, иначе вывод на экран/печать")
    ]
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

    doc.paragraphs[60].text = (
        "На распечатанном бланке отчета автоматически выводится заголовок формы отчета "
        "«Отчет по количеству исполненных заявок за период», интервал дат выборки, ФИО мастера и текущая дата печати."
    )
    doc.paragraphs[62].text = "Источники атрибутов формы с логикой их формирования указаны в Таблице 4."
    doc.paragraphs[63].text = "Таблица 4 – Источники атрибутов формы с логикой формирования"
    
    # Table 6: Sources
    t6 = doc.tables[6]
    # Ensure Table 6 has enough rows
    sources_g = [
        ("1", "Номер заявки", "SR-2026-00102", "VARCHAR", "Service_Request.id_request", "Прямая выборка поля id_request"),
        ("2", "Номер наряда", "WO-2026-00441", "VARCHAR", "Work_Order.id_order", "Прямая выборка поля id_order"),
        ("3", "Дата закрытия", "25.03.2026", "DATE", "Work_Completion_Act.act_date", "DATE(Work_Completion_Act.act_date)"),
        ("4", "Заказчик", "ФГБОУ ВО «УУНиТ»", "VARCHAR", "Client.client_name", "JOIN Vehicle -> Client.client_name"),
        ("5", "Автомобиль", "BMW 530d G30", "VARCHAR", "Vehicle.brand, model", "CONCAT(Vehicle.brand, ' ', Vehicle.model)"),
        ("6", "Пакет тюнинга", "Stage 1 + Euro-2", "VARCHAR", "Service_Request.service_package", "Выборка service_package"),
        ("7", "Сумма по Акту", "41 650,00", "DECIMAL", "Work_Completion_Act.total_amount", "SUM(Work_Completion_Act.total_amount)")
    ]
    while len(t6.rows) - 1 < len(sources_g):
        t6.add_row()
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

    doc.paragraphs[67].text = (
        "Дополнительная логика не предусмотрена. Доступ к формированию отчета ограничивается на уровне ролей (RBAC): "
        "полный доступ предоставляется ролям «Мастер цеха», «Руководитель сервиса» и «Администратор КИС». "
        "Запрещено редактирование сформированных данных отчета на печатной форме."
    )
    
    # Section 4: Testing
    doc.paragraphs[73].text = (
        "Предварительных настроек, сверх необходимых для штатного функционирования КИС «BMW102-ERP» "
        "и СУБД PostgreSQL, не требуется. В тестовой базе должны быть развернуты контрольные транзакции наряд-заказов."
    )
    doc.paragraphs[79].text = "Таблица 5 – Программа приемочных испытаний"
    
    # Table 7: Test Scenarios
    t7 = doc.tables[7]
    tests_g = [
        ("1", "Формирование отчета за стандартный интервал дат (текущий месяц)", "ТР-01", "Проверка корректности выборки и подсчета итогов"),
        ("2", "Фильтрация по конкретному мастеру (Гарифуллин К.Р.)", "ТР-02", "Проверка изоляции данных других мастеров цеха"),
        ("3", "Построение отчета за период без исполненных заявок", "ТР-03", "Проверка обработки пустой выборки без аварийного сбоя"),
        ("4", "Экспорт сформированного отчета в форматы MS Excel (.xlsx) и PDF", "ТР-04", "Проверка совпадения формул и печатной верстки")
    ]
    while len(t7.rows) - 1 < len(tests_g):
        t7.add_row()
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

    doc.paragraphs[83].text = "Таблица 6 – Методика проведения испытаний"
    
    # Table 8: Test Steps
    t8 = doc.tables[8]
    steps_g = [
        ("Сценарий 1", "Шаг 1. Открыть пункт «Отчеты цеха» -> «Исполненные заявки». Задать период 01.03.2026 - 31.03.2026. Нажать «Сформировать».", "Форма строится не более 3.0 с. В таблице отображено 46 исполненных заявок, общая сумма 1 842 500,00 руб."),
        ("Сценарий 2", "Шаг 2. В поле «Мастер-калибровщик» выбрать «Гарифуллин К.Р.», нажать «Сформировать».", "В отчете отображаются только наряды, закрепленные за Гарифуллиным К.Р., итоговая сумма пересчитана."),
        ("Сценарий 3", "Шаг 3. Задать будущий период (01.01.2027 - 31.01.2027), нажать «Сформировать».", "Выводится сообщение: «За выбранный период данные отсутствуют». Таблица пустая, ошибок нет."),
        ("Сценарий 4", "Шаг 4. Нажать кнопку «Экспорт в XLSX». Открыть сохраненный файл.", "Файл открывается в MS Excel, итоговые формулы суммирования функционируют, форматирование сохранено."),
        ("Сценарий 5", "Шаг 5. Нажать кнопку «Печать». Выполнить отправку на принтер.", "Сформирован корректный печатный лист А4 альбомной ориентации, колонтитулы и подписи на месте.")
    ]
    set_cell_text(t8.rows[0].cells[0], "Сценарий", bold=True, size_pt=9, align='center')
    set_cell_text(t8.rows[0].cells[1], "Действие оператора", bold=True, size_pt=9, align='center')
    set_cell_text(t8.rows[0].cells[2], "Ожидаемый результат", bold=True, size_pt=9, align='center')
    for ci in range(3): set_cell_background(t8.rows[0].cells[ci], "E2E8F0")
    for idx, (sc, act, exp) in enumerate(steps_g):
        row = t8.rows[idx + 1]
        set_cell_text(row.cells[0], sc, bold=True, size_pt=8.5, align='center')
        set_cell_text(row.cells[1], act, size_pt=8.5)
        set_cell_text(row.cells[2], exp, size_pt=8.5)

    doc.paragraphs[85].text = (
        "Испытания проводятся на тестовом сервере Исполнителя с использованием эталонной базы данных КИС «BMW102-ERP». "
        "Условия проведения: рабочая станция под управлением ОС Windows 10/11, браузер Google Chrome или тонкий клиент 1С. "
        "По результатам испытаний составляется Протокол приемо-сдаточных испытаний, подписываемый представителями Заказчика и Исполнителя."
    )
    
    # Save temp and repack
    temp_p = os.path.join(base_dir, 'temp_tz_garifullin.docx')
    out_final = os.path.join(base_dir, 'ШАБЛОН_ТЗ_Гарифуллин_К_Р.docx')
    doc.save(temp_p)
    repack_tz(temp_p, out_final)
    if os.path.exists(temp_p): os.remove(temp_p)
    print("Saved Garifullin TZ to:", out_final)

# Build TZ for Gaynanov
def build_gaynanov_tz():
    doc = docx.Document(tpl_tz)
    
    # Title Page
    doc.paragraphs[0].text = "ФГБОУ ВО «Уфимский университет науки и технологий»"
    doc.paragraphs[1].text = "Частное техническое задание"
    doc.paragraphs[2].text = "на разработку отчетной формы «Отчет по клиентам за период»"
    
    # Table 0: Approval
    t0 = doc.tables[0]
    t0_data = [
        ("Утверждено:", "", "", ""),
        ("Начальник транспортного отдела ФГБОУ ВО «УУНиТ»", "", "А.В. Смирнов", "«___» марта 2026 г."),
        ("Согласовано:", "", "", ""),
        ("Главный инженер ООО «BMW102»", "", "Р.Ф. Ахметов", "«___» марта 2026 г."),
        ("Ведущий системный аналитик (студент гр. ЭАС-514С)", "", "Д.И. Гайнанов", "«___» марта 2026 г."),
        ("Преподаватель кафедры АСУ УУНиТ", "", "М.С. Демченко", "«___» марта 2026 г."),
        ("", "", "", ""),
        ("", "", "", ""),
        ("", "", "", "")
    ]
    for ri, rvals in enumerate(t0_data):
        row = t0.rows[ri]
        for ci, val in enumerate(rvals):
            is_bold = (ri in [0, 2] or ci == 0)
            align = 'center' if ci == 2 else ('right' if ci == 3 else 'left')
            set_cell_text(row.cells[ci], val, bold=is_bold, size_pt=9, align=align)
            
    # Table 1: Terms
    t1 = doc.tables[1]
    terms_d = [
        ("КИС", "Корпоративная информационная система; программный комплекс автоматизации автотехцентра (КИС «BMW102-ERP»)."),
        ("ОФ", "Отчетная форма; регламентированный выходной электронный или печатный документ аналитической отчетности."),
        ("ЗДФ", "Запрос на доработку функциональности системы; документ, определяющий требования к расширению функциональности."),
        ("ДБС", "Детальный бизнес-сценарий (102BMW-BS-01 «Организация и исполнение услуг чип-тюнинга и ТО автопарка УУНиТ»)."),
        ("CRM", "Customer Relationship Management; модуль управления клиентской базой, договорами и автопарками контрагентов."),
        ("Контрагент", "Юридическое или физическое лицо — Заказчик услуг сервиса (ФГБОУ ВО «УУНиТ» и др.)."),
        ("Наряд-заказ", "Первичный производственный документ, фиксирующий перечень работ, калькуляцию и автомобиль клиента."),
        ("Дебиторская задолженность", "Сумма задолженности контрагента по фактически выполненным, но не оплаченным наряд-заказам."),
        ("СУБД", "Система управления реляционными базами данных (PostgreSQL), обеспечивающая надежное хранение информации."),
        ("СКД", "Система компоновки данных платформы 1С для интерактивного построения сложных аналитических отчетов.")
    ]
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
    v_rows = [
        ("Гайнанов Д.И.", "25.03.2026", "ЭАС-514С / ООО «BMW102»", "1.0", "Разработка первой редакции частного технического задания на разработку формы «Отчет по клиентам за период»."),
        ("Демченко М.С.", "26.03.2026", "Кафедра АСУ ФГБОУ ВО «УУНиТ»", "1.1", "Рецензирование документа, верификация соответствия СТП, согласование состава показателей и методики испытаний."),
        ("", "", "", "", ""),
        ("", "", "", "", ""),
        ("", "", "", "", ""),
        ("", "", "", "", "")
    ]
    for ri, rvals in enumerate(v_rows):
        row = t2.rows[ri + 1]
        for ci, val in enumerate(rvals):
            set_cell_text(row.cells[ci], val, size_pt=8.5, align='center' if ci in [1, 3] else 'left')

    # Body paragraphs
    doc.paragraphs[22].text = (
        "Настоящий документ разработан в рамках выполнения работ по дисциплине «Корпоративные информационные системы» "
        "в весенний семестр 2026 года. Вариант № 102. Предметная область: «Автоматизация процессов чип-тюнинга и сервисного "
        "обслуживания автопарка УУНиТ в автотехцентре BMW102»."
    )
    doc.paragraphs[23].text = (
        "Подпроцесс «Оформление и регистрация заявок клиентов» и реализует требования Заказчика (ФГБОУ ВО «УУНиТ») "
        "и коммерческого отдела автотехцентра ООО «BMW102» к системе аналитической отчётности клиентской базы."
    )
    doc.paragraphs[24].text = (
        "В документе описаны бизнес-требования к доработке, сценарии использования, варианты технической реализации, "
        "структура сущностей в реляционной модели данных (ERD 3NF) и методы тестирования доработки при ее приемочных испытаниях."
    )
    
    doc.paragraphs[34].text = (
        "В настоящем ЗДФ описаны требования и вариант реализации доработки в области бизнес-процессов клиентского обслуживания "
        "в части добавления функционала формирования и вывода на печать из КИС аналитической отчётной формы "
        "«Отчет по клиентам за период». Действие ЗДФ подчинено рамкам детального бизнес-сценария "
        "102BMW-BS-01 «Организация и исполнение услуг чип-тюнинга и технического обслуживания автопарка»."
    )
    
    doc.paragraphs[36].text = (
        "Реализовать вывод из ИС «BMW102-ERP» формы аналитического отчёта «Отчет по клиентам за период» "
        "следующего вида (рисунок 1 – Форма аналитического отчета). Полный перечень атрибутов данного аналитического отчета "
        "представлен в Таблица 2 Атрибутивный состав формы."
    )
    doc.paragraphs[38].text = "" # remove placeholder
    doc.paragraphs[39].text = "Рисунок 1 – Макет формы аналитического отчета «Отчет по клиентам за период»"
    
    # Replace/Insert figure 1
    doc.paragraphs[40].text = ""
    run_fig1 = doc.paragraphs[40].add_run()
    run_fig1.add_picture(os.path.join(base_dir, 'tz_gaynanov_fig1_report.png'), width=Inches(6.2))
    doc.paragraphs[40].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    
    doc.paragraphs[42].text = (
        "В данном подразделе описывается последовательность взаимодействия менеджера CRM или финансового аналитика "
        "с модулем формирования отчета. Пользователь задает интервал анализа (квартал, полугодие, год), выбирает категорию "
        "клиентов (корпоративные автопарки или физ. лица) и выполняет анализ активности, выручки и дебиторской задолженности."
    )
    doc.paragraphs[44].text = "3.2.1 Сценарий использования «Формирование отчета по клиентам и взаиморасчетам за период»"
    doc.paragraphs[45].text = "Таблица 1 – Спецификация сценария использования ОФ"
    
    # Table 3: Use Case
    t3 = doc.tables[3]
    uc_d = [
        ("Наименование сценария", "Формирование и анализ сводного отчета по клиентам и взаиморасчетам за период"),
        ("Действующие лица", "Менеджер по работе с клиентами / Системный аналитик (Гайнанов Д.И.), Бухгалтер (Смирнов А.В.)"),
        ("Описание сценария использования", "Пользователь запускает сеанс ОФ, задает отчетный период, выбирает категорию контрагентов, производит расчет выручки, скидок и задолженностей и выводит отчет на экран или печать"),
        ("Предварительные условия", "Пользователь авторизован в КИС «BMW102-ERP» с правами «Менеджер CRM» / «Аналитик», в БД имеются зарегистрированные контрагенты, автомобили и наряд-заказы"),
        ("Выходные условия", "На экране отображена сводная ведомость контрагентов с детализацией по автопарку, начислениям и оплатам, доступен экспорт в XLSX и PDF/A"),
        ("Нормальное направление развития (happy flow)", "1. Открытие пункта меню «CRM / Отчеты» -> «Отчет по клиентам за период»;\n2. Задание отчетного интервала (01.01.2026 - 31.03.2026);\n3. Выбор фильтра: «Все категории контрагентов»;\n4. Нажатие кнопки «Сформировать»;\n5. Выполнение запроса к PostgreSQL, агрегация сумм начислений и оплат;\n6. Вывод аналитической таблицы на экран;\n7. Отправка отчета на печать или экспорт в XLSX."),
        ("Альтернативное направление (unhappy flow)", "За выбранный интервал времени у клиентов не зафиксировано обращений. Система выводит сообщение «За выбранный период активности клиентов не зарегистрировано» без сбоев."),
        ("Исключения", "Таймаут выполнения сложного аналитического запроса к СУБД. Система информирует пользователя «Время ожидания ответа БД истекло. Уточните параметры выборки» и записывает лог в журнал."),
        ("Частота использования сценария", "Еженедельно (контроль оплат и дебиторской задолженности) и ежеквартально (оценка динамики клиентской базы и эффективности программ лояльности)."),
        ("Бизнес-правила", "В отчете корректно учитываются корпоративные скидки клиентов (15% для автопарка УУНиТ), а сальдо рассчитывается как разница между начисленной суммой по Актам и поступившими платежами."),
        ("Специальные требования", "Время формирования сводного отчета по всей клиентской базе не должно превышать 3.5 секунд.")
    ]
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

    doc.paragraphs[48].text = (
        "Реализация доработки выполняется в среде КИС «BMW102-ERP» на технологической платформе "
        "1С:Предприятие 8.3 КОРП с использованием Системы компоновки данных (СКД) и внешнего веб-кабинета "
        "корпоративного клиента УУНиТ на базе React/TypeScript."
    )
    doc.paragraphs[50].text = "Рисунок 2 – Макет экранной формы сеанса-конфигуратора параметров отчета"
    
    # Insert Figure 2 after P50
    insert_picture_after(doc.paragraphs[50], os.path.join(base_dir, 'tz_gaynanov_fig2_dialog.png'), width_inches=5.8)
    
    doc.paragraphs[51].text = (
        "Отчетная форма представляет собой интерактивную сводную ведомость контрагентов с детализацией по количеству "
        "закрепленных транспортных средств, объему оказанных услуг, размеру предоставленной скидки и текущему сальдо взаиморасчетов."
    )
    doc.paragraphs[52].text = "Атрибутивный состав формы приведён в Таблице 2."
    doc.paragraphs[53].text = "Таблица 2 – Атрибутивный состав формы отчета"
    
    # Table 4: Attributes of Report Form
    t4 = doc.tables[4]
    attrs_d = [
        ("1", "Номер п/п", "INT", "1, 2, 3...", "Порядковый номер контрагента в отчете"),
        ("2", "Код клиента", "VARCHAR(20)", "CL-01, CL-04...", "Уникальный системный код контрагента в КИС"),
        ("3", "Наименование клиента", "VARCHAR(100)", "ФГБОУ ВО «УУНиТ»", "Полное наименование организации или ФИО клиента"),
        ("4", "Категория клиента", "VARCHAR(50)", "Корпоративный / VIP", "Категория договора и условий обслуживания"),
        ("5", "Количество ТС", "INT", "14 ед.", "Число зарегистрированных автомобилей в карточке клиента"),
        ("6", "Начислено (руб.)", "DECIMAL(12,2)", "542 000,00 руб.", "Общая стоимость услуг по закрытым нарядам за период")
    ]
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

    doc.paragraphs[54].text = "Связи таблиц базы данных КИС, из которых формируется отчет, приведены на схеме:"
    doc.paragraphs[55].text = "Рисунок 3 – Реляционная модель данных (ERD) подпроцесса учета клиентов и заказов"
    
    # Insert Figure 3 after P55
    insert_picture_after(doc.paragraphs[55], os.path.join(base_dir, 'tz_gaynanov_fig3_erd.png'), width_inches=6.0)

    doc.paragraphs[57].text = (
        "Запуск формирования формы выполняется пользователем посредством выбора пункта меню "
        "«CRM / Отчеты / Отчет по клиентам за период» на главной панели команд КИС «BMW102-ERP». "
        "При этом открывается диалоговый сеанс-конфигуратор со следующим набором управляющих атрибутов:"
    )
    doc.paragraphs[58].text = "Таблица 3 – Атрибуты сеанса-конфигуратора ОФ"
    
    # Table 5: Configurator Attributes
    t5 = doc.tables[5]
    cfg_d = [
        ("1", "Период (С .. ПО)", "Дата (с .. по)", "Задаёт отчетный временной интервал (квартал, полугодие, год)"),
        ("2", "Категория контрагента", "Список выбора", "Фильтр: «Все», «Корпоративные автопарки», «Физические лица»"),
        ("3", "Выгрузка в xls/pdf", "Флаг (булево)", "Если флаг активен, выполняется экспорт в файл, иначе вывод на экран/печать")
    ]
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

    doc.paragraphs[60].text = (
        "На распечатанном бланке отчета автоматически выводится заголовок формы отчета "
        "«Отчет по клиентам за период», интервал дат выборки, категория клиентов и текущая дата печати."
    )
    doc.paragraphs[62].text = "Источники атрибутов формы с логикой их формирования указаны в Таблице 4."
    doc.paragraphs[63].text = "Таблица 4 – Источники атрибутов формы с логикой формирования"
    
    # Table 6: Sources
    t6 = doc.tables[6]
    sources_d = [
        ("1", "Код клиента", "CL-01", "VARCHAR", "Client.id_client", "Прямая выборка поля id_client"),
        ("2", "Наименование", "ФГБОУ ВО «УУНиТ»", "VARCHAR", "Client.client_name", "Прямая выборка поля client_name"),
        ("3", "Категория", "Корпоративный", "VARCHAR", "Client.client_type", "Прямая выборка client_type"),
        ("4", "Кол-во ТС", "14", "INT", "Vehicle.vin_code", "COUNT(DISTINCT Vehicle.vin_code)"),
        ("5", "Кол-во заявок", "12", "INT", "Service_Request.id_request", "COUNT(DISTINCT Service_Request.id_request)"),
        ("6", "Начислено (руб)", "542 000,00", "DECIMAL", "Work_Completion_Act.total_amount", "SUM(Work_Completion_Act.total_amount)"),
        ("7", "Скидка (%)", "15.00%", "DECIMAL", "Client.discount_percent", "Прямая выборка discount_percent")
    ]
    while len(t6.rows) - 1 < len(sources_d):
        t6.add_row()
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

    doc.paragraphs[67].text = (
        "Дополнительная логика не предусмотрена. Доступ к просмотру финансовой информации и задолженностей клиентов "
        "ограничивается ролевой моделью КИС: полный доступ открыт для ролей «Менеджер CRM», «Главный бухгалтер» и «Руководство». "
        "Менеджеры видят только агрегированные суммы без возможности прямой правки бухгалтерских проводок."
    )
    
    # Section 4: Testing
    doc.paragraphs[73].text = (
        "Предварительных настроек, сверх необходимых для штатного функционирования КИС «BMW102-ERP» "
        "и СУБД PostgreSQL, не требуется. В тестовой базе должны быть заведены карточки контрагентов и привязаны ТС."
    )
    doc.paragraphs[79].text = "Таблица 5 – Программа приемочных испытаний"
    
    # Table 7: Test Scenarios
    t7 = doc.tables[7]
    tests_d = [
        ("1", "Формирование сводного отчета по всей клиентской базе за 1 квартал", "ТР-01", "Проверка корректности расчета сумм и группировок"),
        ("2", "Фильтрация по ключевому корпоративному заказчику (УУНиТ)", "ТР-02", "Проверка детальной ведомости по автопарку вуза"),
        ("3", "Контроль расчета персональной скидки 15% для автопарка УУНиТ", "ТР-03", "Проверка формулы скидки и сальдо взаиморасчетов"),
        ("4", "Экспорт сформированного клиентского отчета в форматы MS Excel и PDF", "ТР-04", "Проверка выгрузки таблиц, итогов и колонтитулов")
    ]
    while len(t7.rows) - 1 < len(tests_d):
        t7.add_row()
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

    doc.paragraphs[83].text = "Таблица 6 – Методика проведения испытаний"
    
    # Table 8: Test Steps
    t8 = doc.tables[8]
    steps_d = [
        ("Сценарий 1", "Шаг 1. Открыть «CRM / Отчеты» -> «Отчет по клиентам за период». Задать 01.01.2026 - 31.03.2026. Нажать «Сформировать».", "Форма строится не более 3.5 с. В отчете отображено 112 клиентов, 246 автомобилей, начислено 4 980 000,00 руб."),
        ("Сценарий 2", "Шаг 2. В поле «Контрагент» выбрать «ФГБОУ ВО «УУНиТ»», нажать «Сформировать».", "Отчет сформирован по автопарку УУНиТ: 14 ТС, 12 нарядов, начислено 542 000,00 руб., долг 0,00 руб."),
        ("Сценарий 3", "Шаг 3. Проверить значение в колонке «Скидка» для УУНиТ.", "Отображается значение 15%, итоговая сумма рассчитана строго с учетом корпоративного дисконта."),
        ("Сценарий 4", "Шаг 4. Нажать кнопку «Экспорт в XLSX». Открыть сохраненный файл.", "Файл корректно открывается в MS Excel, итоговые формулы суммирования функционируют, форматирование сохранено."),
        ("Сценарий 5", "Шаг 5. Нажать кнопку «Печать». Выполнить предварительный просмотр.", "Сформирован корректный печатный лист А4 альбомной ориентации, колонтитулы и подписи на месте.")
    ]
    set_cell_text(t8.rows[0].cells[0], "Сценарий", bold=True, size_pt=9, align='center')
    set_cell_text(t8.rows[0].cells[1], "Действие оператора", bold=True, size_pt=9, align='center')
    set_cell_text(t8.rows[0].cells[2], "Ожидаемый результат", bold=True, size_pt=9, align='center')
    for ci in range(3): set_cell_background(t8.rows[0].cells[ci], "E2E8F0")
    for idx, (sc, act, exp) in enumerate(steps_d):
        row = t8.rows[idx + 1]
        set_cell_text(row.cells[0], sc, bold=True, size_pt=8.5, align='center')
        set_cell_text(row.cells[1], act, size_pt=8.5)
        set_cell_text(row.cells[2], exp, size_pt=8.5)

    doc.paragraphs[85].text = (
        "Испытания проводятся на тестовом сервере Исполнителя с использованием контрольной базы данных КИС «BMW102-ERP». "
        "Условия проведения: рабочая станция под управлением ОС Windows 10/11, тонкий клиент платформы 1С или браузер. "
        "По результатам испытаний оформляется Протокол приемочных испытаний доработки аналитической отчетности."
    )
    
    # Save temp and repack
    temp_p = os.path.join(base_dir, 'temp_tz_gaynanov.docx')
    out_final = os.path.join(base_dir, 'ШАБЛОН_ТЗ_Гайнанов_Д_И.docx')
    doc.save(temp_p)
    repack_tz(temp_p, out_final)
    if os.path.exists(temp_p): os.remove(temp_p)
    print("Saved Gaynanov TZ to:", out_final)

# Repack ZIP helper to inject logos and fix headers
def repack_tz(source_docx, dest_docx):
    # Load clean header from ШАБЛОН ДБС.docx where both logos are configured!
    dbs_path = os.path.join(base_dir, 'ШАБЛОН ДБС.docx')
    with zipfile.ZipFile(dbs_path, 'r') as z_dbs:
        h1_xml = z_dbs.read('word/header1.xml')
        h1_rels = z_dbs.read('word/_rels/header1.xml.rels')

    with zipfile.ZipFile(source_docx, 'r') as zin, zipfile.ZipFile(dest_docx, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            fname = item.filename
            data = zin.read(fname)
            
            # Replace logos
            if fname == 'word/media/image1.png':
                data = uunit_bytes
            elif fname == 'word/media/image2.png':
                data = exec_bytes
            # Add exec logo to media if not present
            elif fname == 'word/footer3.xml':
                # footer 3 date -> 2026
                data = data.decode('utf-8', errors='ignore').replace('2023', '2026').replace('2025', '2026').encode('utf-8')
            # Use identical clean header table (with both logos and STYLEREF) for headers
            elif fname in ['word/header1.xml', 'word/header2.xml', 'word/header3.xml']:
                data = h1_xml
            elif fname in ['word/_rels/header1.xml.rels', 'word/_rels/header2.xml.rels', 'word/_rels/header3.xml.rels']:
                data = h1_rels
                
            zout.writestr(item, data)
        # Ensure media/image2.png is inside zout
        zout.writestr('word/media/image2.png', exec_bytes)
        zout.writestr('word/media/image1.png', uunit_bytes)

build_garifullin_tz()
build_gaynanov_tz()
print("Both TZ documents generated successfully!")

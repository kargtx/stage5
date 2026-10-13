import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# FILE 1: Практика_2_Гарифуллин.xlsx
# Стиль: Изумрудно-зеленый (хвоя/мята), шрифт Calibri, 1 лист с мини-сводкой
# вверху и аккуратными решениями задач 6-12 внизу.
# =============================================================================

def build_garifullin_file_pr2():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Практика 2 (Задачи 6-12)"
    ws.views.sheetView[0].showGridLines = True
    
    FONT = 'Calibri'
    
    # Стили шрифтов
    f_title = Font(name=FONT, size=13, bold=True, color='FFFFFF')
    f_subtitle = Font(name=FONT, size=10, bold=False, color='E8F5E9')
    f_sec = Font(name=FONT, size=11, bold=True, color='1E4D2B')
    f_th = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    f_reg = Font(name=FONT, size=10, bold=False, color='202020')
    f_bold = Font(name=FONT, size=10, bold=True, color='202020')
    f_math = Font(name=FONT, size=9, italic=True, color='1E4D2B')
    f_res = Font(name=FONT, size=10, bold=True, color='155724')
    f_cond = Font(name=FONT, size=9, italic=True, color='4A4A4A')
    
    # Цвета (Зеленая гамма)
    fill_main = PatternFill(start_color='234E35', end_color='234E35', fill_type='solid')      # темно-зеленый
    fill_sub = PatternFill(start_color='2E6930', end_color='2E6930', fill_type='solid')       # средний зеленый
    fill_th = PatternFill(start_color='3B7A40', end_color='3B7A40', fill_type='solid')        # шапка таблиц
    fill_card = PatternFill(start_color='E8F2EA', end_color='E8F2EA', fill_type='solid')      # светлая подложка
    fill_res = PatternFill(start_color='D4EDDA', end_color='D4EDDA', fill_type='solid')       # результат (мята)
    fill_zebra = PatternFill(start_color='F5F9F6', end_color='F5F9F6', fill_type='solid')
    fill_ans = PatternFill(start_color='E8F5E9', end_color='E8F5E9', fill_type='solid')
    
    # Границы
    bd_gray = Side(style='thin', color='C8D6CB')
    bd_green = Side(style='thin', color='2E6930')
    bd_double = Side(style='double', color='2E6930')
    
    box_cell = Border(left=bd_gray, right=bd_gray, top=bd_gray, bottom=bd_gray)
    box_total = Border(left=bd_gray, right=bd_gray, top=bd_gray, bottom=bd_double)
    box_res = Border(left=bd_green, right=bd_green, top=bd_green, bottom=bd_green)
    
    # Выравнивания
    al_c = Alignment(horizontal='center', vertical='center')
    al_l = Alignment(horizontal='left', vertical='center')
    al_l_wrap = Alignment(horizontal='left', vertical='center', wrap_text=True)
    al_r = Alignment(horizontal='right', vertical='center')
    
    def apply_style(row, c_start, c_end, font=None, fill=None, border=None, al=None):
        for c in range(c_start, c_end + 1):
            cell = ws.cell(row=row, column=c)
            if font: cell.font = font
            if fill: cell.fill = fill
            if border: cell.border = border
            if al: cell.alignment = al

    # Шапка студента
    ws.merge_cells('B2:F2')
    ws['B2'] = "ПРАКТИЧЕСКАЯ РАБОТА № 2"
    ws['B2'].font = f_title; ws['B2'].alignment = al_c
    apply_style(2, 2, 6, fill=fill_main)
    ws.row_dimensions[2].height = 24
    
    ws.merge_cells('B3:F3')
    ws['B3'] = "Аналитическое определение количественных характеристик надежности изделия"
    ws['B3'].font = f_subtitle; ws['B3'].alignment = al_c
    apply_style(3, 2, 6, fill=fill_sub)
    ws.row_dimensions[3].height = 18
    
    ws.merge_cells('B4:F4')
    ws['B4'] = "Выполнил: студент Гарифуллин  |  Дисциплина: Надежность автоматизированных систем"
    ws['B4'].font = Font(name=FONT, size=9, italic=True, color='E8F5E9'); ws['B4'].alignment = al_c
    apply_style(4, 2, 6, fill=fill_sub)
    ws.row_dimensions[4].height = 18

    # Сводная таблица ответов наверху
    ws['B6'] = "Сводные результаты решения задач (№ 6 – 12)"
    ws['B6'].font = f_sec
    ws.row_dimensions[6].height = 20
    
    sum_heads = ["№ задачи", "Краткое описание", "Параметр", "Формула", "Результат", "Ед. изм."]
    for i, h in enumerate(sum_heads, start=2):
        cell = ws.cell(row=7, column=i, value=h)
        cell.font = f_th; cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws.row_dimensions[7].height = 22

    # Генерация задач 6-12, начиная со строки 30
    cur_r = 30
    results_map = {}
    
    def add_task(title, cond, inputs, calcs, note, task_id):
        nonlocal cur_r
        r = cur_r
        
        # Шапка задачи
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.cell(row=r, column=2, value=title).font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
        apply_style(r, 2, 6, fill=fill_sub, al=al_l)
        ws.row_dimensions[r].height = 20
        
        # Условие
        r += 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=6)
        c_c = ws.cell(row=r, column=2, value=f"Условие: {cond}")
        c_c.font = f_cond; c_c.alignment = al_l_wrap
        apply_style(r, 2, 6, fill=fill_card)
        apply_style(r+1, 2, 6, fill=fill_card)
        ws.row_dimensions[r].height = 16
        ws.row_dimensions[r+1].height = 16
        r += 1
        
        # Исходные данные
        r += 1
        ws.cell(row=r, column=2, value="Исходные данные:").font = f_bold
        r += 1
        for ci, h in enumerate(["Параметр", "Обозначение", "Значение", "Ед. изм."], start=2):
            cell = ws.cell(row=r, column=ci, value=h)
            cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
            cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
        ws.row_dimensions[r].height = 19
        
        inp_cells = {}
        for row_i in inputs:
            r += 1
            ws.row_dimensions[r].height = 18
            p_name, p_sym, p_val, p_unit = row_i
            if isinstance(p_val, str) and '{' in p_val:
                p_val = p_val.format(**inp_cells)
            
            c1 = ws.cell(row=r, column=2, value=p_name); c1.font = f_reg; c1.border = box_cell
            c2 = ws.cell(row=r, column=3, value=p_sym); c2.font = f_bold; c2.alignment = al_c; c2.border = box_cell
            c3 = ws.cell(row=r, column=4, value=p_val); c3.font = f_bold; c3.alignment = al_r; c3.border = box_cell
            if isinstance(p_val, (int, float)):
                if abs(p_val) < 0.0001 and p_val != 0:
                    c3.number_format = '0.00E+00'
                elif isinstance(p_val, float):
                    c3.number_format = '#,##0.####'
                else:
                    c3.number_format = '#,##0'
            c4 = ws.cell(row=r, column=5, value=p_unit); c4.font = f_reg; c4.alignment = al_c; c4.border = box_cell
            
            inp_cells[p_sym] = f"D{r}"
            inp_cells[p_sym.replace('(', '_').replace(')', '')] = f"D{r}"
        
        # Расчет
        r += 2
        ws.cell(row=r, column=2, value="Расчет показателей:").font = f_bold
        r += 1
        for ci, h in enumerate(["Показатель", "Формула", "Расчет в Excel", "Значение", "Ед. изм."], start=2):
            cell = ws.cell(row=r, column=ci, value=h)
            cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
            cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
        ws.row_dimensions[r].height = 19
        
        calc_cells = {}
        for row_c in calcs:
            r += 1
            ws.row_dimensions[r].height = 20
            c_name, c_math, c_template, c_unit, c_fmt, c_key = row_c
            formula = c_template.format(**inp_cells, **calc_cells)
            
            c1 = ws.cell(row=r, column=2, value=c_name); c1.font = f_reg; c1.border = box_cell
            c2 = ws.cell(row=r, column=3, value=c_math); c2.font = f_math; c2.alignment = al_c; c2.border = box_cell
            c3 = ws.cell(row=r, column=4, value=formula); c3.font = f_math; c3.border = box_cell
            c4 = ws.cell(row=r, column=5, value=formula); c4.font = f_res; c4.fill = fill_res; c4.alignment = al_r; c4.border = box_res
            c4.number_format = c_fmt
            c5 = ws.cell(row=r, column=6, value=c_unit); c5.font = f_reg; c5.alignment = al_c; c5.border = box_cell
            
            calc_cells[c_key] = f"E{r}"
            results_map[f"{task_id}_{c_key}"] = f"E{r}"
        
        # Ответ
        r += 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ans_cell = ws.cell(row=r, column=2, value=f"Ответ: {note}")
        ans_cell.font = f_bold
        apply_style(r, 2, 6, fill=fill_ans, border=box_cell, al=al_l)
        ws.row_dimensions[r].height = 20
        
        cur_r = r + 2

    # Добавление задач 6-11
    add_task(
        "Задача 6. Показатели надежности системы при L1 = 460 ч (t = 800 ч)",
        "Средняя наработка до отказа L1 системы равна 460 ч. Для экспоненциального закона надежности определить P(t), f(t) и λ(t) за 800 ч работы системы.",
        [("Средняя наработка L1", "L1", 460, "ч"), ("Время работы t", "t", 800, "ч")],
        [
            ("Интенсивность отказов", "1 / L1", "=1/{L1}", "1/ч", "0.000000", "l"),
            ("Вероятность безотказной работы", "exp(-λ·t)", "=EXP(-{l}*{t})", "—", "0.0000", "p"),
            ("Частота отказов", "λ · P(t)", "={l}*{p}", "1/ч", "0.000000", "f")
        ],
        "λ(800) = 0,002174 1/ч; P(800) = 0,1757 (17,57%); f(800) = 0,000382 1/ч.",
        "t6"
    )

    add_task(
        "Задача 7. Надежность системы из 6000 последовательных элементов (T = 50 ч)",
        "Система управления состоит из 6000 элементов, соединенных последовательно, средняя интенсивность отказов которых λср = 0,2∙10⁻⁶ ч⁻¹. Найти Pc(50) и mtc.",
        [("Число элементов", "n", 6000, "шт."), ("Интенсивность λср", "l_elem", 0.0000002, "1/ч"), ("Время Т", "T", 50, "ч")],
        [
            ("Интенсивность системы λc", "n · λср", "={n}*{l_elem}", "1/ч", "0.000000", "lc"),
            ("Вероятность безотказной работы", "exp(-λc·T)", "=EXP(-{lc}*{T})", "—", "0.0000", "pc"),
            ("Среднее время работы системы", "1 / λc", "=1/{lc}", "ч", "0.00", "mtc")
        ],
        "Pc(50) = 0,9418 (94,18%); mtc = 833,33 ч (λc = 1,2·10⁻³ 1/ч).",
        "t7"
    )

    add_task(
        "Задача 8. Надежность системы из 5000 последовательных элементов (T = 120 ч)",
        "Система управления состоит из 5000 элементов, соединенных последовательно, средняя интенсивность отказов которых λср = 0,3∙10⁻⁶ ч⁻¹. Найти Pc(120) и mtc.",
        [("Число элементов", "n", 5000, "шт."), ("Интенсивность λср", "l_elem", 0.0000003, "1/ч"), ("Время Т", "T", 120, "ч")],
        [
            ("Интенсивность системы λc", "n · λср", "={n}*{l_elem}", "1/ч", "0.000000", "lc"),
            ("Вероятность безотказной работы", "exp(-λc·T)", "=EXP(-{lc}*{T})", "—", "0.0000", "pc"),
            ("Среднее время работы системы", "1 / λc", "=1/{lc}", "ч", "0.00", "mtc")
        ],
        "Pc(120) = 0,8353 (83,53%); mtc = 666,67 ч (λc = 1,5·10⁻³ 1/ч).",
        "t8"
    )

    add_task(
        "Задача 9. Надежность невосстанавливаемой ЭВМ из 200000 элементов (сутки)",
        "Электронная машина состоит из 200000 последовательных элементов с λср = 0,3∙10⁻⁶ ч⁻¹. Определить Pc(24) в течение суток и mtc.",
        [("Число элементов", "n", 200000, "шт."), ("Интенсивность λср", "l_elem", 0.0000003, "1/ч"), ("Время в сутках t", "t", 24, "ч")],
        [
            ("Интенсивность ЭВМ λc", "n · λср", "={n}*{l_elem}", "1/ч", "0.000000", "lc"),
            ("Вероятность работы за сутки", "exp(-λc·t)", "=EXP(-{lc}*{t})", "—", "0.0000", "pc"),
            ("Среднее время работы ЭВМ", "1 / λc", "=1/{lc}", "ч", "0.00", "mtc")
        ],
        "Pc(24) = 0,2369 (23,69%); mtc = 16,67 ч (16 ч 40 мин; λc = 0,06 1/ч).",
        "t9"
    )

    add_task(
        "Задача 10. Расчет Pc(t), qc(t), fc(t), mtc для системы из 12600 элементов",
        "Система состоит из 12600 элементов, средняя интенсивность отказов которых λср = 0,32∙10⁻⁶ ч⁻¹. Определить Pc(t), qc(t), fc(t), mtc для t = 50 ч.",
        [("Число элементов", "n", 12600, "шт."), ("Интенсивность λср", "l_elem", 0.00000032, "1/ч"), ("Время наработки t", "t", 50, "ч")],
        [
            ("Интенсивность системы λc", "n · λср", "={n}*{l_elem}", "1/ч", "0.000000", "lc"),
            ("Вероятность безотказной работы", "exp(-λc·t)", "=EXP(-{lc}*{t})", "—", "0.0000", "pc"),
            ("Вероятность отказа qc", "1 - Pc(t)", "=1-{pc}", "—", "0.0000", "qc"),
            ("Частота отказов fc", "λc · Pc(t)", "={lc}*{pc}", "1/ч", "0.000000", "fc"),
            ("Среднее время работы mtc", "1 / λc", "=1/{lc}", "ч", "0.00", "mtc")
        ],
        "Pc(50) = 0,8174; qc(50) = 0,1826; fc(50) = 0,003296 1/ч; mtc = 248,02 ч.",
        "t10"
    )

    add_task(
        "Задача 11. Характеристики прибора с λ = 0,1∙10⁻⁶ ч⁻¹ (t = 180 ч)",
        "Время работы прибора до отказа подчинено экспоненциальному закону с λ = 0,1∙10⁻⁶ ч⁻¹; t = 180 ч. Определить p(t), q(t), f(t), mt.",
        [("Интенсивность λ", "l_val", 0.0000001, "1/ч"), ("Время наработки t", "t", 180, "ч")],
        [
            ("Вероятность безотказной работы", "exp(-λ·t)", "=EXP(-{l_val}*{t})", "—", "0.000000", "p"),
            ("Вероятность отказа q", "1 - p(t)", "=1-{p}", "—", "0.000000", "q"),
            ("Частота отказов f", "λ · p(t)", "={l_val}*{p}", "1/ч", "0.00E+00", "f"),
            ("Среднее время работы mt", "1 / λ", "=1/{l_val}", "ч", "#,##0", "mt")
        ],
        "p(180) = 0,999982; q(180) = 0,000018 (1,8·10⁻⁵); f(180) = 1,00·10⁻⁷ 1/ч; mt = 10 000 000 ч.",
        "t11"
    )

    # Задача 12 (Система из 5 приборов)
    r = cur_r
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2, value="Задача 12. Система из 5 приборов").font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    apply_style(r, 2, 6, fill=fill_sub, al=al_l)
    ws.row_dimensions[r].height = 20
    
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=6)
    c_c12 = ws.cell(row=r, column=2, value="Условие: Система из 5 приборов: mt1 = 83 ч, mt2 = 220 ч, mt3 = 280 ч, mt4 = 400 ч, mt5 = 700 ч. Экспоненциальный закон. Найти mtc системы.")
    c_c12.font = f_cond; c_c12.alignment = al_l_wrap
    apply_style(r, 2, 6, fill=fill_card)
    apply_style(r+1, 2, 6, fill=fill_card)
    ws.row_dimensions[r].height = 16
    ws.row_dimensions[r+1].height = 16
    r += 1
    
    r += 1
    ws.cell(row=r, column=2, value="Исходные данные и расчет интенсивностей приборов:").font = f_bold
    r += 1
    for ci, h in enumerate(["№", "Обозначение", "mt_i (ч)", "λ_i = 1/mt_i (1/ч)", "Примечание"], start=2):
        cell = ws.cell(row=r, column=ci, value=h)
        cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
        cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws.row_dimensions[r].height = 19
    
    devs = [
        ("1", "mt1", 83, "Прибор 1"),
        ("2", "mt2", 220, "Прибор 2"),
        ("3", "mt3", 280, "Прибор 3"),
        ("4", "mt4", 400, "Прибор 4"),
        ("5", "mt5", 700, "Прибор 5")
    ]
    st_d = r + 1
    for dev in devs:
        r += 1
        ws.row_dimensions[r].height = 18
        ws.cell(row=r, column=2, value=dev[0]).font = f_reg; ws.cell(row=r, column=2).alignment = al_c; ws.cell(row=r, column=2).border = box_cell
        ws.cell(row=r, column=3, value=dev[1]).font = f_bold; ws.cell(row=r, column=3).alignment = al_c; ws.cell(row=r, column=3).border = box_cell
        c_m = ws.cell(row=r, column=4, value=dev[2]); c_m.font = f_bold; c_m.alignment = al_r; c_m.border = box_cell; c_m.number_format = '#,##0'
        c_l = ws.cell(row=r, column=5, value=f"=1/D{r}"); c_l.font = f_bold; c_l.alignment = al_r; c_l.border = box_cell; c_l.number_format = '0.000000'
        ws.cell(row=r, column=6, value=dev[3]).font = f_cond; ws.cell(row=r, column=6).border = box_cell
    end_d = r
    
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws.cell(row=r, column=2, value="Суммарная интенсивность λc = Σ λ_i :").font = f_bold
    ws.cell(row=r, column=2).alignment = al_r
    apply_style(r, 2, 4, fill=fill_card, border=box_total)
    c_slc = ws.cell(row=r, column=5, value=f"=SUM(E{st_d}:E{end_d})")
    c_slc.font = f_res; c_slc.alignment = al_r; c_slc.fill = fill_res; c_slc.border = box_total; c_slc.number_format = '0.000000'
    ws.cell(row=r, column=6, value="1/ч").font = f_bold; ws.cell(row=r, column=6).alignment = al_c; ws.cell(row=r, column=6).border = box_total
    ref_lc_12 = f"E{r}"
    results_map["t12_lc"] = ref_lc_12
    
    r += 2
    ws.cell(row=r, column=2, value="Итоговый расчет mtc системы:").font = f_bold
    r += 1
    for ci, h in enumerate(["Показатель", "Формула", "Расчет в Excel", "Значение", "Ед. изм."], start=2):
        cell = ws.cell(row=r, column=ci, value=h)
        cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
        cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws.row_dimensions[r].height = 19
    
    r += 1
    ws.row_dimensions[r].height = 20
    ws.cell(row=r, column=2, value="Средняя наработка системы mtc").font = f_reg; ws.cell(row=r, column=2).border = box_cell
    ws.cell(row=r, column=3, value="1 / λc").font = f_math; ws.cell(row=r, column=3).alignment = al_c; ws.cell(row=r, column=3).border = box_cell
    ws.cell(row=r, column=4, value=f"=1/{ref_lc_12}").font = f_math; ws.cell(row=r, column=4).border = box_cell
    c_fmtc = ws.cell(row=r, column=5, value=f"=1/{ref_lc_12}")
    c_fmtc.font = f_res; c_fmtc.fill = fill_res; c_fmtc.alignment = al_r; c_fmtc.border = box_res; c_fmtc.number_format = '0.00'
    ws.cell(row=r, column=6, value="ч").font = f_reg; ws.cell(row=r, column=6).alignment = al_c; ws.cell(row=r, column=6).border = box_cell
    results_map["t12_mtc"] = f"E{r}"
    
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2, value="Ответ: Интенсивность отказов λc = 0,024094 1/ч; средняя наработка системы mtc = 41,50 ч.").font = f_bold
    apply_style(r, 2, 6, fill=fill_ans, border=box_cell, al=al_l)
    ws.row_dimensions[r].height = 20

    # Заполнение сводной таблицы наверху (строки 8 - 24)
    sum_rows = [
        ("Задание 6", "L1=460 ч, t=800 ч. Интенсивность отказов", "λ(800)", "λ = 1 / L1", f"={results_map['t6_l']}", "1/ч", "0.000000"),
        ("Задание 6", "L1=460 ч, t=800 ч. Вероятность безотказной работы", "P(800)", "P = exp(-λ·t)", f"={results_map['t6_p']}", "—", "0.0000"),
        ("Задание 6", "L1=460 ч, t=800 ч. Частота отказов", "f(800)", "f = λ · P(t)", f"={results_map['t6_f']}", "1/ч", "0.000000"),
        ("Задание 7", "6000 элементов, λср=0.2·10⁻⁶ ч⁻¹, T=50 ч", "Pc(50)", "Pc = exp(-λc·T)", f"={results_map['t7_pc']}", "—", "0.0000"),
        ("Задание 7", "6000 элементов, λср=0.2·10⁻⁶ ч⁻¹, T=50 ч", "mtc", "mtc = 1 / λc", f"={results_map['t7_mtc']}", "ч", "0.00"),
        ("Задание 8", "5000 элементов, λср=0.3·10⁻⁶ ч⁻¹, T=120 ч", "Pc(120)", "Pc = exp(-λc·T)", f"={results_map['t8_pc']}", "—", "0.0000"),
        ("Задание 8", "5000 элементов, λср=0.3·10⁻⁶ ч⁻¹, T=120 ч", "mtc", "mtc = 1 / λc", f"={results_map['t8_mtc']}", "ч", "0.00"),
        ("Задание 9", "200000 элементов, λср=0.3·10⁻⁶ ч⁻¹, сутки (t=24 ч)", "Pc(24)", "Pc = exp(-λc·t)", f"={results_map['t9_pc']}", "—", "0.0000"),
        ("Задание 9", "200000 элементов, λср=0.3·10⁻⁶ ч⁻¹, сутки (t=24 ч)", "mtc", "mtc = 1 / λc", f"={results_map['t9_mtc']}", "ч", "0.00"),
        ("Задание 10", "12600 элементов, λср=0.32·10⁻⁶ ч⁻¹, t=50 ч", "Pc(50)", "Pc = exp(-λc·t)", f"={results_map['t10_pc']}", "—", "0.0000"),
        ("Задание 10", "12600 элементов, λср=0.32·10⁻⁶ ч⁻¹, t=50 ч", "qc(50)", "qc = 1 - Pc", f"={results_map['t10_qc']}", "—", "0.0000"),
        ("Задание 10", "12600 элементов, λср=0.32·10⁻⁶ ч⁻¹, t=50 ч", "fc(50)", "fc = λc · Pc", f"={results_map['t10_fc']}", "1/ч", "0.000000"),
        ("Задание 10", "12600 элементов, λср=0.32·10⁻⁶ ч⁻¹, t=50 ч", "mtc", "mtc = 1 / λc", f"={results_map['t10_mtc']}", "ч", "0.00"),
        ("Задание 11", "Прибор с λ=0.1·10⁻⁶ ч⁻¹, t=180 ч", "p(180)", "p = exp(-λ·t)", f"={results_map['t11_p']}", "—", "0.000000"),
        ("Задание 11", "Прибор с λ=0.1·10⁻⁶ ч⁻¹, t=180 ч", "q(180)", "q = 1 - p", f"={results_map['t11_q']}", "—", "0.000000"),
        ("Задание 11", "Прибор с λ=0.1·10⁻⁶ ч⁻¹, t=180 ч", "f(180)", "f = λ · p", f"={results_map['t11_f']}", "1/ч", "0.00E+00"),
        ("Задание 11", "Прибор с λ=0.1·10⁻⁶ ч⁻¹, t=180 ч", "mt", "mt = 1 / λ", f"={results_map['t11_mt']}", "ч", "#,##0"),
        ("Задание 12", "Система из 5 приборов (mti = {83, 220, 280, 400, 700})", "λc", "λc = Σ (1/mti)", f"={results_map['t12_lc']}", "1/ч", "0.000000"),
        ("Задание 12", "Система из 5 приборов (mti = {83, 220, 280, 400, 700})", "mtc", "mtc = 1 / λc", f"={results_map['t12_mtc']}", "ч", "0.00")
    ]
    
    for idx, s_row in enumerate(sum_rows, start=8):
        ws.row_dimensions[idx].height = 19
        bg = fill_zebra if idx % 2 == 0 else PatternFill(fill_type=None)
        
        c1 = ws.cell(row=idx, column=2, value=s_row[0]); c1.font = f_bold; c1.alignment = al_c; c1.border = box_cell; c1.fill = bg
        c2 = ws.cell(row=idx, column=3, value=s_row[1]); c2.font = f_reg; c2.alignment = al_l; c2.border = box_cell; c2.fill = bg
        c3 = ws.cell(row=idx, column=4, value=s_row[2]); c3.font = f_bold; c3.alignment = al_c; c3.border = box_cell; c3.fill = bg
        c4 = ws.cell(row=idx, column=5, value=s_row[3]); c4.font = f_math; c4.alignment = al_l; c4.border = box_cell; c4.fill = bg
        c5 = ws.cell(row=idx, column=6, value=s_row[4]); c5.font = f_res; c5.fill = fill_res; c5.alignment = al_r; c5.border = box_res
        c5.number_format = s_row[6]
        c6 = ws.cell(row=idx, column=7, value=s_row[5]); c6.font = f_reg; c6.alignment = al_c; c6.border = box_cell; c6.fill = bg

    col_w = {"A": 3, "B": 14, "C": 42, "D": 14, "E": 22, "F": 16, "G": 12}
    for col, width in col_w.items():
        ws.column_dimensions[col].width = width

    fname = "Практика_2_Гарифуллин.xlsx"
    wb.save(fname)
    print(f"Created: {fname}")


# =============================================================================
# FILE 2: Практика_2_Гайнетдинов.xlsx
# Стиль: Графит / Стальной синий (Slate), 2 листа:
# Лист 1: "Итоговая таблица"
# Лист 2: "Решения (№6-12)"
# =============================================================================

def build_gaynetdinov_file_pr2():
    wb = openpyxl.Workbook()
    FONT = 'Calibri'
    
    # Стили шрифтов
    f_title = Font(name=FONT, size=13, bold=True, color='FFFFFF')
    f_subtitle = Font(name=FONT, size=10, bold=False, color='F2F4F4')
    f_sec = Font(name=FONT, size=11, bold=True, color='2C3E50')
    f_th = Font(name=FONT, size=9, bold=True, color='FFFFFF')
    f_reg = Font(name=FONT, size=9, bold=False, color='2C3E50')
    f_bold = Font(name=FONT, size=9, bold=True, color='2C3E50')
    f_math = Font(name=FONT, size=9, italic=True, color='2980B9')
    f_res = Font(name=FONT, size=10, bold=True, color='1B4F72')
    f_cond = Font(name=FONT, size=9, italic=True, color='5D6D7E')
    
    # Цвета (Графит / Сине-серый slate)
    fill_dark = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')     # графит
    fill_med = PatternFill(start_color='34495E', end_color='34495E', fill_type='solid')      # сланец
    fill_th = PatternFill(start_color='4A6572', end_color='4A6572', fill_type='solid')       # шапка таблиц
    fill_card = PatternFill(start_color='F4F6F7', end_color='F4F6F7', fill_type='solid')     # светло-серый
    fill_res = PatternFill(start_color='D6EAF8', end_color='D6EAF8', fill_type='solid')      # светло-голубой результат
    fill_zebra = PatternFill(start_color='F8F9F9', end_color='F8F9F9', fill_type='solid')
    fill_ans = PatternFill(start_color='EBF5FB', end_color='EBF5FB', fill_type='solid')
    
    # Границы
    bd_gray = Side(style='thin', color='BDC3C7')
    bd_blue = Side(style='thin', color='2980B9')
    bd_double = Side(style='double', color='2C3E50')
    
    box_cell = Border(left=bd_gray, right=bd_gray, top=bd_gray, bottom=bd_gray)
    box_total = Border(left=bd_gray, right=bd_gray, top=bd_gray, bottom=bd_double)
    box_res = Border(left=bd_blue, right=bd_blue, top=bd_blue, bottom=bd_blue)
    
    # Выравнивания
    al_c = Alignment(horizontal='center', vertical='center')
    al_l = Alignment(horizontal='left', vertical='center')
    al_l_wrap = Alignment(horizontal='left', vertical='center', wrap_text=True)
    al_r = Alignment(horizontal='right', vertical='center')
    
    def apply_style(ws, row, c_start, c_end, font=None, fill=None, border=None, al=None):
        for c in range(c_start, c_end + 1):
            cell = ws.cell(row=row, column=c)
            if font: cell.font = font
            if fill: cell.fill = fill
            if border: cell.border = border
            if al: cell.alignment = al

    # ЛИСТ 2: РЕШЕНИЯ (№6-12)
    ws_sol = wb.active
    ws_sol.title = "Решения (№6-12)"
    ws_sol.views.sheetView[0].showGridLines = True
    
    # Шапка листа решений
    ws_sol.merge_cells('B2:G2')
    ws_sol['B2'] = "РАСЧЕТНАЯ ЧАСТЬ ПРАКТИЧЕСКОЙ РАБОТЫ № 2 (ЗАДАНИЯ 6 – 12)"
    ws_sol['B2'].font = f_title; ws_sol['B2'].alignment = al_c
    apply_style(ws_sol, 2, 2, 7, fill=fill_dark)
    ws_sol.row_dimensions[2].height = 24
    
    ws_sol.merge_cells('B3:G3')
    ws_sol['B3'] = "Аналитическая оценка показателей надежности  |  Студент: Гайнетдинов"
    ws_sol['B3'].font = f_subtitle; ws_sol['B3'].alignment = al_c
    apply_style(ws_sol, 3, 2, 7, fill=fill_med)
    ws_sol.row_dimensions[3].height = 18

    cur_r = 5
    results_map = {}
    
    def add_card(title, cond, inputs, calcs, note, task_id):
        nonlocal cur_r
        r = cur_r
        
        ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws_sol.cell(row=r, column=2, value=title).font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
        apply_style(ws_sol, r, 2, 7, fill=fill_med, al=al_l)
        ws_sol.row_dimensions[r].height = 20
        
        r += 1
        ws_sol.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=7)
        ws_sol.cell(row=r, column=2, value=f"Текст задания: {cond}").font = f_cond
        ws_sol.cell(row=r, column=2).alignment = al_l_wrap
        apply_style(ws_sol, r, 2, 7, fill=fill_card)
        apply_style(ws_sol, r+1, 2, 7, fill=fill_card)
        ws_sol.row_dimensions[r].height = 16
        ws_sol.row_dimensions[r+1].height = 16
        r += 1
        
        r += 1
        ws_sol.cell(row=r, column=2, value="1. Исходные параметры").font = f_bold
        r += 1
        for ci, h in enumerate(["Наименование", "Символ", "Значение", "Размерность"], start=2):
            cell = ws_sol.cell(row=r, column=ci, value=h)
            cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
            cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
        ws_sol.row_dimensions[r].height = 18
        
        inp_cells = {}
        for row_i in inputs:
            r += 1
            ws_sol.row_dimensions[r].height = 18
            p_name, p_sym, p_val, p_unit = row_i
            if isinstance(p_val, str) and '{' in p_val:
                p_val = p_val.format(**inp_cells)
            
            c1 = ws_sol.cell(row=r, column=2, value=p_name); c1.font = f_reg; c1.border = box_cell
            c2 = ws_sol.cell(row=r, column=3, value=p_sym); c2.font = f_bold; c2.alignment = al_c; c2.border = box_cell
            c3 = ws_sol.cell(row=r, column=4, value=p_val); c3.font = f_bold; c3.alignment = al_r; c3.border = box_cell
            if isinstance(p_val, (int, float)):
                if abs(p_val) < 0.0001 and p_val != 0:
                    c3.number_format = '0.00E+00'
                elif isinstance(p_val, float):
                    c3.number_format = '#,##0.####'
                else:
                    c3.number_format = '#,##0'
            c4 = ws_sol.cell(row=r, column=5, value=p_unit); c4.font = f_reg; c4.alignment = al_c; c4.border = box_cell
            
            inp_cells[p_sym] = f"D{r}"
            inp_cells[p_sym.replace('(', '_').replace(')', '')] = f"D{r}"
        
        r += 2
        ws_sol.cell(row=r, column=2, value="2. Аналитический расчет характеристик").font = f_bold
        r += 1
        for ci, h in enumerate(["Искомый показатель", "Матем. формула", "Формула в Excel", "Значение", "Размерность"], start=2):
            cell = ws_sol.cell(row=r, column=ci, value=h)
            cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
            cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
        ws_sol.row_dimensions[r].height = 18
        
        calc_cells = {}
        for row_c in calcs:
            r += 1
            ws_sol.row_dimensions[r].height = 19
            c_name, c_math, c_template, c_unit, c_fmt, c_key = row_c
            formula = c_template.format(**inp_cells, **calc_cells)
            
            c1 = ws_sol.cell(row=r, column=2, value=c_name); c1.font = f_reg; c1.border = box_cell
            c2 = ws_sol.cell(row=r, column=3, value=c_math); c2.font = f_math; c2.alignment = al_c; c2.border = box_cell
            c3 = ws_sol.cell(row=r, column=4, value=formula); c3.font = f_math; c3.border = box_cell
            c4 = ws_sol.cell(row=r, column=5, value=formula); c4.font = f_res; c4.fill = fill_res; c4.alignment = al_r; c4.border = box_res
            c4.number_format = c_fmt
            c5 = ws_sol.cell(row=r, column=6, value=c_unit); c5.font = f_reg; c5.alignment = al_c; c5.border = box_cell
            
            calc_cells[c_key] = f"E{r}"
            results_map[f"{task_id}_{c_key}"] = f"E{r}"
        
        r += 1
        ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ans_cell = ws_sol.cell(row=r, column=2, value=f"Вывод: {note}")
        ans_cell.font = f_bold
        apply_style(ws_sol, r, 2, 7, fill=fill_ans, border=box_cell, al=al_l)
        ws_sol.row_dimensions[r].height = 20
        
        cur_r = r + 2

    # Добавление задач 6-11
    add_card(
        "Задание 6. Расчет λ(t), P(t), f(t) при L1 = 460 ч (t = 800 ч)",
        "Средняя наработка до отказа L1 системы равна 460 ч. Для экспоненциального закона надежности определить P(t), f(t) и λ(t) за 800 ч работы системы.",
        [("Средняя наработка L1", "L1", 460, "ч"), ("Период эксплуатации t", "t", 800, "ч")],
        [
            ("Интенсивность отказов λ", "1 / L1", "=1/{L1}", "1/ч", "0.000000", "l"),
            ("Вероятность безотказной работы P", "exp(-λ·t)", "=EXP(-{l}*{t})", "—", "0.0000", "p"),
            ("Частота отказов f", "λ · P(t)", "={l}*{p}", "1/ч", "0.000000", "f")
        ],
        "При t = 800 ч: λ = 0,002174 1/ч, P(800) = 0,1757, f(800) = 0,000382 1/ч.",
        "t6"
    )

    add_card(
        "Задание 7. Система управления из 6000 элементов (T = 50 ч)",
        "Система управления состоит из 6000 последовательных элементов с λср = 0,2∙10⁻⁶ ч⁻¹. Определить Pc(50) и mtc.",
        [("Количество элементов n", "n", 6000, "шт."), ("Интенсивность элемента λср", "l_elem", 0.0000002, "1/ч"), ("Время T", "T", 50, "ч")],
        [
            ("Интенсивность отказов системы λc", "n · λср", "={n}*{l_elem}", "1/ч", "0.000000", "lc"),
            ("Вероятность безотказной работы Pc", "exp(-λc·T)", "=EXP(-{lc}*{T})", "—", "0.0000", "pc"),
            ("Среднее время работы системы mtc", "1 / λc", "=1/{lc}", "ч", "0.00", "mtc")
        ],
        "Pc(50) = 0,9418 (94,18%); средняя наработка mtc = 833,33 ч.",
        "t7"
    )

    add_card(
        "Задание 8. Система управления из 5000 элементов (T = 120 ч)",
        "Система управления состоит из 5000 последовательных элементов с λср = 0,3∙10⁻⁶ ч⁻¹. Определить Pc(120) и mtc.",
        [("Количество элементов n", "n", 5000, "шт."), ("Интенсивность элемента λср", "l_elem", 0.0000003, "1/ч"), ("Время T", "T", 120, "ч")],
        [
            ("Интенсивность отказов системы λc", "n · λср", "={n}*{l_elem}", "1/ч", "0.000000", "lc"),
            ("Вероятность безотказной работы Pc", "exp(-λc·T)", "=EXP(-{lc}*{T})", "—", "0.0000", "pc"),
            ("Среднее время работы системы mtc", "1 / λc", "=1/{lc}", "ч", "0.00", "mtc")
        ],
        "Pc(120) = 0,8353 (83,53%); средняя наработка mtc = 666,67 ч.",
        "t8"
    )

    add_card(
        "Задание 9. Электронная машина из 200000 элементов (t = 24 ч)",
        "Невосстанавливаемая машина состоит из 200000 последовательных элементов с λср = 0,3∙10⁻⁶ ч⁻¹. Найти Pc(24) за сутки и mtc.",
        [("Количество элементов n", "n", 200000, "шт."), ("Интенсивность элемента λср", "l_elem", 0.0000003, "1/ч"), ("Суточный интервал t", "t", 24, "ч")],
        [
            ("Интенсивность отказов машины λc", "n · λср", "={n}*{l_elem}", "1/ч", "0.000000", "lc"),
            ("Вероятность безотказной работы Pc", "exp(-λc·t)", "=EXP(-{lc}*{t})", "—", "0.0000", "pc"),
            ("Среднее время работы машины mtc", "1 / λc", "=1/{lc}", "ч", "0.00", "mtc")
        ],
        "Pc(24) = 0,2369; средняя наработка mtc = 16,67 ч (16 ч 40 мин).",
        "t9"
    )

    add_card(
        "Задание 10. Расчет Pc(t), qc(t), fc(t), mtc для системы из 12600 элементов",
        "Система состоит из 12600 элементов с λср = 0,32∙10⁻⁶ ч⁻¹. Определить Pc(50), qc(50), fc(50), mtc для t = 50 ч.",
        [("Количество элементов n", "n", 12600, "шт."), ("Интенсивность элемента λср", "l_elem", 0.00000032, "1/ч"), ("Время t", "t", 50, "ч")],
        [
            ("Интенсивность отказов системы λc", "n · λср", "={n}*{l_elem}", "1/ч", "0.000000", "lc"),
            ("Вероятность безотказной работы Pc", "exp(-λc·t)", "=EXP(-{lc}*{t})", "—", "0.0000", "pc"),
            ("Вероятность отказа qc", "1 - Pc", "=1-{pc}", "—", "0.0000", "qc"),
            ("Частота отказов fc", "λc · Pc", "={lc}*{pc}", "1/ч", "0.000000", "fc"),
            ("Средняя наработка системы mtc", "1 / λc", "=1/{lc}", "ч", "0.00", "mtc")
        ],
        "Pc(50) = 0,8174; qc(50) = 0,1826; fc(50) = 0,003296 1/ч; mtc = 248,02 ч.",
        "t10"
    )

    add_card(
        "Задание 11. Характеристики прибора с λ = 0,1∙10⁻⁶ ч⁻¹ (t = 180 ч)",
        "Время работы прибора подчинено экспоненциальному закону с λ = 0,1∙10⁻⁶ ч⁻¹; t = 180 ч. Определить p(180), q(180), f(180), mt.",
        [("Параметр λ", "l_val", 0.0000001, "1/ч"), ("Время работы t", "t", 180, "ч")],
        [
            ("Вероятность безотказной работы p", "exp(-λ·t)", "=EXP(-{l_val}*{t})", "—", "0.000000", "p"),
            ("Вероятность отказа q", "1 - p", "=1-{p}", "—", "0.000000", "q"),
            ("Частота отказов f", "λ · p", "={l_val}*{p}", "1/ч", "0.00E+00", "f"),
            ("Средняя наработка mt", "1 / λ", "=1/{l_val}", "ч", "#,##0", "mt")
        ],
        "p(180) = 0,999982; q(180) = 0,000018; f(180) = 1,00·10⁻⁷ 1/ч; mt = 10 000 000 ч.",
        "t11"
    )

    # Задание 12 (Система из 5 приборов)
    r = cur_r
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_sol.cell(row=r, column=2, value="Задание 12. Система из пяти последовательных приборов").font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    apply_style(ws_sol, r, 2, 7, fill=fill_med, al=al_l)
    ws_sol.row_dimensions[r].height = 20
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=7)
    c_c12 = ws_sol.cell(row=r, column=2, value="Текст задания: Система состоит из 5 приборов: mt1 = 83 ч, mt2 = 220 ч, mt3 = 280 ч, mt4 = 400 ч, mt5 = 700 ч. Экспоненциальный закон. Определить mtc системы.")
    c_c12.font = f_cond; c_c12.alignment = al_l_wrap
    apply_style(ws_sol, r, 2, 7, fill=fill_card)
    apply_style(ws_sol, r+1, 2, 7, fill=fill_card)
    ws_sol.row_dimensions[r].height = 16
    ws_sol.row_dimensions[r+1].height = 16
    r += 1
    
    r += 1
    ws_sol.cell(row=r, column=2, value="1. Исходные наработки и расчет интенсивностей").font = f_bold
    r += 1
    for ci, h in enumerate(["№ прибора", "Обозначение", "mt_i (ч)", "λ_i = 1/mt_i (1/ч)", "Описание"], start=2):
        cell = ws_sol.cell(row=r, column=ci, value=h)
        cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
        cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws_sol.row_dimensions[r].height = 18
    
    devs = [
        ("Прибор 1", "mt1", 83, "Прибор 1"),
        ("Прибор 2", "mt2", 220, "Прибор 2"),
        ("Прибор 3", "mt3", 280, "Прибор 3"),
        ("Прибор 4", "mt4", 400, "Прибор 4"),
        ("Прибор 5", "mt5", 700, "Прибор 5")
    ]
    st_d = r + 1
    for dev in devs:
        r += 1
        ws_sol.row_dimensions[r].height = 18
        ws_sol.cell(row=r, column=2, value=dev[0]).font = f_reg; ws_sol.cell(row=r, column=2).alignment = al_c; ws_sol.cell(row=r, column=2).border = box_cell
        ws_sol.cell(row=r, column=3, value=dev[1]).font = f_bold; ws_sol.cell(row=r, column=3).alignment = al_c; ws_sol.cell(row=r, column=3).border = box_cell
        c_m = ws_sol.cell(row=r, column=4, value=dev[2]); c_m.font = f_bold; c_m.alignment = al_r; c_m.border = box_cell; c_m.number_format = '#,##0'
        c_l = ws_sol.cell(row=r, column=5, value=f"=1/D{r}"); c_l.font = f_bold; c_l.alignment = al_r; c_l.border = box_cell; c_l.number_format = '0.000000'
        ws_sol.cell(row=r, column=6, value=dev[3]).font = f_cond; ws_sol.cell(row=r, column=6).border = box_cell
    end_d = r
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws_sol.cell(row=r, column=2, value="Суммарная интенсивность λc = Σ λ_i :").font = f_bold
    ws_sol.cell(row=r, column=2).alignment = al_r
    apply_style(ws_sol, r, 2, 4, fill=fill_card, border=box_total)
    c_slc = ws_sol.cell(row=r, column=5, value=f"=SUM(E{st_d}:E{end_d})")
    c_slc.font = f_res; c_slc.alignment = al_r; c_slc.fill = fill_res; c_slc.border = box_total; c_slc.number_format = '0.000000'
    ws_sol.cell(row=r, column=6, value="1/ч").font = f_bold; ws_sol.cell(row=r, column=6).alignment = al_c; ws_sol.cell(row=r, column=6).border = box_total
    ref_lc_12 = f"E{r}"
    results_map["t12_lc"] = ref_lc_12
    
    r += 2
    ws_sol.cell(row=r, column=2, value="2. Расчет среднего времени безотказной работы системы").font = f_bold
    r += 1
    for ci, h in enumerate(["Искомый показатель", "Матем. формула", "Формула в Excel", "Значение", "Размерность"], start=2):
        cell = ws_sol.cell(row=r, column=ci, value=h)
        cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
        cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws_sol.row_dimensions[r].height = 18
    
    r += 1
    ws_sol.row_dimensions[r].height = 19
    ws_sol.cell(row=r, column=2, value="Средняя наработка системы mtc").font = f_reg; ws_sol.cell(row=r, column=2).border = box_cell
    ws_sol.cell(row=r, column=3, value="1 / λc").font = f_math; ws_sol.cell(row=r, column=3).alignment = al_c; ws_sol.cell(row=r, column=3).border = box_cell
    ws_sol.cell(row=r, column=4, value=f"=1/{ref_lc_12}").font = f_math; ws_sol.cell(row=r, column=4).border = box_cell
    c_fmtc = ws_sol.cell(row=r, column=5, value=f"=1/{ref_lc_12}")
    c_fmtc.font = f_res; c_fmtc.fill = fill_res; c_fmtc.alignment = al_r; c_fmtc.border = box_res; c_fmtc.number_format = '0.00'
    ws_sol.cell(row=r, column=6, value="ч").font = f_reg; ws_sol.cell(row=r, column=6).alignment = al_c; ws_sol.cell(row=r, column=6).border = box_cell
    results_map["t12_mtc"] = f"E{r}"
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_sol.cell(row=r, column=2, value="Вывод: Интенсивность отказов системы λc = 0,024094 1/ч; средняя наработка системы mtc = 41,50 ч.").font = f_bold
    apply_style(ws_sol, r, 2, 7, fill=fill_ans, border=box_cell, al=al_l)
    ws_sol.row_dimensions[r].height = 20

    col_w_sol = {"A": 3, "B": 24, "C": 14, "D": 22, "E": 18, "F": 14, "G": 20}
    for col, width in col_w_sol.items():
        ws_sol.column_dimensions[col].width = width

    # ЛИСТ 1: ИТОГОВАЯ ТАБЛИЦА (СВОДКА)
    ws_sum = wb.create_sheet(title="Итоговая таблица", index=0)
    ws_sum.views.sheetView[0].showGridLines = True
    
    ws_sum.merge_cells('B2:G2')
    ws_sum['B2'] = "СВОДНЫЕ РЕЗУЛЬТАТЫ ПРАКТИЧЕСКОЙ РАБОТЫ № 2"
    ws_sum['B2'].font = f_title; ws_sum['B2'].alignment = al_c
    apply_style(ws_sum, 2, 2, 7, fill=fill_dark)
    ws_sum.row_dimensions[2].height = 24
    
    ws_sum.merge_cells('B3:G3')
    ws_sum['B3'] = "Студент: Гайнетдинов  |  Вариант / Задания № 6 – 12"
    ws_sum['B3'].font = f_subtitle; ws_sum['B3'].alignment = al_c
    apply_style(ws_sum, 3, 2, 7, fill=fill_med)
    ws_sum.row_dimensions[3].height = 18
    
    ws_sum['B5'] = "Таблица итоговых ответов по аналитическому расчету надежности:"
    ws_sum['B5'].font = f_sec
    ws_sum.row_dimensions[5].height = 20
    
    sum_heads_gay = ["№ задачи", "Краткое условие", "Параметр", "Формула", "Результат", "Ед. изм."]
    for i, h in enumerate(sum_heads_gay, start=2):
        cell = ws_sum.cell(row=6, column=i, value=h)
        cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
        cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws_sum.row_dimensions[6].height = 22

    sum_rows_gay = [
        ("Задание 6", "L1=460 ч, t=800 ч. Интенсивность отказов", "λ(800)", "λ = 1 / L1", f"='Решения (№6-12)'!{results_map['t6_l']}", "1/ч", "0.000000"),
        ("Задание 6", "L1=460 ч, t=800 ч. Вероятность безотказной работы", "P(800)", "P = exp(-λ·t)", f"='Решения (№6-12)'!{results_map['t6_p']}", "—", "0.0000"),
        ("Задание 6", "L1=460 ч, t=800 ч. Частота отказов", "f(800)", "f = λ · P(t)", f"='Решения (№6-12)'!{results_map['t6_f']}", "1/ч", "0.000000"),
        ("Задание 7", "6000 элементов, λср=0.2·10⁻⁶ ч⁻¹, T=50 ч", "Pc(50)", "Pc = exp(-λc·T)", f"='Решения (№6-12)'!{results_map['t7_pc']}", "—", "0.0000"),
        ("Задание 7", "6000 элементов, λср=0.2·10⁻⁶ ч⁻¹, T=50 ч", "mtc", "mtc = 1 / λc", f"='Решения (№6-12)'!{results_map['t7_mtc']}", "ч", "0.00"),
        ("Задание 8", "5000 элементов, λср=0.3·10⁻⁶ ч⁻¹, T=120 ч", "Pc(120)", "Pc = exp(-λc·T)", f"='Решения (№6-12)'!{results_map['t8_pc']}", "—", "0.0000"),
        ("Задание 8", "5000 элементов, λср=0.3·10⁻⁶ ч⁻¹, T=120 ч", "mtc", "mtc = 1 / λc", f"='Решения (№6-12)'!{results_map['t8_mtc']}", "ч", "0.00"),
        ("Задание 9", "200000 элементов, λср=0.3·10⁻⁶ ч⁻¹, сутки (t=24 ч)", "Pc(24)", "Pc = exp(-λc·t)", f"='Решения (№6-12)'!{results_map['t9_pc']}", "—", "0.0000"),
        ("Задание 9", "200000 элементов, λср=0.3·10⁻⁶ ч⁻¹, сутки (t=24 ч)", "mtc", "mtc = 1 / λc", f"='Решения (№6-12)'!{results_map['t9_mtc']}", "ч", "0.00"),
        ("Задание 10", "12600 элементов, λср=0.32·10⁻⁶ ч⁻¹, t=50 ч", "Pc(50)", "Pc = exp(-λc·t)", f"='Решения (№6-12)'!{results_map['t10_pc']}", "—", "0.0000"),
        ("Задание 10", "12600 элементов, λср=0.32·10⁻⁶ ч⁻¹, t=50 ч", "qc(50)", "qc = 1 - Pc", f"='Решения (№6-12)'!{results_map['t10_qc']}", "—", "0.0000"),
        ("Задание 10", "12600 элементов, λср=0.32·10⁻⁶ ч⁻¹, t=50 ч", "fc(50)", "fc = λc · Pc", f"='Решения (№6-12)'!{results_map['t10_fc']}", "1/ч", "0.000000"),
        ("Задание 10", "12600 элементов, λср=0.32·10⁻⁶ ч⁻¹, t=50 ч", "mtc", "mtc = 1 / λc", f"='Решения (№6-12)'!{results_map['t10_mtc']}", "ч", "0.00"),
        ("Задание 11", "Прибор с λ=0.1·10⁻⁶ ч⁻¹, t=180 ч", "p(180)", "p = exp(-λ·t)", f"='Решения (№6-12)'!{results_map['t11_p']}", "—", "0.000000"),
        ("Задание 11", "Прибор с λ=0.1·10⁻⁶ ч⁻¹, t=180 ч", "q(180)", "q = 1 - p", f"='Решения (№6-12)'!{results_map['t11_q']}", "—", "0.000000"),
        ("Задание 11", "Прибор с λ=0.1·10⁻⁶ ч⁻¹, t=180 ч", "f(180)", "f = λ · p", f"='Решения (№6-12)'!{results_map['t11_f']}", "1/ч", "0.00E+00"),
        ("Задание 11", "Прибор с λ=0.1·10⁻⁶ ч⁻¹, t=180 ч", "mt", "mt = 1 / λ", f"='Решения (№6-12)'!{results_map['t11_mt']}", "ч", "#,##0"),
        ("Задание 12", "Система из 5 приборов (mti = {83, 220, 280, 400, 700})", "λc", "λc = Σ (1/mti)", f"='Решения (№6-12)'!{results_map['t12_lc']}", "1/ч", "0.000000"),
        ("Задание 12", "Система из 5 приборов (mti = {83, 220, 280, 400, 700})", "mtc", "mtc = 1 / λc", f"='Решения (№6-12)'!{results_map['t12_mtc']}", "ч", "0.00")
    ]
    
    for idx, s_row in enumerate(sum_rows_gay, start=7):
        ws_sum.row_dimensions[idx].height = 19
        bg = fill_zebra if idx % 2 == 0 else PatternFill(fill_type=None)
        
        c1 = ws_sum.cell(row=idx, column=2, value=s_row[0]); c1.font = f_bold; c1.alignment = al_c; c1.border = box_cell; c1.fill = bg
        c2 = ws_sum.cell(row=idx, column=3, value=s_row[1]); c2.font = f_reg; c2.alignment = al_l; c2.border = box_cell; c2.fill = bg
        c3 = ws_sum.cell(row=idx, column=4, value=s_row[2]); c3.font = f_bold; c3.alignment = al_c; c3.border = box_cell; c3.fill = bg
        c4 = ws_sum.cell(row=idx, column=5, value=s_row[3]); c4.font = f_math; c4.alignment = al_l; c4.border = box_cell; c4.fill = bg
        c5 = ws_sum.cell(row=idx, column=6, value=s_row[4]); c5.font = f_res; c5.fill = fill_res; c5.alignment = al_r; c5.border = box_res
        c5.number_format = s_row[6]
        c6 = ws_sum.cell(row=idx, column=7, value=s_row[5]); c6.font = f_reg; c6.alignment = al_c; c6.border = box_cell; c6.fill = bg

    col_w_sum = {"A": 3, "B": 14, "C": 42, "D": 14, "E": 22, "F": 16, "G": 12}
    for col, width in col_w_sum.items():
        ws_sum.column_dimensions[col].width = width

    fname = "Практика_2_Гайнетдинов.xlsx"
    wb.save(fname)
    print(f"Created: {fname}")


if __name__ == '__main__':
    build_garifullin_file_pr2()
    build_gaynetdinov_file_pr2()

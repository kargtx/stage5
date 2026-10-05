import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# FILE 1: ПРАКТИКА_ГАРИФУЛЛИН.xlsx
# Стиль: Изумрудно-зеленый (хвоя/мята), шрифт Calibri, 1 лист с мини-сводкой
# вверху и аккуратными решениями задач 6-15 внизу.
# =============================================================================

def build_garifullin_file():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Практика 1 (Задачи 6-15)"
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
    ws['B2'] = "ПРАКТИЧЕСКАЯ РАБОТА № 1"
    ws['B2'].font = f_title
    ws['B2'].alignment = al_c
    apply_style(2, 2, 6, fill=fill_main)
    ws.row_dimensions[2].height = 24
    
    ws.merge_cells('B3:F3')
    ws['B3'] = "Определение количественных характеристик надежности по статистическим данным"
    ws['B3'].font = f_subtitle
    ws['B3'].alignment = al_c
    apply_style(3, 2, 6, fill=fill_sub)
    ws.row_dimensions[3].height = 18
    
    ws.merge_cells('B4:F4')
    ws['B4'] = "Выполнил: студент Гарифуллин  |  Дисциплина: Надежность автоматизированных систем"
    ws['B4'].font = Font(name=FONT, size=9, italic=True, color='E8F5E9')
    ws['B4'].alignment = al_c
    apply_style(4, 2, 6, fill=fill_sub)
    ws.row_dimensions[4].height = 18

    # Сводная таблица ответов наверху
    ws['B6'] = "Сводные результаты решения задач (№ 6 – 15)"
    ws['B6'].font = f_sec
    ws.row_dimensions[6].height = 20
    
    sum_heads = ["№ задачи", "Краткое описание", "Параметр", "Формула", "Результат", "Ед. изм."]
    for i, h in enumerate(sum_heads, start=2):
        cell = ws.cell(row=7, column=i, value=h)
        cell.font = f_th
        cell.fill = fill_th
        cell.alignment = al_c
        cell.border = box_cell
    ws.row_dimensions[7].height = 22

    # Резервируем строки 8-22 под сводку, заполним ссылки после генерации задач
    sum_start_row = 8
    
    # Генерация задач 6-15, начиная с строки 25
    cur_r = 25
    results_map = {}
    
    def add_task(title, cond, inputs, calcs, note, task_id):
        nonlocal cur_r
        r = cur_r
        task_start = r
        
        # Шапка задачи
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.cell(row=r, column=2, value=title).font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
        apply_style(r, 2, 6, fill=fill_sub, al=al_l)
        ws.row_dimensions[r].height = 20
        
        # Условие
        r += 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=6)
        c_c = ws.cell(row=r, column=2, value=f"Условие: {cond}")
        c_c.font = f_cond
        c_c.alignment = al_l_wrap
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
            cell.fill = fill_th
            cell.alignment = al_c
            cell.border = box_cell
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
            if isinstance(p_val, (int, float)): c3.number_format = '#,##0.####'
            c4 = ws.cell(row=r, column=5, value=p_unit); c4.font = f_reg; c4.alignment = al_c; c4.border = box_cell
            
            inp_cells[p_sym] = f"D{r}"
            inp_cells[p_sym.replace('(', '_').replace(')', '')] = f"D{r}"
        
        # Расчет
        r += 2
        ws.cell(row=r, column=2, value="Расчет:").font = f_bold
        r += 1
        for ci, h in enumerate(["Показатель", "Формула", "Расчет в Excel", "Значение", "Ед. изм."], start=2):
            cell = ws.cell(row=r, column=ci, value=h)
            cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
            cell.fill = fill_th
            cell.alignment = al_c
            cell.border = box_cell
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

    # Добавление задач
    add_task("Задача 6. Интенсивность и частота отказов", 
             "На испытание поставлено 100 изделий. За 4000 ч отказало 50 изделий. За интервал 4000-4100 ч отказало еще 20 изделий. Найти f(t), λ(t) при t = 4000 ч.",
             [("Число изделий", "N", 100, "шт."), ("Время t", "t", 4000, "ч"), ("Отказов к моменту t", "n_отк", 50, "шт."),
              ("Интервал времени", "dt", 100, "ч"), ("Отказов за интервал dt", "dn", 20, "шт.")],
             [("Работоспособно к 4000 ч", "n(t) = N - n_отк", "={N}-{n_отк}", "шт.", "0", "nt"),
              ("Частота отказов f(4000)", "f = dn / (N · dt)", "={dn}/({N}*{dt})", "1/ч", "0.00000", "f"),
              ("Интенсивность отказов λ(4000)", "λ = dn / (n(t) · dt)", "={dn}/({nt}*{dt})", "1/ч", "0.00000", "l")],
             "f(4000) = 0,002 1/ч; λ(4000) = 0,004 1/ч.", "t6")

    add_task("Задача 7. Вероятность безотказной работы и отказа", 
             "На испытание поставлено 100 изделий. За 4000 ч отказало 50 изделий. Найти p(t) и q(t) при t = 4000 ч.",
             [("Число изделий", "N", 100, "шт."), ("Время t", "t", 4000, "ч"), ("Отказов к моменту t", "n_отк", 50, "шт.")],
             [("Работоспособно к 4000 ч", "n(t) = N - n_отк", "={N}-{n_отк}", "шт.", "0", "nt"),
              ("Вероятность безотказной работы", "p = n(t) / N", "={nt}/{N}", "—", "0.0000", "p"),
              ("Вероятность отказа", "q = 1 - p", "=1-{p}", "—", "0.0000", "q")],
             "p(4000) = 0,50; q(4000) = 0,50.", "t7")

    add_task("Задача 8. Надежность гироскопов", 
             "В течение 1000 ч из 10 гироскопов отказало 2. За интервал 1000-1100 ч отказал еще 1 гироскоп. Найти f(t), λ(t) при t = 1000 ч.",
             [("Число гироскопов", "N", 10, "шт."), ("Время t", "t", 1000, "ч"), ("Отказов к моменту t", "n_отк", 2, "шт."),
              ("Интервал времени", "dt", 100, "ч"), ("Отказов за интервал dt", "dn", 1, "шт.")],
             [("Работоспособно к 1000 ч", "n(t) = N - n_отк", "={N}-{n_отк}", "шт.", "0", "nt"),
              ("Частота отказов f(1000)", "f = dn / (N · dt)", "={dn}/({N}*{dt})", "1/ч", "0.00000", "f"),
              ("Интенсивность отказов λ(1000)", "λ = dn / (n(t) · dt)", "={dn}/({nt}*{dt})", "1/ч", "0.00000", "l")],
             "f(1000) = 0,001 1/ч; λ(1000) = 0,00125 1/ч.", "t8")

    add_task("Задача 9. Надежность ламп", 
             "На испытание поставлено 1000 ламп. За первые 3000 ч отказало 80 ламп, за 3000-4000 ч отказало еще 50 ламп. Найти p(t), q(t) при t = 4000 ч.",
             [("Число ламп", "N", 1000, "шт."), ("Отказов за 3000 ч", "n_отк1", 80, "шт."), ("Отказов за 3000-4000 ч", "dn", 50, "шт.")],
             [("Сумма отказов к 4000 ч", "n_отк = n_отк1 + dn", "={n_отк1}+{dn}", "шт.", "0", "notk"),
              ("Работоспособно к 4000 ч", "n(4000) = N - n_отк", "={N}-{notk}", "шт.", "0", "nt"),
              ("Вероятность безотказной работы", "p = n(4000) / N", "={nt}/{N}", "—", "0.0000", "p"),
              ("Вероятность отказа", "q = 1 - p", "=1-{p}", "—", "0.0000", "q")],
             "p(4000) = 0,87; q(4000) = 0,13.", "t9")

    add_task("Задача 10. Характеристики изделий (N=1000)", 
             "На испытание поставлено 1000 изделий. К 1300 ч вышло из строя 288 изделий. За 1300-1400 ч вышло еще 13 изделий. Вычислить p(1300), p(1400), f(1300), λ(1300).",
             [("Число изделий", "N", 1000, "шт."), ("Время t1", "t1", 1300, "ч"), ("Отказов к t1", "n_отк1", 288, "шт."),
              ("Интервал dt", "dt", 100, "ч"), ("Отказов за dt", "dn", 13, "шт.")],
             [("Работоспособно к 1300 ч", "n(1300) = N - n_отк1", "={N}-{n_отк1}", "шт.", "0", "nt1"),
              ("Работоспособно к 1400 ч", "n(1400) = n(1300) - dn", "={nt1}-{dn}", "шт.", "0", "nt2"),
              ("Вероятность p(1300)", "p(1300) = n(1300) / N", "={nt1}/{N}", "—", "0.0000", "p1"),
              ("Вероятность p(1400)", "p(1400) = n(1400) / N", "={nt2}/{N}", "—", "0.0000", "p2"),
              ("Частота f(1300)", "f = dn / (N · dt)", "={dn}/({N}*{dt})", "1/ч", "0.000000", "f"),
              ("Интенсивность λ(1300)", "λ = dn / (n(1300) · dt)", "={dn}/({nt1}*{dt})", "1/ч", "0.000000", "l")],
             "p(1300) = 0,712; p(1400) = 0,699; f(1300) = 0,00013 1/ч; λ(1300) = 0,000183 1/ч.", "t10")

    add_task("Задача 11. Характеристики изделий (N=45)", 
             "На испытание поставлено 45 изделий. К 60 ч вышло из строя 35 изделий. За 60-65 ч вышло еще 3 изделия. Вычислить p(60), p(65), f(60), λ(60).",
             [("Число изделий", "N", 45, "шт."), ("Время t1", "t1", 60, "ч"), ("Отказов к t1", "n_отк1", 35, "шт."),
              ("Интервал dt", "dt", 5, "ч"), ("Отказов за dt", "dn", 3, "шт.")],
             [("Работоспособно к 60 ч", "n(60) = N - n_отк1", "={N}-{n_отк1}", "шт.", "0", "nt1"),
              ("Работоспособно к 65 ч", "n(65) = n(60) - dn", "={nt1}-{dn}", "шт.", "0", "nt2"),
              ("Вероятность p(60)", "p(60) = n(60) / N", "={nt1}/{N}", "—", "0.0000", "p1"),
              ("Вероятность p(65)", "p(65) = n(65) / N", "={nt2}/{N}", "—", "0.0000", "p2"),
              ("Частота f(60)", "f = dn / (N · dt)", "={dn}/({N}*{dt})", "1/ч", "0.000000", "f"),
              ("Интенсивность λ(60)", "λ = dn / (n(60) · dt)", "={dn}/({nt1}*{dt})", "1/ч", "0.000000", "l")],
             "p(60) = 0,2222; p(65) = 0,1556; f(60) = 0,013333 1/ч; λ(60) = 0,060000 1/ч.", "t11")

    # Задача 12 (Интервальный ряд)
    r = cur_r
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2, value="Задача 12. Среднее время безотказной работы по группированным данным").font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    apply_style(r, 2, 6, fill=fill_sub, al=al_l)
    ws.row_dimensions[r].height = 20
    
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=6)
    ws.cell(row=r, column=2, value="Условие: Наблюдение за 45 образцами РЭО (после 80-часовой приработки). Данные по интервалам до первого отказа сведены в таблицу. Найти mt*.").font = f_cond
    ws.cell(row=r, column=2).alignment = al_l_wrap
    apply_style(r, 2, 6, fill=fill_card)
    apply_style(r+1, 2, 6, fill=fill_card)
    ws.row_dimensions[r].height = 16
    ws.row_dimensions[r+1].height = 16
    r += 1
    
    r += 1
    ws.cell(row=r, column=2, value="Расчетная таблица интервалов:").font = f_bold
    r += 1
    for ci, h in enumerate(["№", "Интервал Δti, ч", "Середина tср.i, ч", "Отказов ni, шт.", "ni · tср.i, ч"], start=2):
        cell = ws.cell(row=r, column=ci, value=h)
        cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
        cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws.row_dimensions[r].height = 19
    
    int_data = [(1, "0-10", 5, 19), (2, "10-20", 15, 13), (3, "20-30", 25, 8),
                (4, "30-40", 35, 3), (5, "40-50", 45, 0), (6, "50-60", 55, 1), (7, "60-70", 65, 1)]
    st_t12 = r + 1
    for item in int_data:
        r += 1
        ws.row_dimensions[r].height = 18
        ws.cell(row=r, column=2, value=item[0]).alignment = al_c; ws.cell(row=r, column=2).font = f_reg; ws.cell(row=r, column=2).border = box_cell
        ws.cell(row=r, column=3, value=item[1]).alignment = al_c; ws.cell(row=r, column=3).font = f_reg; ws.cell(row=r, column=3).border = box_cell
        ws.cell(row=r, column=4, value=item[2]).alignment = al_r; ws.cell(row=r, column=4).font = f_reg; ws.cell(row=r, column=4).border = box_cell
        ws.cell(row=r, column=5, value=item[3]).alignment = al_r; ws.cell(row=r, column=5).font = f_bold; ws.cell(row=r, column=5).border = box_cell
        c5 = ws.cell(row=r, column=6, value=f"=D{r}*E{r}"); c5.font = f_bold; c5.alignment = al_r; c5.border = box_cell
        c5.number_format = '#,##0.0'
    end_t12 = r
    
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws.cell(row=r, column=2, value="Сумма:").font = f_bold; ws.cell(row=r, column=2).alignment = al_r
    apply_style(r, 2, 4, fill=fill_card, border=box_total)
    c_s_n = ws.cell(row=r, column=5, value=f"=SUM(E{st_t12}:E{end_t12})"); c_s_n.font = f_bold; c_s_n.alignment = al_r; c_s_n.fill = fill_card; c_s_n.border = box_total
    c_s_prod = ws.cell(row=r, column=6, value=f"=SUM(F{st_t12}:F{end_t12})"); c_s_prod.font = f_bold; c_s_prod.alignment = al_r; c_s_prod.fill = fill_card; c_s_prod.border = box_total
    r_sum12 = r
    
    r += 2
    ws.cell(row=r, column=2, value="Среднее время безотказной работы mt* = Σ(ni·tср.i) / N:").font = f_reg
    c_mt12 = ws.cell(row=r, column=5, value=f"=F{r_sum12}/E{r_sum12}")
    c_mt12.font = f_res; c_mt12.fill = fill_res; c_mt12.border = box_res; c_mt12.alignment = al_r
    c_mt12.number_format = '0.00'
    ws.cell(row=r, column=6, value="ч").font = f_reg; ws.cell(row=r, column=6).alignment = al_c
    results_map["t12_mt"] = f"E{r}"
    
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2, value="Ответ: mt* = 15,89 ч (с учетом приработки 80 ч: T = 95,89 ч).").font = f_bold
    apply_style(r, 2, 6, fill=fill_ans, border=box_cell, al=al_l)
    cur_r = r + 2

    # Задача 13 (Выборка 8 изделий)
    r = cur_r
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2, value="Задача 13. Среднее время и дисперсия по наработкам 8 изделий").font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    apply_style(r, 2, 6, fill=fill_sub, al=al_l)
    ws.row_dimensions[r].height = 20
    
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=6)
    ws.cell(row=r, column=2, value="Условие: На испытание поставлено 8 изделий: ti = {560, 700, 800, 650, 580, 760, 920, 850} ч. Найти среднее время mt* и дисперсию.").font = f_cond
    ws.cell(row=r, column=2).alignment = al_l_wrap
    apply_style(r, 2, 6, fill=fill_card)
    apply_style(r+1, 2, 6, fill=fill_card)
    ws.row_dimensions[r].height = 16
    ws.row_dimensions[r+1].height = 16
    r += 1
    
    r += 1
    ws.cell(row=r, column=2, value="Выборка наработок изделий:").font = f_bold
    r += 1
    for ci, h in enumerate(["№ изделия", "Обозначение", "Время работы ti, ч", ""], start=2):
        if h:
            cell = ws.cell(row=r, column=ci, value=h)
            cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
            cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws.row_dimensions[r].height = 19
    
    t13_data = [560, 700, 800, 650, 580, 760, 920, 850]
    st_t13 = r + 1
    for i, val in enumerate(t13_data, start=1):
        r += 1
        ws.row_dimensions[r].height = 18
        ws.cell(row=r, column=2, value=i).alignment = al_c; ws.cell(row=r, column=2).font = f_reg; ws.cell(row=r, column=2).border = box_cell
        ws.cell(row=r, column=3, value=f"t{i}").alignment = al_c; ws.cell(row=r, column=3).font = f_bold; ws.cell(row=r, column=3).border = box_cell
        c_v = ws.cell(row=r, column=4, value=val); c_v.alignment = al_r; c_v.font = f_bold; c_v.border = box_cell
    end_t13 = r
    
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.cell(row=r, column=2, value="Сумма Σti:").font = f_bold; ws.cell(row=r, column=2).alignment = al_r
    apply_style(r, 2, 3, fill=fill_card, border=box_total)
    c_sum13 = ws.cell(row=r, column=4, value=f"=SUM(D{st_t13}:D{end_t13})"); c_sum13.font = f_bold; c_sum13.alignment = al_r; c_sum13.fill = fill_card; c_sum13.border = box_total
    
    r += 2
    ws.cell(row=r, column=2, value="Среднее время mt* = Σti / N:").font = f_reg
    c_mt13 = ws.cell(row=r, column=5, value=f"=AVERAGE(D{st_t13}:D{end_t13})")
    c_mt13.font = f_res; c_mt13.fill = fill_res; c_mt13.border = box_res; c_mt13.alignment = al_r; c_mt13.number_format = '0.00'
    ws.cell(row=r, column=6, value="ч").font = f_reg; ws.cell(row=r, column=6).alignment = al_c
    results_map["t13_mt"] = f"E{r}"
    
    r += 1
    ws.cell(row=r, column=2, value="Выборочная дисперсия Dt*:").font = f_reg
    c_var13 = ws.cell(row=r, column=5, value=f"=VAR.S(D{st_t13}:D{end_t13})")
    c_var13.font = f_bold; c_var13.border = box_cell; c_var13.alignment = al_r; c_var13.number_format = '0.00'
    ws.cell(row=r, column=6, value="ч²").font = f_reg; ws.cell(row=r, column=6).alignment = al_c
    
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2, value="Ответ: mt* = 727,50 ч; Dt* = 16307,14 ч² (СКО σ = 127,70 ч).").font = f_bold
    apply_style(r, 2, 6, fill=fill_ans, border=box_cell, al=al_l)
    cur_r = r + 2

    # Задача 14 (Время восстановления)
    r = cur_r
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2, value="Задача 14. Среднее время восстановления аппаратуры").font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    apply_style(r, 2, 6, fill=fill_sub, al=al_l)
    ws.row_dimensions[r].height = 20
    
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=6)
    ws.cell(row=r, column=2, value="Условие: Зарегистрировано 6 отказов. Время восстановления: 15, 20, 10, 28, 22, 30 мин. Найти среднее время восстановления mt.").font = f_cond
    ws.cell(row=r, column=2).alignment = al_l_wrap
    apply_style(r, 2, 6, fill=fill_card)
    apply_style(r+1, 2, 6, fill=fill_card)
    ws.row_dimensions[r].height = 16
    ws.row_dimensions[r+1].height = 16
    r += 1
    
    r += 1
    ws.cell(row=r, column=2, value="Данные по отказам:").font = f_bold
    r += 1
    for ci, h in enumerate(["№ отказа", "Обозначение", "Время восстановления tв.i, мин", ""], start=2):
        if h:
            cell = ws.cell(row=r, column=ci, value=h)
            cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
            cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws.row_dimensions[r].height = 19
    
    t14_data = [15, 20, 10, 28, 22, 30]
    st_t14 = r + 1
    for i, val in enumerate(t14_data, start=1):
        r += 1
        ws.row_dimensions[r].height = 18
        ws.cell(row=r, column=2, value=i).alignment = al_c; ws.cell(row=r, column=2).font = f_reg; ws.cell(row=r, column=2).border = box_cell
        ws.cell(row=r, column=3, value=f"t{i}").alignment = al_c; ws.cell(row=r, column=3).font = f_bold; ws.cell(row=r, column=3).border = box_cell
        ws.cell(row=r, column=4, value=val).alignment = al_r; ws.cell(row=r, column=4).font = f_bold; ws.cell(row=r, column=4).border = box_cell
    end_t14 = r
    
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.cell(row=r, column=2, value="Сумма времени:").font = f_bold; ws.cell(row=r, column=2).alignment = al_r
    apply_style(r, 2, 3, fill=fill_card, border=box_total)
    ws.cell(row=r, column=4, value=f"=SUM(D{st_t14}:D{end_t14})").font = f_bold; ws.cell(row=r, column=4).alignment = al_r; ws.cell(row=r, column=4).fill = fill_card; ws.cell(row=r, column=4).border = box_total
    
    r += 2
    ws.cell(row=r, column=2, value="Среднее время восстановления mt = Σtв.i / N:").font = f_reg
    c_mt14 = ws.cell(row=r, column=5, value=f"=AVERAGE(D{st_t14}:D{end_t14})")
    c_mt14.font = f_res; c_mt14.fill = fill_res; c_mt14.border = box_res; c_mt14.alignment = al_r; c_mt14.number_format = '0.00'
    ws.cell(row=r, column=6, value="мин").font = f_reg; ws.cell(row=r, column=6).alignment = al_c
    results_map["t14_mt"] = f"E{r}"
    
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2, value="Ответ: Среднее время восстановления аппаратуры mt = 20,83 мин (0,35 ч).").font = f_bold
    apply_style(r, 2, 6, fill=fill_ans, border=box_cell, al=al_l)
    cur_r = r + 2

    # Задача 15
    add_task("Задача 15. Характеристики изделий при t = 11000 ч и 12000 ч", 
             "На испытание поставлено 1000 изделий. К 11000 ч вышло из строя 410 изделий. За 11000-12000 ч вышло еще 40 изделий. Найти p(11000), p(12000), f(11000), λ(11000).",
             [("Число изделий", "N", 1000, "шт."), ("Время t1", "t1", 11000, "ч"), ("Отказов к t1", "n_отк1", 410, "шт."),
              ("Интервал dt", "dt", 1000, "ч"), ("Отказов за dt", "dn", 40, "шт.")],
             [("Работоспособно к 11000 ч", "n(11000) = N - n_отк1", "={N}-{n_отк1}", "шт.", "0", "nt1"),
              ("Работоспособно к 12000 ч", "n(12000) = n(11000) - dn", "={nt1}-{dn}", "шт.", "0", "nt2"),
              ("Вероятность p(11000)", "p(11000) = n(11000) / N", "={nt1}/{N}", "—", "0.0000", "p1"),
              ("Вероятность p(12000)", "p(12000) = n(12000) / N", "={nt2}/{N}", "—", "0.0000", "p2"),
              ("Частота f(11000)", "f = dn / (N · dt)", "={dn}/({N}*{dt})", "1/ч", "0.000000", "f"),
              ("Интенсивность λ(11000)", "λ = dn / (n(11000) · dt)", "={dn}/({nt1}*{dt})", "1/ч", "0.000000", "l")],
             "p(11000) = 0,5900; p(12000) = 0,5500; f(11000) = 4,00·10⁻⁵ 1/ч; λ(11000) = 6,78·10⁻⁵ 1/ч.", "t15")

    # Теперь заполняем сводную таблицу вверху (строки 8..21)
    sum_rows = [
        ("Задача 6", "100 изделий, t=4000 ч", "Частота f(4000)", "f = Δn/(N·Δt)", f"={results_map['t6_f']}", "1/ч", "0.00000"),
        ("Задача 6", "100 изделий, t=4000 ч", "Интенсивность λ(4000)", "λ = Δn/(n(t)·Δt)", f"={results_map['t6_l']}", "1/ч", "0.00000"),
        ("Задача 7", "100 изделий, t=4000 ч", "Вероятность p(4000)", "p = n(t)/N", f"={results_map['t7_p']}", "—", "0.0000"),
        ("Задача 7", "100 изделий, t=4000 ч", "Вероятность q(4000)", "q = 1 - p", f"={results_map['t7_q']}", "—", "0.0000"),
        ("Задача 8", "10 гироскопов, t=1000 ч", "Частота f(1000)", "f = Δn/(N·Δt)", f"={results_map['t8_f']}", "1/ч", "0.00000"),
        ("Задача 8", "10 гироскопов, t=1000 ч", "Интенсивность λ(1000)", "λ = Δn/(n(t)·Δt)", f"={results_map['t8_l']}", "1/ч", "0.00000"),
        ("Задача 9", "1000 ламп, t=4000 ч", "Вероятность p(4000)", "p = n(4000)/N", f"={results_map['t9_p']}", "—", "0.0000"),
        ("Задача 9", "1000 ламп, t=4000 ч", "Вероятность q(4000)", "q = 1 - p", f"={results_map['t9_q']}", "—", "0.0000"),
        ("Задача 10", "1000 изделий, t=1300 ч", "Вероятности p, f, λ", "p(1300); p(1400); f; λ", f"={results_map['t10_p1']}", "—", "0.0000"),
        ("Задача 11", "45 изделий, t=60 ч", "Вероятности p, f, λ", "p(60); p(65); f; λ", f"={results_map['t11_p1']}", "—", "0.0000"),
        ("Задача 12", "45 образцов РЭО", "Среднее время mt*", "mt* = Σ(ni·tср)/N", f"={results_map['t12_mt']}", "ч", "0.00"),
        ("Задача 13", "8 изделий (наработки)", "Среднее время mt*", "mt* = Σti/N", f"={results_map['t13_mt']}", "ч", "0.00"),
        ("Задача 14", "6 отказов (восстановление)", "Среднее время mt", "mt = Σtв/N", f"={results_map['t14_mt']}", "мин", "0.00"),
        ("Задача 15", "1000 изделий, t=11000 ч", "Вероятность p(11000)", "p = n(11000)/N", f"={results_map['t15_p1']}", "—", "0.0000"),
    ]
    
    for idx, s_row in enumerate(sum_rows, start=sum_start_row):
        ws.row_dimensions[idx].height = 19
        bg = fill_zebra if idx % 2 == 0 else PatternFill(fill_type=None)
        
        c1 = ws.cell(row=idx, column=2, value=s_row[0]); c1.font = f_bold; c1.alignment = al_c; c1.border = box_cell; c1.fill = bg
        c2 = ws.cell(row=idx, column=3, value=s_row[1]); c2.font = f_reg; c2.alignment = al_l; c2.border = box_cell; c2.fill = bg
        c3 = ws.cell(row=idx, column=4, value=s_row[2]); c3.font = f_bold; c3.alignment = al_l; c3.border = box_cell; c3.fill = bg
        c4 = ws.cell(row=idx, column=5, value=s_row[3]); c4.font = f_math; c4.alignment = al_l; c4.border = box_cell; c4.fill = bg
        c5 = ws.cell(row=idx, column=6, value=s_row[4]); c5.font = f_res; c5.fill = fill_res; c5.alignment = al_r; c5.border = box_res
        c5.number_format = s_row[6]
        c6 = ws.cell(row=idx, column=7, value=s_row[5]); c6.font = f_reg; c6.alignment = al_c; c6.border = box_cell; c6.fill = bg

    # Подгонка колонок
    col_w = {"A": 3, "B": 24, "C": 26, "D": 26, "E": 20, "F": 18, "G": 10}
    for col, width in col_w.items():
        ws.column_dimensions[col].width = width

    fname = "Практика_Гарифуллин.xlsx"
    wb.save(fname)
    print(f"Created: {fname}")


# =============================================================================
# FILE 2: ПРАКТИКА_ГАЙНЕТДИНОВ.xlsx
# Стиль: Графитово-стальной / Slate Gray, шрифт Arial, 2 листа:
# Лист 1: "Результаты" (сводная таблица ответов)
# Лист 2: "Решения (№6-15)" (подробные карточки)
# =============================================================================

def build_gaynetdinov_file():
    wb = openpyxl.Workbook()
    
    FONT = 'Arial'
    
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

    # ЛИСТ 2: РЕШЕНИЯ (№6-15)
    ws_sol = wb.active
    ws_sol.title = "Решения (№6-15)"
    ws_sol.views.sheetView[0].showGridLines = True
    
    # Шапка листа решений
    ws_sol.merge_cells('B2:G2')
    ws_sol['B2'] = "РАСЧЕТНАЯ ЧАСТЬ ПРАКТИЧЕСКОЙ РАБОТЫ № 1 (ЗАДАНИЯ 6 – 15)"
    ws_sol['B2'].font = f_title; ws_sol['B2'].alignment = al_c
    apply_style(ws_sol, 2, 2, 7, fill=fill_dark)
    ws_sol.row_dimensions[2].height = 24
    
    ws_sol.merge_cells('B3:G3')
    ws_sol['B3'] = "Оценка надежности изделий по статистике отказов  |  Студент: Гайнетдинов"
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
            if isinstance(p_val, (int, float)): c3.number_format = '#,##0.####'
            c4 = ws_sol.cell(row=r, column=5, value=p_unit); c4.font = f_reg; c4.alignment = al_c; c4.border = box_cell
            
            inp_cells[p_sym] = f"D{r}"
            inp_cells[p_sym.replace('(', '_').replace(')', '')] = f"D{r}"
        
        r += 2
        ws_sol.cell(row=r, column=2, value="2. Расчетные формулы и вычисления").font = f_bold
        r += 1
        for ci, h in enumerate(["Искомый показатель", "Обозначение", "Матем. формула", "Формула в Excel", "Результат", "Ед."], start=2):
            cell = ws_sol.cell(row=r, column=ci, value=h)
            cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
            cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
        ws_sol.row_dimensions[r].height = 18
        
        calc_cells = {}
        for row_c in calcs:
            r += 1
            ws_sol.row_dimensions[r].height = 19
            c_name, c_sym, c_math, c_template, c_unit, c_fmt, c_key = row_c
            formula = c_template.format(**inp_cells, **calc_cells)
            
            c1 = ws_sol.cell(row=r, column=2, value=c_name); c1.font = f_reg; c1.border = box_cell
            c2 = ws_sol.cell(row=r, column=3, value=c_sym); c2.font = f_bold; c2.alignment = al_c; c2.border = box_cell
            c3 = ws_sol.cell(row=r, column=4, value=c_math); c3.font = f_math; c3.alignment = al_c; c3.border = box_cell
            c4 = ws_sol.cell(row=r, column=5, value=formula); c4.font = f_math; c4.border = box_cell
            c5 = ws_sol.cell(row=r, column=6, value=formula); c5.font = f_res; c5.fill = fill_res; c5.alignment = al_r; c5.border = box_res
            c5.number_format = c_fmt
            c6 = ws_sol.cell(row=r, column=7, value=c_unit); c6.font = f_reg; c6.alignment = al_c; c6.border = box_cell
            
            calc_cells[c_key] = f"F{r}"
            results_map[f"{task_id}_{c_key}"] = f"F{r}"
        
        r += 1
        ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ans = ws_sol.cell(row=r, column=2, value=f"Итог: {note}")
        ans.font = f_bold
        apply_style(ws_sol, r, 2, 7, fill=fill_ans, border=box_cell, al=al_l)
        ws_sol.row_dimensions[r].height = 20
        
        cur_r = r + 2

    # Задания 6-11
    add_card("Задание №6. Интенсивность и частота отказов", 
             "На испытании N=100 изделий. К t=4000 ч отказало 50 шт. На участке 4000-4100 ч отказало 20 шт. Найти f(4000) и λ(4000).",
             [("Количество изделий", "N", 100, "шт."), ("Время t", "t", 4000, "ч"), ("Число отказов к моменту t", "notk", 50, "шт."),
              ("Период времени", "dt", 100, "ч"), ("Отказов за период", "dn", 20, "шт.")],
             [("Число исправных к t", "n(t)", "N - notk", "={N}-{notk}", "шт.", "0", "nt"),
              ("Частота отказов", "f(4000)", "dn / (N · dt)", "={dn}/({N}*{dt})", "1/ч", "0.00000", "f"),
              ("Интенсивность отказов", "λ(4000)", "dn / (n(t) · dt)", "={dn}/({nt}*{dt})", "1/ч", "0.00000", "l")],
             "f(4000) = 0,002 1/ч, λ(4000) = 0,004 1/ч.", "t6")

    add_card("Задание №7. Вероятность безотказной работы и вероятность отказа", 
             "На испытании N=100 изделий. К наработке t=4000 ч отказало 50 шт. Найти p(4000) и q(4000).",
             [("Количество изделий", "N", 100, "шт."), ("Наработка t", "t", 4000, "ч"), ("Число отказов", "notk", 50, "шт.")],
             [("Исправных изделий", "n(t)", "N - notk", "={N}-{notk}", "шт.", "0", "nt"),
              ("Вероятность безотказной работы", "p(4000)", "n(t) / N", "={nt}/{N}", "—", "0.0000", "p"),
              ("Вероятность отказа", "q(4000)", "1 - p", "=1-{p}", "—", "0.0000", "q")],
             "p(4000) = 0,5000; q(4000) = 0,5000.", "t7")

    add_card("Задание №8. Показатели безотказности гироскопов", 
             "За 1000 ч из N=10 гироскопов отказало 2. В интервале 1000-1100 ч отказал еще 1 гироскоп. Найти f(1000) и λ(1000).",
             [("Число гироскопов", "N", 10, "шт."), ("Время t", "t", 1000, "ч"), ("Число отказов к моменту t", "notk", 2, "шт."),
              ("Шаг по времени", "dt", 100, "ч"), ("Отказов за интервал", "dn", 1, "шт.")],
             [("Исправных гироскопов к t", "n(t)", "N - notk", "={N}-{notk}", "шт.", "0", "nt"),
              ("Частота отказов", "f(1000)", "dn / (N · dt)", "={dn}/({N}*{dt})", "1/ч", "0.00000", "f"),
              ("Интенсивность отказов", "λ(1000)", "dn / (n(t) · dt)", "={dn}/({nt}*{dt})", "1/ч", "0.00000", "l")],
             "f(1000) = 0,001 1/ч, λ(1000) = 0,00125 1/ч.", "t8")

    add_card("Задание №9. Испытание электронных ламп", 
             "Партия ламп N=1000. За 3000 ч отказало 80 ламп, за 3000-4000 ч еще 50 ламп. Найти p(4000) и q(4000).",
             [("Объем выборки ламп", "N", 1000, "шт."), ("Отказов до 3000 ч", "notk1", 80, "шт."), ("Отказов за 3000-4000 ч", "dn", 50, "шт.")],
             [("Всего отказов к 4000 ч", "notk_tot", "notk1 + dn", "={notk1}+{dn}", "шт.", "0", "notk"),
              ("Не отказавших ламп", "n(4000)", "N - notk_tot", "={N}-{notk}", "шт.", "0", "nt"),
              ("Вероятность безотказной работы", "p(4000)", "n(4000) / N", "={nt}/{N}", "—", "0.0000", "p"),
              ("Вероятность отказа", "q(4000)", "1 - p", "=1-{p}", "—", "0.0000", "q")],
             "p(4000) = 0,8700; q(4000) = 0,1300.", "t9")

    add_card("Задание №10. Характеристики изделий (N=1000)", 
             "Поставлено 1000 изделий. К 1300 ч вышло из строя 288 шт., за 1300-1400 ч еще 13 шт. Найти p(1300), p(1400), f(1300), λ(1300).",
             [("Общее число изделий", "N", 1000, "шт."), ("Время t1", "t1", 1300, "ч"), ("Отказов к t1", "notk1", 288, "шт."),
              ("Интервал dt", "dt", 100, "ч"), ("Отказов в интервале", "dn", 13, "шт.")],
             [("Исправных изделий к 1300 ч", "n(1300)", "N - notk1", "={N}-{notk1}", "шт.", "0", "nt1"),
              ("Исправных изделий к 1400 ч", "n(1400)", "n(1300) - dn", "={nt1}-{dn}", "шт.", "0", "nt2"),
              ("Вероятность p(1300)", "p(1300)", "n(1300) / N", "={nt1}/{N}", "—", "0.0000", "p1"),
              ("Вероятность p(1400)", "p(1400)", "n(1400) / N", "={nt2}/{N}", "—", "0.0000", "p2"),
              ("Частота отказов f(1300)", "f(1300)", "dn / (N · dt)", "={dn}/({N}*{dt})", "1/ч", "0.000000", "f"),
              ("Интенсивность отказов λ(1300)", "λ(1300)", "dn / (n(1300) · dt)", "={dn}/({nt1}*{dt})", "1/ч", "0.000000", "l")],
             "p(1300)=0,7120; p(1400)=0,6990; f=1,30·10⁻⁴ 1/ч; λ=1,83·10⁻⁴ 1/ч.", "t10")

    add_card("Задание №11. Характеристики изделий (N=45)", 
             "Поставлено 45 изделий. К 60 ч вышло из строя 35 шт., за 60-65 ч еще 3 шт. Найти p(60), p(65), f(60), λ(60).",
             [("Общее число изделий", "N", 45, "шт."), ("Время t1", "t1", 60, "ч"), ("Отказов к t1", "notk1", 35, "шт."),
              ("Интервал dt", "dt", 5, "ч"), ("Отказов в интервале", "dn", 3, "шт.")],
             [("Исправных изделий к 60 ч", "n(60)", "N - notk1", "={N}-{notk1}", "шт.", "0", "nt1"),
              ("Исправных изделий к 65 ч", "n(65)", "n(60) - dn", "={nt1}-{dn}", "шт.", "0", "nt2"),
              ("Вероятность p(60)", "p(60)", "n(60) / N", "={nt1}/{N}", "—", "0.0000", "p1"),
              ("Вероятность p(65)", "p(65)", "n(65) / N", "={nt2}/{N}", "—", "0.0000", "p2"),
              ("Частота отказов f(60)", "f(60)", "dn / (N · dt)", "={dn}/({N}*{dt})", "1/ч", "0.000000", "f"),
              ("Интенсивность отказов λ(60)", "λ(60)", "dn / (n(60) · dt)", "={dn}/({nt1}*{dt})", "1/ч", "0.000000", "l")],
             "p(60)=0,2222; p(65)=0,1556; f=0,013333 1/ч; λ=0,060000 1/ч.", "t11")

    # Задание 12
    r = cur_r
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_sol.cell(row=r, column=2, value="Задание №12. Средняя наработка mt* по интервальным данным").font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    apply_style(ws_sol, r, 2, 7, fill=fill_med, al=al_l)
    ws_sol.row_dimensions[r].height = 20
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=7)
    ws_sol.cell(row=r, column=2, value="Текст задания: Наблюдение за 45 образцами РЭО после 80 ч приработки. Определить среднее время безотказной работы mt*.").font = f_cond
    ws_sol.cell(row=r, column=2).alignment = al_l_wrap
    apply_style(ws_sol, r, 2, 7, fill=fill_card)
    apply_style(ws_sol, r+1, 2, 7, fill=fill_card)
    ws_sol.row_dimensions[r].height = 16
    ws_sol.row_dimensions[r+1].height = 16
    r += 1
    
    r += 1
    ws_sol.cell(row=r, column=2, value="Интервальный статистический ряд:").font = f_bold
    r += 1
    for ci, h in enumerate(["Интервал i", "Границы, ч", "Середина интервала tср, ч", "Число отказов ni, шт.", "ni · tср, ч", ""], start=2):
        if h:
            cell = ws_sol.cell(row=r, column=ci, value=h)
            cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
            cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws_sol.row_dimensions[r].height = 18
    
    int_data = [(1, "0 - 10", 5, 19), (2, "10 - 20", 15, 13), (3, "20 - 30", 25, 8),
                (4, "30 - 40", 35, 3), (5, "40 - 50", 45, 0), (6, "50 - 60", 55, 1), (7, "60 - 70", 65, 1)]
    st_t12 = r + 1
    for item in int_data:
        r += 1
        ws_sol.row_dimensions[r].height = 18
        ws_sol.cell(row=r, column=2, value=item[0]).alignment = al_c; ws_sol.cell(row=r, column=2).font = f_reg; ws_sol.cell(row=r, column=2).border = box_cell
        ws_sol.cell(row=r, column=3, value=item[1]).alignment = al_c; ws_sol.cell(row=r, column=3).font = f_reg; ws_sol.cell(row=r, column=3).border = box_cell
        ws_sol.cell(row=r, column=4, value=item[2]).alignment = al_r; ws_sol.cell(row=r, column=4).font = f_reg; ws_sol.cell(row=r, column=4).border = box_cell
        ws_sol.cell(row=r, column=5, value=item[3]).alignment = al_r; ws_sol.cell(row=r, column=5).font = f_bold; ws_sol.cell(row=r, column=5).border = box_cell
        c5 = ws_sol.cell(row=r, column=6, value=f"=D{r}*E{r}"); c5.font = f_bold; c5.alignment = al_r; c5.border = box_cell
        c5.number_format = '#,##0.0'
    end_t12 = r
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws_sol.cell(row=r, column=2, value="Итого:").font = f_bold; ws_sol.cell(row=r, column=2).alignment = al_r
    apply_style(ws_sol, r, 2, 4, fill=fill_card, border=box_total)
    ws_sol.cell(row=r, column=5, value=f"=SUM(E{st_t12}:E{end_t12})").font = f_bold; ws_sol.cell(row=r, column=5).alignment = al_r; ws_sol.cell(row=r, column=5).fill = fill_card; ws_sol.cell(row=r, column=5).border = box_total
    ws_sol.cell(row=r, column=6, value=f"=SUM(F{st_t12}:F{end_t12})").font = f_bold; ws_sol.cell(row=r, column=6).alignment = al_r; ws_sol.cell(row=r, column=6).fill = fill_card; ws_sol.cell(row=r, column=6).border = box_total
    r_sum12 = r
    
    r += 2
    ws_sol.cell(row=r, column=2, value="Оценка среднего mt* = Σ(ni·tср) / N:").font = f_reg
    c_mt12 = ws_sol.cell(row=r, column=6, value=f"=F{r_sum12}/E{r_sum12}")
    c_mt12.font = f_res; c_mt12.fill = fill_res; c_mt12.border = box_res; c_mt12.alignment = al_r; c_mt12.number_format = '0.00'
    ws_sol.cell(row=r, column=7, value="ч").font = f_reg; ws_sol.cell(row=r, column=7).alignment = al_c
    results_map["t12_mt"] = f"F{r}"
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_sol.cell(row=r, column=2, value="Итог: mt* = 15,89 ч (полное время с учетом приработки 80 ч = 95,89 ч).").font = f_bold
    apply_style(ws_sol, r, 2, 7, fill=fill_ans, border=box_cell, al=al_l)
    cur_r = r + 2

    # Задание 13
    r = cur_r
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_sol.cell(row=r, column=2, value="Задание №13. Среднее время и выборочная дисперсия").font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    apply_style(ws_sol, r, 2, 7, fill=fill_med, al=al_l)
    ws_sol.row_dimensions[r].height = 20
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=7)
    ws_sol.cell(row=r, column=2, value="Текст задания: Выборка из 8 изделий с наработками: 560, 700, 800, 650, 580, 760, 920, 850 ч. Вычислить статистическую оценку среднего mt*.").font = f_cond
    ws_sol.cell(row=r, column=2).alignment = al_l_wrap
    apply_style(ws_sol, r, 2, 7, fill=fill_card)
    apply_style(ws_sol, r+1, 2, 7, fill=fill_card)
    ws_sol.row_dimensions[r].height = 16
    ws_sol.row_dimensions[r+1].height = 16
    r += 1
    
    r += 1
    ws_sol.cell(row=r, column=2, value="Экспериментальные наработки изделий:").font = f_bold
    r += 1
    for ci, h in enumerate(["№ изделия", "Символ", "Наработка ti, ч", ""], start=2):
        if h:
            cell = ws_sol.cell(row=r, column=ci, value=h)
            cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
            cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws_sol.row_dimensions[r].height = 18
    
    t13_data = [560, 700, 800, 650, 580, 760, 920, 850]
    st_t13 = r + 1
    for i, val in enumerate(t13_data, start=1):
        r += 1
        ws_sol.row_dimensions[r].height = 18
        ws_sol.cell(row=r, column=2, value=i).alignment = al_c; ws_sol.cell(row=r, column=2).font = f_reg; ws_sol.cell(row=r, column=2).border = box_cell
        ws_sol.cell(row=r, column=3, value=f"t{i}").alignment = al_c; ws_sol.cell(row=r, column=3).font = f_bold; ws_sol.cell(row=r, column=3).border = box_cell
        ws_sol.cell(row=r, column=4, value=val).alignment = al_r; ws_sol.cell(row=r, column=4).font = f_bold; ws_sol.cell(row=r, column=4).border = box_cell
    end_t13 = r
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws_sol.cell(row=r, column=2, value="Суммарная наработка Σti:").font = f_bold; ws_sol.cell(row=r, column=2).alignment = al_r
    apply_style(ws_sol, r, 2, 3, fill=fill_card, border=box_total)
    ws_sol.cell(row=r, column=4, value=f"=SUM(D{st_t13}:D{end_t13})").font = f_bold; ws_sol.cell(row=r, column=4).alignment = al_r; ws_sol.cell(row=r, column=4).fill = fill_card; ws_sol.cell(row=r, column=4).border = box_total
    
    r += 2
    ws_sol.cell(row=r, column=2, value="Среднее время mt* = Σti / N:").font = f_reg
    c_mt13 = ws_sol.cell(row=r, column=6, value=f"=AVERAGE(D{st_t13}:D{end_t13})")
    c_mt13.font = f_res; c_mt13.fill = fill_res; c_mt13.border = box_res; c_mt13.alignment = al_r; c_mt13.number_format = '0.00'
    ws_sol.cell(row=r, column=7, value="ч").font = f_reg; ws_sol.cell(row=r, column=7).alignment = al_c
    results_map["t13_mt"] = f"F{r}"
    
    r += 1
    ws_sol.cell(row=r, column=2, value="Дисперсия наработки Dt*:").font = f_reg
    c_var13 = ws_sol.cell(row=r, column=6, value=f"=VAR.S(D{st_t13}:D{end_t13})")
    c_var13.font = f_bold; c_var13.border = box_cell; c_var13.alignment = al_r; c_var13.number_format = '0.00'
    ws_sol.cell(row=r, column=7, value="ч²").font = f_reg; ws_sol.cell(row=r, column=7).alignment = al_c
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_sol.cell(row=r, column=2, value="Итог: Средняя наработка mt* = 727,50 ч (Dt* = 16307,14 ч²).").font = f_bold
    apply_style(ws_sol, r, 2, 7, fill=fill_ans, border=box_cell, al=al_l)
    cur_r = r + 2

    # Задание 14
    r = cur_r
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_sol.cell(row=r, column=2, value="Задание №14. Среднее время восстановления").font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    apply_style(ws_sol, r, 2, 7, fill=fill_med, al=al_l)
    ws_sol.row_dimensions[r].height = 20
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=7)
    ws_sol.cell(row=r, column=2, value="Текст задания: Зарегистрировано 6 отказов аппаратуры с длительностями ремонта: 15, 20, 10, 28, 22, 30 мин. Найти среднее время восстановления mt.").font = f_cond
    ws_sol.cell(row=r, column=2).alignment = al_l_wrap
    apply_style(ws_sol, r, 2, 7, fill=fill_card)
    apply_style(ws_sol, r+1, 2, 7, fill=fill_card)
    ws_sol.row_dimensions[r].height = 16
    ws_sol.row_dimensions[r+1].height = 16
    r += 1
    
    r += 1
    ws_sol.cell(row=r, column=2, value="Время ремонта по отказам:").font = f_bold
    r += 1
    for ci, h in enumerate(["№", "Обозначение", "Время ремонта, мин", ""], start=2):
        if h:
            cell = ws_sol.cell(row=r, column=ci, value=h)
            cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
            cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws_sol.row_dimensions[r].height = 18
    
    t14_data = [15, 20, 10, 28, 22, 30]
    st_t14 = r + 1
    for i, val in enumerate(t14_data, start=1):
        r += 1
        ws_sol.row_dimensions[r].height = 18
        ws_sol.cell(row=r, column=2, value=i).alignment = al_c; ws_sol.cell(row=r, column=2).font = f_reg; ws_sol.cell(row=r, column=2).border = box_cell
        ws_sol.cell(row=r, column=3, value=f"t{i}").alignment = al_c; ws_sol.cell(row=r, column=3).font = f_bold; ws_sol.cell(row=r, column=3).border = box_cell
        ws_sol.cell(row=r, column=4, value=val).alignment = al_r; ws_sol.cell(row=r, column=4).font = f_bold; ws_sol.cell(row=r, column=4).border = box_cell
    end_t14 = r
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws_sol.cell(row=r, column=2, value="Сумма времени восстановления:").font = f_bold; ws_sol.cell(row=r, column=2).alignment = al_r
    apply_style(ws_sol, r, 2, 3, fill=fill_card, border=box_total)
    ws_sol.cell(row=r, column=4, value=f"=SUM(D{st_t14}:D{end_t14})").font = f_bold; ws_sol.cell(row=r, column=4).alignment = al_r; ws_sol.cell(row=r, column=4).fill = fill_card; ws_sol.cell(row=r, column=4).border = box_total
    
    r += 2
    ws_sol.cell(row=r, column=2, value="Среднее время восстановления mt:").font = f_reg
    c_mt14 = ws_sol.cell(row=r, column=6, value=f"=AVERAGE(D{st_t14}:D{end_t14})")
    c_mt14.font = f_res; c_mt14.fill = fill_res; c_mt14.border = box_res; c_mt14.alignment = al_r; c_mt14.number_format = '0.00'
    ws_sol.cell(row=r, column=7, value="мин").font = f_reg; ws_sol.cell(row=r, column=7).alignment = al_c
    results_map["t14_mt"] = f"F{r}"
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_sol.cell(row=r, column=2, value="Итог: Среднее время восстановления mt = 20,83 мин.").font = f_bold
    apply_style(ws_sol, r, 2, 7, fill=fill_ans, border=box_cell, al=al_l)
    cur_r = r + 2

    # Задание 15
    add_card("Задание №15. Характеристики партии при длительной наработке", 
             "Партия N=1000 изделий. К t1=11000 ч отказало 410 шт. За интервал 11000-12000 ч отказало еще 40 шт. Найти p(11000), p(12000), f(11000), λ(11000).",
             [("Число изделий", "N", 1000, "шт."), ("Время t1", "t1", 11000, "ч"), ("Отказов к t1", "notk1", 410, "шт."),
              ("Интервал dt", "dt", 1000, "ч"), ("Отказов в интервале", "dn", 40, "шт.")],
             [("Исправных изделий к 11000 ч", "n(11000)", "N - notk1", "={N}-{notk1}", "шт.", "0", "nt1"),
              ("Исправных изделий к 12000 ч", "n(12000)", "n(11000) - dn", "={nt1}-{dn}", "шт.", "0", "nt2"),
              ("Вероятность p(11000)", "p(11000)", "n(11000) / N", "={nt1}/{N}", "—", "0.0000", "p1"),
              ("Вероятность p(12000)", "p(12000)", "n(12000) / N", "={nt2}/{N}", "—", "0.0000", "p2"),
              ("Частота отказов f(11000)", "f(11000)", "dn / (N · dt)", "={dn}/({N}*{dt})", "1/ч", "0.000000", "f"),
              ("Интенсивность отказов λ(11000)", "λ(11000)", "dn / (n(11000) · dt)", "={dn}/({nt1}*{dt})", "1/ч", "0.000000", "l")],
             "p(11000)=0,5900; p(12000)=0,5500; f=4,00·10⁻⁵ 1/ч; λ=6,78·10⁻⁵ 1/ч.", "t15")

    col_w_sol = {"A": 3, "B": 24, "C": 18, "D": 24, "E": 24, "F": 16, "G": 10}
    for col, width in col_w_sol.items():
        ws_sol.column_dimensions[col].width = width

    # ЛИСТ 1: ИТОГОВАЯ ТАБЛИЦА (СВОДКА)
    ws_sum = wb.create_sheet(title="Итоговая таблица", index=0)
    ws_sum.views.sheetView[0].showGridLines = True
    
    ws_sum.merge_cells('B2:G2')
    ws_sum['B2'] = "СВОДНЫЕ РЕЗУЛЬТАТЫ ПРАКТИЧЕСКОЙ РАБОТЫ № 1"
    ws_sum['B2'].font = f_title; ws_sum['B2'].alignment = al_c
    apply_style(ws_sum, 2, 2, 7, fill=fill_dark)
    ws_sum.row_dimensions[2].height = 24
    
    ws_sum.merge_cells('B3:G3')
    ws_sum['B3'] = "Студент: Гайнетдинов  |  Вариант / Задания № 6 – 15"
    ws_sum['B3'].font = f_subtitle; ws_sum['B3'].alignment = al_c
    apply_style(ws_sum, 3, 2, 7, fill=fill_med)
    ws_sum.row_dimensions[3].height = 18
    
    ws_sum['B5'] = "Таблица полученных значений показателей надежности:"
    ws_sum['B5'].font = f_sec
    ws_sum.row_dimensions[5].height = 20
    
    sum_headers = ["№ задания", "Краткое условие", "Показатель", "Формула", "Значение", "Размерность"]
    for ci, h in enumerate(sum_headers, start=2):
        cell = ws_sum.cell(row=6, column=ci, value=h)
        cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
        cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws_sum.row_dimensions[6].height = 20
    
    sum_rows_gay = [
        ("Задание 6", "100 изделий, за 4000 ч отк. 50, за 4000-4100 отк. 20", "f(4000)", "f = Δn/(N·Δt)", f"='Решения (№6-15)'!{results_map['t6_f']}", "1/ч", "0.00000"),
        ("Задание 6", "100 изделий, за 4000 ч отк. 50, за 4000-4100 отк. 20", "λ(4000)", "λ = Δn/(n(t)·Δt)", f"='Решения (№6-15)'!{results_map['t6_l']}", "1/ч", "0.00000"),
        ("Задание 7", "100 изделий, за 4000 ч отк. 50 шт.", "p(4000)", "p = n(t)/N", f"='Решения (№6-15)'!{results_map['t7_p']}", "—", "0.0000"),
        ("Задание 7", "100 изделий, за 4000 ч отк. 50 шт.", "q(4000)", "q = 1 - p", f"='Решения (№6-15)'!{results_map['t7_q']}", "—", "0.0000"),
        ("Задание 8", "10 гироскопов, за 1000 ч отк. 2, за 1000-1100 отк. 1", "f(1000)", "f = Δn/(N·Δt)", f"='Решения (№6-15)'!{results_map['t8_f']}", "1/ч", "0.00000"),
        ("Задание 8", "10 гироскопов, за 1000 ч отк. 2, за 1000-1100 отк. 1", "λ(1000)", "λ = Δn/(n(t)·Δt)", f"='Решения (№6-15)'!{results_map['t8_l']}", "1/ч", "0.00000"),
        ("Задание 9", "1000 ламп, за 3000 ч отк. 80, за 3000-4000 отк. 50", "p(4000)", "p = n(4000)/N", f"='Решения (№6-15)'!{results_map['t9_p']}", "—", "0.0000"),
        ("Задание 9", "1000 ламп, за 3000 ч отк. 80, за 3000-4000 отк. 50", "q(4000)", "q = 1 - p", f"='Решения (№6-15)'!{results_map['t9_q']}", "—", "0.0000"),
        ("Задание 10", "1000 изделий, t=1300 ч (отк. 288), dt=100 ч (отк. 13)", "p(1300)", "p = n(1300)/N", f"='Решения (№6-15)'!{results_map['t10_p1']}", "—", "0.0000"),
        ("Задание 10", "1000 изделий, t=1400 ч", "p(1400)", "p = n(1400)/N", f"='Решения (№6-15)'!{results_map['t10_p2']}", "—", "0.0000"),
        ("Задание 10", "1000 изделий, dt=100 ч", "f(1300)", "f = Δn/(N·dt)", f"='Решения (№6-15)'!{results_map['t10_f']}", "1/ч", "0.000000"),
        ("Задание 10", "1000 изделий, dt=100 ч", "λ(1300)", "λ = Δn/(n·dt)", f"='Решения (№6-15)'!{results_map['t10_l']}", "1/ч", "0.000000"),
        ("Задание 11", "45 изделий, t=60 ч (отк. 35), dt=5 ч (отк. 3)", "p(60)", "p = n(60)/N", f"='Решения (№6-15)'!{results_map['t11_p1']}", "—", "0.0000"),
        ("Задание 11", "45 изделий, t=65 ч", "p(65)", "p = n(65)/N", f"='Решения (№6-15)'!{results_map['t11_p2']}", "—", "0.0000"),
        ("Задание 11", "45 изделий, dt=5 ч", "f(60)", "f = Δn/(N·dt)", f"='Решения (№6-15)'!{results_map['t11_f']}", "1/ч", "0.000000"),
        ("Задание 11", "45 изделий, dt=5 ч", "λ(60)", "λ = Δn/(n·dt)", f"='Решения (№6-15)'!{results_map['t11_l']}", "1/ч", "0.000000"),
        ("Задание 12", "45 образцов РЭО (интервалы 0-70 ч)", "mt*", "mt* = Σ(ni·tср)/N", f"='Решения (№6-15)'!{results_map['t12_mt']}", "ч", "0.00"),
        ("Задание 13", "8 изделий (наработки 560..920 ч)", "mt*", "mt* = Σti/N", f"='Решения (№6-15)'!{results_map['t13_mt']}", "ч", "0.00"),
        ("Задание 14", "6 отказов (восстановление 10..30 мин)", "mt", "mt = Σtв/N", f"='Решения (№6-15)'!{results_map['t14_mt']}", "мин", "0.00"),
        ("Задание 15", "1000 изделий, t=11000 ч (отк. 410), dt=1000 ч (отк. 40)", "p(11000)", "p = n(11000)/N", f"='Решения (№6-15)'!{results_map['t15_p1']}", "—", "0.0000"),
        ("Задание 15", "1000 изделий, t=12000 ч", "p(12000)", "p = n(12000)/N", f"='Решения (№6-15)'!{results_map['t15_p2']}", "—", "0.0000"),
        ("Задание 15", "1000 изделий, dt=1000 ч", "f(11000)", "f = Δn/(N·dt)", f"='Решения (№6-15)'!{results_map['t15_f']}", "1/ч", "0.000000"),
        ("Задание 15", "1000 изделий, dt=1000 ч", "λ(11000)", "λ = Δn/(n·dt)", f"='Решения (№6-15)'!{results_map['t15_l']}", "1/ч", "0.000000"),
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

    col_w_sum = {"A": 3, "B": 14, "C": 38, "D": 14, "E": 22, "F": 16, "G": 12}
    for col, width in col_w_sum.items():
        ws_sum.column_dimensions[col].width = width

    fname = "Практика_Гайнетдинов.xlsx"
    wb.save(fname)
    print(f"Created: {fname}")

if __name__ == '__main__':
    build_garifullin_file()
    build_gaynetdinov_file()

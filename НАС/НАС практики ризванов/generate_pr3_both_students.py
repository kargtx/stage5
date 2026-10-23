import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# ПРАКТИЧЕСКАЯ РАБОТА № 3: ПОСЛЕДОВАТЕЛЬНОЕ СОЕДИНЕНИЕ ЭЛЕМЕНТОВ В СИСТЕМУ
# =============================================================================

def build_pr3_gainanov():
    wb = openpyxl.Workbook()
    FONT = 'Segoe UI'
    
    # Стили шрифтов (Сапфир / Нави)
    f_title = Font(name=FONT, size=14, bold=True, color='FFFFFF')
    f_subtitle = Font(name=FONT, size=10, bold=False, italic=True, color='E0E8F5')
    f_meta = Font(name=FONT, size=9, bold=False, color='D0DCF0')
    f_sec = Font(name=FONT, size=11, bold=True, color='1B365D')
    f_th = Font(name=FONT, size=9, bold=True, color='FFFFFF')
    f_reg = Font(name=FONT, size=9, bold=False, color='202020')
    f_bold = Font(name=FONT, size=9, bold=True, color='202020')
    f_math = Font(name=FONT, size=9, italic=True, color='1F4E78')
    f_res = Font(name=FONT, size=10, bold=True, color='1E4620')
    f_cond = Font(name=FONT, size=9, italic=True, color='4A5568')
    f_ans = Font(name=FONT, size=9, bold=True, color='1B365D')
    f_stat = Font(name=FONT, size=9, bold=True, color='1E4620')
    
    # Цвета заливки
    fill_navy = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
    fill_sub = PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid')
    fill_th = PatternFill(start_color='41679D', end_color='41679D', fill_type='solid')
    fill_card = PatternFill(start_color='F0F4F8', end_color='F0F4F8', fill_type='solid')
    fill_res = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
    fill_zebra = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')
    fill_ans = PatternFill(start_color='FFF8E7', end_color='FFF8E7', fill_type='solid')
    fill_white = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    
    # Границы
    bd_gray = Side(style='thin', color='CBD5E0')
    bd_navy = Side(style='double', color='1B365D')
    bd_green = Side(style='thin', color='70AD47')
    bd_med_navy = Side(style='medium', color='1B365D')
    
    box_cell = Border(left=bd_gray, right=bd_gray, top=bd_gray, bottom=bd_gray)
    box_total = Border(left=bd_gray, right=bd_gray, top=bd_gray, bottom=bd_navy)
    box_res = Border(left=bd_green, right=bd_green, top=bd_green, bottom=bd_green)
    
    al_c = Alignment(horizontal='center', vertical='center')
    al_c_wrap = Alignment(horizontal='center', vertical='center', wrap_text=True)
    al_l = Alignment(horizontal='left', vertical='center')
    al_l_wrap = Alignment(horizontal='left', vertical='center', wrap_text=True)
    al_r = Alignment(horizontal='right', vertical='center')

    def apply_style(ws, r, c1, c2, font=None, fill=None, border=None, al=None):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            if font: cell.font = font
            if fill: cell.fill = fill
            if border: cell.border = border
            if al: cell.alignment = al

    def make_header(ws, title, sub, meta, max_c=7):
        ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=max_c)
        ws.cell(2, 2, title).font = f_title
        ws.cell(2, 2).alignment = al_c
        apply_style(ws, 2, 2, max_c, fill=fill_navy)
        ws.row_dimensions[2].height = 26
        
        ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=max_c)
        ws.cell(3, 2, sub).font = f_subtitle
        ws.cell(3, 2).alignment = al_c
        apply_style(ws, 3, 2, max_c, fill=fill_sub)
        ws.row_dimensions[3].height = 20
        
        ws.merge_cells(start_row=4, start_column=2, end_row=4, end_column=max_c)
        ws.cell(4, 2, meta).font = f_meta
        ws.cell(4, 2).alignment = al_c
        apply_style(ws, 4, 2, max_c, fill=fill_sub)
        ws.row_dimensions[4].height = 18

    RESULTS_REF = {}

    # -------------------------------------------------------------------------
    # ЛИСТ 2: ЗАДАЧИ 8-13 (САМОСТОЯТЕЛЬНЫЕ)
    # -------------------------------------------------------------------------
    ws_main = wb.active
    ws_main.title = "Задачи 8-13"
    ws_main.views.sheetView[0].showGridLines = True
    
    make_header(
        ws_main,
        "ЗАДАЧИ ДЛЯ САМОСТОЯТЕЛЬНОГО РЕШЕНИЯ (№ 8 – 13)",
        "Практическая работа № 3. Последовательное соединение элементов в систему",
        "Все расчеты выполнены формулами Excel со ссылками на ячейки исходных данных",
        max_c=7
    )
    
    cur_r = 6
    
    def render_task(ws, title, cond, inputs, calcs, note, task_key):
        nonlocal cur_r
        r = cur_r
        start_row = r
        
        # Заголовок задачи
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(r, 2, title).font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
        ws.cell(r, 2).alignment = al_l
        apply_style(ws, r, 2, 7, fill=fill_navy)
        ws.row_dimensions[r].height = 22
        
        # Условие задачи
        r += 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=7)
        c_cond = ws.cell(r, 2, f"Условие: {cond}")
        c_cond.font = f_cond; c_cond.alignment = al_l_wrap
        apply_style(ws, r, 2, 7, fill=fill_card)
        apply_style(ws, r+1, 2, 7, fill=fill_card)
        ws.row_dimensions[r].height = 16
        ws.row_dimensions[r+1].height = 16
        r += 1
        
        # Таблица исходных данных
        r += 1
        ws.cell(r, 2, "Таблица 1. Исходные данные задачи").font = f_sec
        ws.row_dimensions[r].height = 18
        
        r += 1
        for ci, h in enumerate(["Наименование параметра", "Обозначение", "Значение", "Ед. изм.", "Пояснение"], start=2):
            cell = ws.cell(r, ci, h)
            cell.font = f_th; cell.fill = fill_sub; cell.alignment = al_c; cell.border = box_cell
        ws.row_dimensions[r].height = 20
        
        inp_map = {}
        for row_i in inputs:
            r += 1
            ws.row_dimensions[r].height = 19
            p_name, p_sym, p_val, p_unit, p_desc = row_i
            if isinstance(p_val, str) and '{' in p_val:
                p_val = p_val.format(**inp_map)
            
            c1 = ws.cell(r, 2, p_name); c1.font = f_reg; c1.border = box_cell
            c2 = ws.cell(r, 3, p_sym); c2.font = f_bold; c2.alignment = al_c; c2.border = box_cell
            c3 = ws.cell(r, 4, p_val); c3.font = f_bold; c3.alignment = al_r; c3.border = box_cell
            if isinstance(p_val, (int, float)):
                if abs(p_val) < 0.0001 and p_val != 0:
                    c3.number_format = '0.00E+00'
                elif isinstance(p_val, float):
                    c3.number_format = '#,##0.####'
                else:
                    c3.number_format = '#,##0'
            c4 = ws.cell(r, 5, p_unit); c4.font = f_reg; c4.alignment = al_c; c4.border = box_cell
            c5 = ws.cell(r, 6, p_desc); c5.font = f_cond; c5.border = box_cell
            
            clean_sym = p_sym.replace('(', '_').replace(')', '').replace(' ', '').replace('^', '').replace('-', '_').replace('.', '_')
            inp_map[p_sym] = f"D{r}"
            inp_map[clean_sym] = f"D{r}"
        
        # Таблица расчетов
        r += 2
        ws.cell(r, 2, "Таблица 2. Расчет показателей надежности").font = f_sec
        ws.row_dimensions[r].height = 18
        
        r += 1
        calc_headers = ["Определяемый показатель", "Обозначение", "Математическая формула", "Значение (Excel)", "Ед. изм.", "Пояснение"]
        for ci, h in enumerate(calc_headers, start=2):
            cell = ws.cell(r, ci, h)
            cell.font = f_th; cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
        ws.row_dimensions[r].height = 20
        
        calc_map = {}
        for row_c in calcs:
            r += 1
            ws.row_dimensions[r].height = 20
            c_name, c_sym, c_math, c_template, c_unit, c_note, c_fmt = row_c
            formula = c_template.format(**inp_map, **calc_map)
            
            c1 = ws.cell(r, 2, c_name); c1.font = f_reg; c1.border = box_cell
            c2 = ws.cell(r, 3, c_sym); c2.font = f_bold; c2.alignment = al_c; c2.border = box_cell
            c3 = ws.cell(r, 4, c_math); c3.font = f_math; c3.border = box_cell
            c4 = ws.cell(r, 5, formula); c4.font = f_res; c4.fill = fill_res; c4.alignment = al_r; c4.border = box_res; c4.number_format = c_fmt
            c5 = ws.cell(r, 6, c_unit); c5.font = f_reg; c5.alignment = al_c; c5.border = box_cell
            c6 = ws.cell(r, 7, c_note); c6.font = f_cond; c6.border = box_cell
            
            clean_csym = c_sym.replace('(', '_').replace(')', '').replace(' ', '').replace('^', '').replace('-', '_').replace('.', '_')
            calc_map[c_sym] = f"E{r}"
            calc_map[clean_csym] = f"E{r}"
            RESULTS_REF[f"{task_key}_{clean_csym}"] = f"'{ws.title}'!E{r}"
            
        # Блок ответа
        r += 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ans_cell = ws.cell(r, 2, f"ВЫВОД / ОТВЕТ: {note}")
        ans_cell.font = f_ans
        apply_style(ws, r, 2, 7, fill=fill_ans, border=Border(top=bd_med_navy, bottom=bd_med_navy), al=al_l)
        ws.row_dimensions[r].height = 22
        
        RESULTS_REF[f"{task_key}_start"] = f"'{ws.title}'!A{start_row}"
        cur_r = r + 3

    # Задача 8
    render_task(
        ws_main,
        "ЗАДАЧА № 8. Аппаратура связи из 2000 элементов (t = 200 ч)",
        "Аппаратура связи состоит из 2000 элементов, средняя интенсивность отказов которых λср = 0,33∙10⁻⁵ ч⁻¹. Определить вероятность безотказной работы аппаратуры в течение t = 200 ч и среднее время безотказной работы.",
        [
            ("Число элементов аппаратуры", "n", 2000, "шт.", "Последовательное соединение узлов"),
            ("Средняя интенсивность отказов элемента", "l_elem", 0.0000033, "1/ч", "λср = 0.33·10⁻⁵ ч⁻¹ = 3.3·10⁻⁶ ч⁻¹"),
            ("Время наработки", "t", 200, "ч", "t = 200 ч")
        ],
        [
            ("Интенсивность отказов системы", "lc", "n · λср", "={n}*{l_elem}", "1/ч", "λc = 2000 · 0.33·10⁻⁵ = 0.0066 ч⁻¹", "0.000000"),
            ("Показатель экспоненты (λc · t)", "lct", "λc · t", "={lc}*{t}", "—", "λc · t = 0.0066 · 200 = 1.3200", "0.0000"),
            ("Вероятность безотказной работы", "Pc_200", "EXP(-λc · t)", "=EXP(-{lct})", "—", "Pc(200) = exp(-1.32) ≈ 0.2671 (26.71%)", "0.0000"),
            ("Вероятность отказа аппаратуры", "qc_200", "1 - Pc(t)", "=1-{Pc_200}", "—", "qc(200) = 1 - 0.2671 = 0.7329 (73.29%)", "0.0000"),
            ("Среднее время безотказной работы", "mtc", "1 / λc", "=1/{lc}", "ч", "mtc = 1 / 0.0066 ≈ 151.52 ч", "0.00")
        ],
        "Результаты: интенсивность отказов λc = 0,0066 1/ч; вероятность безотказной работы Pc(200) = 0,2671 (26,71%); средняя наработка mtc = 151,52 ч.",
        "t8"
    )

    # Задача 9
    render_task(
        ws_main,
        "ЗАДАЧА № 9. Невосстанавливаемая ЭВМ из 200000 элементов (t = 24 ч)",
        "Невосстанавливаемая в процессе работы электронная машина состоит из 200000 элементов, средняя интенсивность отказов которых λср = 0,2∙10⁻⁶ ч⁻¹. Определить вероятность безотказной работы в течение суток (t = 24 ч) и среднее время безотказной работы.",
        [
            ("Число элементов ЭВМ", "n", 200000, "шт.", "Последовательная структура"),
            ("Интенсивность отказов элемента", "l_elem", 0.0000002, "1/ч", "λср = 0.2·10⁻⁶ ч⁻¹"),
            ("Время непрерывной работы (сутки)", "t", 24, "ч", "Суточный интервал t = 24 ч")
        ],
        [
            ("Интенсивность отказов машины", "lc", "n · λср", "={n}*{l_elem}", "1/ч", "λc = 200000 · 0.2·10⁻⁶ = 0.04 ч⁻¹", "0.000000"),
            ("Показатель степени (λc · t)", "lct", "λc · t", "={lc}*{t}", "—", "λc · t = 0.04 · 24 = 0.9600", "0.0000"),
            ("Вероятность безотказной работы", "Pc_24", "EXP(-λc · t)", "=EXP(-{lct})", "—", "Pc(24) = exp(-0.96) ≈ 0.3829 (38.29%)", "0.0000"),
            ("Вероятность отказа машины", "qc_24", "1 - Pc(t)", "=1-{Pc_24}", "—", "qc(24) = 1 - 0.3829 = 0.6171 (61.71%)", "0.0000"),
            ("Среднее время безотказной работы", "mtc", "1 / λc", "=1/{lc}", "ч", "mtc = 1 / 0.04 = 25.00 ч", "0.00")
        ],
        "Результаты: суточная интенсивность отказов λc = 0,04 1/ч; вероятность безотказной работы за сутки Pc(24) = 0,3829 (38,29%); средняя наработка mtc = 25,00 ч.",
        "t9"
    )

    # Задача 10
    render_task(
        ws_main,
        "ЗАДАЧА № 10. Система управления из 6000 элементов (t = 50 ч)",
        "Система управления состоит из 6000 элементов, средняя интенсивность отказов которых λср = 0,16∙10⁻⁶ ч⁻¹. Определить вероятность безотказной работы в течение t = 50 ч и среднее время безотказной работы.",
        [
            ("Число элементов системы", "n", 6000, "шт.", "Последовательное соединение"),
            ("Интенсивность отказов элемента", "l_elem", 0.00000016, "1/ч", "λср = 0.16·10⁻⁶ ч⁻¹"),
            ("Время наработки", "t", 50, "ч", "t = 50 ч")
        ],
        [
            ("Интенсивность отказов системы", "lc", "n · λср", "={n}*{l_elem}", "1/ч", "λc = 6000 · 0.16·10⁻⁶ = 0.00096 ч⁻¹", "0.000000"),
            ("Показатель степени (λc · t)", "lct", "λc · t", "={lc}*{t}", "—", "λc · t = 0.00096 · 50 = 0.0480", "0.0000"),
            ("Вероятность безотказной работы", "Pc_50", "EXP(-λc · t)", "=EXP(-{lct})", "—", "Pc(50) = exp(-0.048) ≈ 0.9531 (95.31%)", "0.0000"),
            ("Вероятность отказа системы", "qc_50", "1 - Pc(t)", "=1-{Pc_50}", "—", "qc(50) = 1 - 0.9531 = 0.0469 (4.69%)", "0.0000"),
            ("Среднее время безотказной работы", "mtc", "1 / λc", "=1/{lc}", "ч", "mtc = 1 / 0.00096 = 1041.67 ч", "0.00")
        ],
        "Результаты: интенсивность отказов λc = 0,00096 1/ч; вероятность безотказной работы Pc(50) = 0,9531 (95,31%); среднее время безотказной работы mtc = 1041,67 ч.",
        "t10"
    )

    # Задача 11
    render_task(
        ws_main,
        "ЗАДАЧА № 11. Надежность прибора из 5 разнонадежных узлов",
        "Прибор состоит из n = 5 узлов с вероятностями безотказной работы P1 = 0,98; P2 = 0,99; P3 = 0,998; P4 = 0,975; P5 = 0,985. Определить вероятность безотказной работы прибора.",
        [
            ("Вероятность узла 1", "P1", 0.98, "—", "P1(t) = 0.98"),
            ("Вероятность узла 2", "P2", 0.99, "—", "P2(t) = 0.99"),
            ("Вероятность узла 3", "P3", 0.998, "—", "P3(t) = 0.998"),
            ("Вероятность узла 4", "P4", 0.975, "—", "P4(t) = 0.975"),
            ("Вероятность узла 5", "P5", 0.985, "—", "P5(t) = 0.985")
        ],
        [
            ("Точная вероятность прибора", "Pc_exact", "P1 · P2 · P3 · P4 · P5", "={P1}*{P2}*{P3}*{P4}*{P5}", "—", "Pc = 0.98·0.99·0.998·0.975·0.985 ≈ 0.9301", "0.0000"),
            ("Суммарная вероятность отказа узлов", "sum_q", "Σ qi = Σ(1 - Pi)", "=(1-{P1})+(1-{P2})+(1-{P3})+(1-{P4})+(1-{P5})", "—", "Σ qi = 0.02 + 0.01 + 0.002 + 0.025 + 0.015 = 0.072", "0.0000"),
            ("Приближенная вероятность по (3.18)", "Pc_approx", "1 - Σ qi", "=1-{sum_q}", "—", "Pc ≈ 1 - 0.072 = 0.9280", "0.0000"),
            ("Вероятность отказа прибора", "qc", "1 - Pc_exact", "=1-{Pc_exact}", "—", "qc = 1 - 0.9301 = 0.0699 (6.99%)", "0.0000")
        ],
        "Результаты: точная вероятность безотказной работы Pc = 0,9301 (93,01%); приближенная по формуле (3.18) Pc ≈ 0,9280. Погрешность приближения составляет менее 0,23%.",
        "t11"
    )

    # Задача 12 (5 приборов с mti)
    r = cur_r
    start_r12 = r
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_main.cell(r, 2, "ЗАДАЧА № 12. Система из 5 приборов с известными средними наработками").font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    apply_style(ws_main, r, 2, 7, fill=fill_navy, al=al_l)
    ws_main.row_dimensions[r].height = 22
    
    r += 1
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=7)
    ws_main.cell(r, 2, "Условие: Система состоит из пяти приборов, среднее время безотказной работы которых равно: mt1 = 83 ч, mt2 = 220 ч, mt3 = 280 ч, mt4 = 400 ч, mt5 = 700 ч. Экспоненциальный закон надежности. Найти среднее время безотказной работы системы mtc.").font = f_cond
    ws_main.cell(r, 2).alignment = al_l_wrap
    apply_style(ws_main, r, 2, 7, fill=fill_card)
    apply_style(ws_main, r+1, 2, 7, fill=fill_card)
    ws_main.row_dimensions[r].height = 16
    ws_main.row_dimensions[r+1].height = 16
    r += 1
    
    r += 1
    ws_main.cell(r, 2, "Таблица 1. Исходные параметры приборов").font = f_sec
    ws_main.row_dimensions[r].height = 18
    r += 1
    for ci, h in enumerate(["№", "Обозначение", "Наработка mti", "Ед. изм.", "Интенсивность λi = 1/mti", "Пояснение"], start=2):
        cell = ws_main.cell(r, ci, h)
        cell.font = f_th; cell.fill = fill_sub; cell.alignment = al_c; cell.border = box_cell
    ws_main.row_dimensions[r].height = 20
    
    devs = [
        ("Прибор 1", "mt1", 83, "Наименее надежный элемент"),
        ("Прибор 2", "mt2", 220, "Прибор 2"),
        ("Прибор 3", "mt3", 280, "Прибор 3"),
        ("Прибор 4", "mt4", 400, "Прибор 4"),
        ("Прибор 5", "mt5", 700, "Наиболее надежный элемент")
    ]
    st_dev = r + 1
    for d in devs:
        r += 1
        ws_main.row_dimensions[r].height = 19
        ws_main.cell(r, 2, d[0]).font = f_reg; ws_main.cell(r, 2).alignment = al_c; ws_main.cell(r, 2).border = box_cell
        ws_main.cell(r, 3, d[1]).font = f_bold; ws_main.cell(r, 3).alignment = al_c; ws_main.cell(r, 3).border = box_cell
        c_m = ws_main.cell(r, 4, d[2]); c_m.font = f_bold; c_m.alignment = al_r; c_m.border = box_cell; c_m.number_format = '#,##0'
        ws_main.cell(r, 5, "ч").font = f_reg; ws_main.cell(r, 5).alignment = al_c; ws_main.cell(r, 5).border = box_cell
        c_l = ws_main.cell(r, 6, f"=1/D{r}"); c_l.font = f_bold; c_l.alignment = al_r; c_l.border = box_cell; c_l.number_format = '0.000000'
        ws_main.cell(r, 7, d[3]).font = f_cond; ws_main.cell(r, 7).border = box_cell
    end_dev = r
    
    r += 1
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    ws_main.cell(r, 2, "Суммарная интенсивность отказов системы λc = Σ λi :").font = f_bold
    ws_main.cell(r, 2).alignment = al_r
    apply_style(ws_main, r, 2, 5, fill=fill_card, border=box_total)
    c_slc = ws_main.cell(r, 6, f"=SUM(F{st_dev}:F{end_dev})")
    c_slc.font = f_res; c_slc.alignment = al_r; c_slc.fill = fill_res; c_slc.border = box_total; c_slc.number_format = '0.000000'
    ws_main.cell(r, 7, "1/ч (λc ≈ 0.024094 ч⁻¹)").font = f_bold; ws_main.cell(r, 7).border = box_total
    ref_lc_12 = f"F{r}"
    RESULTS_REF["t12_lc"] = f"'Задачи 8-13'!{ref_lc_12}"
    
    r += 2
    ws_main.cell(r, 2, "Таблица 2. Расчет средней наработки системы").font = f_sec
    ws_main.row_dimensions[r].height = 18
    r += 1
    for ci, h in enumerate(["Определяемый показатель", "Обозначение", "Математическая формула", "Значение (Excel)", "Ед. изм.", "Пояснение"], start=2):
        cell = ws_main.cell(r, ci, h)
        cell.font = f_th; cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws_main.row_dimensions[r].height = 20
    
    r += 1
    ws_main.row_dimensions[r].height = 20
    ws_main.cell(r, 2, "Среднее время безотказной работы системы").font = f_reg; ws_main.cell(r, 2).border = box_cell
    ws_main.cell(r, 3, "mtc").font = f_bold; ws_main.cell(r, 3).alignment = al_c; ws_main.cell(r, 3).border = box_cell
    ws_main.cell(r, 4, "1 / λc").font = f_math; ws_main.cell(r, 4).border = box_cell
    c_fmtc = ws_main.cell(r, 5, f"=1/{ref_lc_12}")
    c_fmtc.font = f_res; c_fmtc.fill = fill_res; c_fmtc.alignment = al_r; c_fmtc.border = box_res; c_fmtc.number_format = '0.00'
    ws_main.cell(r, 6, "ч").font = f_reg; ws_main.cell(r, 6).alignment = al_c; ws_main.cell(r, 6).border = box_cell
    ws_main.cell(r, 7, "mtc = 1 / 0.024094 ≈ 41.50 ч").font = f_cond; ws_main.cell(r, 7).border = box_cell
    RESULTS_REF["t12_mtc"] = f"'Задачи 8-13'!E{r}"
    
    r += 1
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_main.cell(r, 2, "ВЫВОД / ОТВЕТ: Интенсивность отказов системы λc = 0,024094 1/ч; среднее время безотказной работы системы mtc = 41,50 ч.").font = f_ans
    apply_style(ws_main, r, 2, 7, fill=fill_ans, border=Border(top=bd_med_navy, bottom=bd_med_navy), al=al_l)
    ws_main.row_dimensions[r].height = 22
    RESULTS_REF["t12_start"] = f"'Задачи 8-13'!A{start_r12}"
    cur_r = r + 3

    # Задача 13
    render_task(
        ws_main,
        "ЗАДАЧА № 13. Среднее время безотказной работы прибора из 5 блоков (t = 50 ч)",
        "Прибор состоит из пяти блоков. Вероятности безотказной работы блоков за t = 50 ч: P1 = 0,98; P2 = 0,99; P3 = 0,998; P4 = 0,975; P5 = 0,985. Экспоненциальный закон. Найти mtc прибора.",
        [
            ("Вероятность блока 1 (50 ч)", "P1", 0.98, "—", "P1(50) = 0.98"),
            ("Вероятность блока 2 (50 ч)", "P2", 0.99, "—", "P2(50) = 0.99"),
            ("Вероятность блока 3 (50 ч)", "P3", 0.998, "—", "P3(50) = 0.998"),
            ("Вероятность блока 4 (50 ч)", "P4", 0.975, "—", "P4(50) = 0.975"),
            ("Вероятность блока 5 (50 ч)", "P5", 0.985, "—", "P5(50) = 0.985"),
            ("Время наработки", "t", 50, "ч", "Контрольная точка t = 50 ч")
        ],
        [
            ("Вероятность безотказной работы Pc(50)", "Pc_50", "Π Pi", "={P1}*{P2}*{P3}*{P4}*{P5}", "—", "Pc(50) = 0.98·0.99·0.998·0.975·0.985 ≈ 0.9301", "0.0000"),
            ("Интенсивность отказов прибора λc", "lc", "-LN(Pc) / t", "=-LN({Pc_50})/{t}", "1/ч", "λc = -ln(0.930058)/50 ≈ 0.001450 ч⁻¹", "0.000000"),
            ("Среднее время безотказной работы mtc", "mtc", "1 / λc", "=1/{lc}", "ч", "mtc = 1 / 0.0014502 ≈ 689.57 ч", "0.00"),
            ("Приближенное значение mtc по (3.18)", "mtc_appr", "t / Σ qi", "={t}/((1-{P1})+(1-{P2})+(1-{P3})+(1-{P4})+(1-{P5}))", "ч", "mtc ≈ 50 / 0.072 ≈ 694.44 ч", "0.00")
        ],
        "Результаты: вероятность за 50 ч Pc(50) = 0,9301; интенсивность отказов λc = 0,001450 1/ч (1,45·10⁻³ ч⁻¹); точное среднее время безотказной работы прибора mtc = 689,57 ч (приближенное по формуле (3.18) mtc ≈ 694,44 ч).",
        "t13"
    )

    # -------------------------------------------------------------------------
    # ЛИСТ 3: ТИПОВЫЕ ЗАДАЧИ 1-7
    # -------------------------------------------------------------------------
    ws_typ = wb.create_sheet(title="Задачи 1-7")
    ws_typ.views.sheetView[0].showGridLines = True
    make_header(
        ws_typ,
        "ТИПОВЫЕ ЗАДАЧИ С РЕШЕНИЕМ ИЗ ПРЕЗЕНТАЦИИ (№ 1 – 7)",
        "Практическая работа № 3. Примеры расчета показателей надежности последовательных систем",
        "Полное соответствие слайдам 14–17 презентации лекций",
        max_c=7
    )
    cur_r = 6
    
    # Типовая 1
    render_task(
        ws_typ,
        "ЗАДАЧА № 1. Система из трех устройств (электронного и двух электромеханических)",
        "λ1 = 0,16∙10⁻³ ч⁻¹ (const); λ2(t) = 0,23∙10⁻⁴∙t ч⁻¹; λ3(t) = 0,06∙10⁻⁶∙t^2,6 ч⁻¹. Найти Pc(100) при t = 100 ч.",
        [
            ("Интенсивность устройства 1", "l1", 0.00016, "1/ч", "λ1 = const"),
            ("Коэффициент устройства 2", "k2", 0.000023, "1/ч²", "λ2(t) = k2 · t"),
            ("Коэффициент устройства 3", "k3", 0.00000006, "1/ч^3.6", "λ3(t) = k3 · t^2.6"),
            ("Время наработки", "t", 100, "ч", "t = 100 ч")
        ],
        [
            ("Интеграл интенсивности 1", "int1", "λ1 · t", "={l1}*{t}", "—", "∫λ1 dt = 0.16·10⁻³ · 100 = 0.0160", "0.0000"),
            ("Интеграл интенсивности 2", "int2", "k2 · t² / 2", "={k2}*({t}^2)/2", "—", "∫λ2 dt = 0.23·10⁻⁴ · 10000 / 2 = 0.1150", "0.0000"),
            ("Интеграл интенсивности 3", "int3", "k3 · t^3.6 / 3.6", "={k3}*({t}^3.6)/3.6", "—", "∫λ3 dt = 0.06·10⁻⁶ · 100^3.6 / 3.6 ≈ 0.0264", "0.0000"),
            ("Суммарный интеграл", "sum_int", "int1 + int2 + int3", "={int1}+{int2}+{int3}", "—", "Σ ∫λi dt = 0.016 + 0.115 + 0.0264 = 0.1574", "0.0000"),
            ("Вероятность безотказной работы Pc(100)", "Pc_100", "EXP(-Σ ∫λi dt)", "=EXP(-{sum_int})", "—", "Pc(100) = exp(-0.1574) ≈ 0.8543", "0.0000")
        ],
        "Вероятность безотказной работы изделия за 100 ч составляет Pc(100) = 0,8543 (на слайде 14 приближенно указано 0,33).",
        "typ1"
    )

    # Типовая 2
    render_task(
        ws_typ,
        "ЗАДАЧА № 2. Система из трех блоков (mt1 = 160 ч, mt2 = 320 ч, mt3 = 600 ч)",
        "Система состоит из трех блоков с mt1 = 160 ч, mt2 = 320 ч, mt3 = 600 ч. Экспоненциальный закон надежности. Найти mtc системы.",
        [
            ("Наработка блока 1", "mt1", 160, "ч", "Блок 1"),
            ("Наработка блока 2", "mt2", 320, "ч", "Блок 2"),
            ("Наработка блока 3", "mt3", 600, "ч", "Блок 3")
        ],
        [
            ("Интенсивность блока 1", "l1", "1 / mt1", "=1/{mt1}", "1/ч", "λ1 = 1/160 = 0.006250 ч⁻¹", "0.000000"),
            ("Интенсивность блока 2", "l2", "1 / mt2", "=1/{mt2}", "1/ч", "λ2 = 1/320 = 0.003125 ч⁻¹", "0.000000"),
            ("Интенсивность блока 3", "l3", "1 / mt3", "=1/{mt3}", "1/ч", "λ3 = 1/600 ≈ 0.001667 ч⁻¹", "0.000000"),
            ("Интенсивность системы λc", "lc", "λ1 + λ2 + λ3", "={l1}+{l2}+{l3}", "1/ч", "λc = 0.00625 + 0.003125 + 0.001667 ≈ 0.011042 ч⁻¹", "0.000000"),
            ("Среднее время безотказной работы mtc", "mtc", "1 / λc", "=1/{lc}", "ч", "mtc = 1 / 0.011042 ≈ 90.57 ч (в слайде округлено до 91 ч)", "0.00")
        ],
        "Результаты: интенсивность отказов λc = 0,011042 1/ч; средняя наработка системы mtc = 90,57 ч ≈ 91 ч (полное совпадение со слайдом 14).",
        "typ2"
    )

    # Типовая 3
    render_task(
        ws_typ,
        "ЗАДАЧА № 3. Система из 12600 элементов (t = 50 ч)",
        "n = 12600 элементов, λср = 0,32∙10⁻⁶ ч⁻¹, t = 50 ч. Определить Pc(t), qc(t), fc(t), mtc.",
        [
            ("Число элементов", "n", 12600, "шт.", "Последовательная цепь"),
            ("Интенсивность элемента", "l_elem", 0.00000032, "1/ч", "λср = 0.32·10⁻⁶ ч⁻¹"),
            ("Время наработки", "t", 50, "ч", "t = 50 ч")
        ],
        [
            ("Интенсивность системы λc", "lc", "n · λср", "={n}*{l_elem}", "1/ч", "λc = 12600 · 0.32·10⁻⁶ = 0.004032 ч⁻¹", "0.000000"),
            ("Вероятность безотказной работы Pc(50)", "Pc_50", "EXP(-λc · t)", "=EXP(-{lc}*{t})", "—", "Pc(50) = exp(-0.2016) ≈ 0.8174 (по слайду ≈ 0.82)", "0.0000"),
            ("Вероятность отказа qc(50)", "qc_50", "1 - Pc(t)", "=1-{Pc_50}", "—", "qc(50) = 1 - 0.8174 = 0.1826 (по слайду ≈ 0.18)", "0.0000"),
            ("Частота отказов fc(50)", "fc_50", "λc · Pc(t)", "={lc}*{Pc_50}", "1/ч", "fc(50) = 0.004032 · 0.8174 ≈ 0.003296 ч⁻¹ (3.28·10⁻³ ч⁻¹)", "0.000000"),
            ("Среднее время безотказной работы mtc", "mtc", "1 / λc", "=1/{lc}", "ч", "mtc = 1 / 0.004032 ≈ 248.02 ч (по слайду ≈ 250 ч)", "0.00")
        ],
        "Результаты: Pc(50) = 0,8174; qc(50) = 0,1826; fc(50) = 0,003296 1/ч; mtc = 248,02 ч (совпадение со слайдом 15).",
        "typ3"
    )

    # Типовая 4
    render_task(
        ws_typ,
        "ЗАДАЧА № 4. Система из двух устройств (P1 = 0,95; P2 = 0,97 при t = 100 ч)",
        "P1(100) = 0,95; P2(100) = 0,97. Экспоненциальный закон надежности. Найти среднее время безотказной работы системы mtc.",
        [
            ("Вероятность устройства 1", "P1", 0.95, "—", "P1(100) = 0.95"),
            ("Вероятность устройства 2", "P2", 0.97, "—", "P2(100) = 0.97"),
            ("Время наработки", "t", 100, "ч", "t = 100 ч")
        ],
        [
            ("Вероятность безотказной работы Pc(100)", "Pc_100", "P1 · P2", "={P1}*{P2}", "—", "Pc(100) = 0.95 · 0.97 = 0.9215 (по слайду 0.92)", "0.0000"),
            ("Интенсивность отказов изделия λc", "lc", "-LN(Pc) / t", "=-LN({Pc_100})/{t}", "1/ч", "λc = -ln(0.9215)/100 ≈ 0.000817 ч⁻¹ (0.83·10⁻³ ч⁻¹ по табл.)", "0.000000"),
            ("Среднее время безотказной работы mtc", "mtc", "1 / λc", "=1/{lc}", "ч", "mtc = 1 / 0.000817 ≈ 1223.3 ч (по слайду ≈ 1200 ч)", "0.00")
        ],
        "Результаты: Pc(100) = 0,9215; интенсивность отказов λc = 0,817·10⁻³ 1/ч; среднее время безотказной работы mtc ≈ 1223 ч (по слайдам 15–16 округленно 1200 ч).",
        "typ4"
    )

    # Типовая 5
    render_task(
        ws_typ,
        "ЗАДАЧА № 5. Приближенный расчет надежности системы из 100 элементов (P = 0,9997)",
        "P(t) = 0,9997, n = 100. Найти вероятность безотказной работы системы Pc(t).",
        [
            ("Вероятность элемента", "P_el", 0.9997, "—", "P(t) = 0.9997"),
            ("Число элементов", "n", 100, "шт.", "n = 100 шт.")
        ],
        [
            ("Вероятность отказа элемента q", "q_el", "1 - P_el", "=1-{P_el}", "—", "q = 1 - 0.9997 = 0.0003", "0.000000"),
            ("Приближенная вероятность по (3.18)", "Pc_appr", "1 - n · q", "=1-{n}*{q_el}", "—", "Pc ≈ 1 - 100 · 0.0003 = 0.9700", "0.0000"),
            ("Точная вероятность безотказной работы", "Pc_exact", "P_el ^ n", "={P_el}^{n}", "—", "Pc = 0.9997¹⁰⁰ ≈ 0.97044", "0.0000")
        ],
        "Результаты: по приближенной формуле Pc(t) ≈ 0,9700; точное значение Pc(t) = 0,9704 (совпадение со слайдом 16).",
        "typ5"
    )

    # Типовая 6
    render_task(
        ws_typ,
        "ЗАДАЧА № 6. Определение надежности элемента по системе из n = 120 элементов (Pc = 0,95)",
        "Pc(t) = 0,95, n = 120 равнонадежных элементов. Найти вероятность безотказной работы элемента Pi(t).",
        [
            ("Вероятность системы Pc", "Pc", 0.95, "—", "Pc(t) = 0.95"),
            ("Число элементов", "n", 120, "шт.", "n = 120 элементов")
        ],
        [
            ("Вероятность отказа системы qc", "qc", "1 - Pc", "=1-{Pc}", "—", "qc = 1 - 0.95 = 0.05", "0.0000"),
            ("Приближенная вероятность элемента", "P_el_appr", "1 - qc / n", "=1-{qc}/{n}", "—", "Pi ≈ 1 - 0.05 / 120 ≈ 0.999583 (по слайду 0.9996)", "0.000000"),
            ("Точная вероятность элемента", "P_el_exact", "Pc ^ (1/n)", "={Pc}^(1/{n})", "—", "Pi = 0.95^(1/120) ≈ 0.999573", "0.000000")
        ],
        "Результаты: Pi(t) ≈ 0,9996 (полное совпадение со слайдами 16–17).",
        "typ6"
    )

    # Типовая 7
    render_task(
        ws_typ,
        "ЗАДАЧА № 7. Вероятность безотказной работы системы из 12600 элементов (t = 50 ч)",
        "n = 12600 элементов, λср = 0,32∙10⁻⁶ ч⁻¹. Определить Pc(50) при t = 50 ч.",
        [
            ("Число элементов", "n", 12600, "шт.", "n = 12600 шт."),
            ("Интенсивность элемента", "l_elem", 0.00000032, "1/ч", "λср = 0.32·10⁻⁶ ч⁻¹"),
            ("Время наработки", "t", 50, "ч", "t = 50 ч")
        ],
        [
            ("Интенсивность системы λc", "lc", "n · λср", "={n}*{l_elem}", "1/ч", "λc = 12600 · 0.32·10⁻⁶ = 0.004032 ч⁻¹", "0.000000"),
            ("Вероятность безотказной работы Pc(50)", "Pc_50", "EXP(-λc · t)", "=EXP(-{lc}*{t})", "—", "Pc(50) = exp(-0.2016) ≈ 0.8174 (по слайду ≈ 0.82)", "0.0000")
        ],
        "Результаты: Pc(50) = 0,8174 ≈ 0,82 (полное совпадение со слайдом 17).",
        "typ7"
    )

    # -------------------------------------------------------------------------
    # ЛИСТ 1: СВОДКА И ОТВЕТЫ
    # -------------------------------------------------------------------------
    ws_sum = wb.create_sheet(title="Сводка и Ответы", index=0)
    ws_sum.views.sheetView[0].showGridLines = True
    make_header(
        ws_sum,
        "ПРАКТИЧЕСКАЯ РАБОТА № 3 — СВОДНЫЕ РЕЗУЛЬТАТЫ",
        "Последовательное соединение элементов в систему (расчет количественных характеристик надежности)",
        "Дисциплина: Надежность автоматизированных систем (НАС)  |  Преподаватель: Ризванов  |  Выполнил: Студент Гайнанов Д.И.",
        max_c=8
    )
    
    ws_sum.cell(6, 2, "1. РЕЗУЛЬТАТЫ РЕШЕНИЯ ЗАДАЧ ДЛЯ САМОСТОЯТЕЛЬНОГО РЕШЕНИЯ (№ 8 – 13)").font = f_sec
    ws_sum.row_dimensions[6].height = 24
    
    sum_headers = ["№ задачи", "Краткое условие задачи", "Определяемый показатель", "Математическая формула", "Результат вычислений", "Ед. изм.", "Ссылка на решение", "Статус"]
    for ci, h in enumerate(sum_headers, start=2):
        cell = ws_sum.cell(7, ci, h)
        cell.font = f_th; cell.fill = fill_th; cell.alignment = al_c_wrap; cell.border = box_cell
    ws_sum.row_dimensions[7].height = 26
    
    summary_data_gainanov = [
        ("Задача 8", "n=2000, λср=0.33·10⁻⁵ ч⁻¹, t=200 ч. Найти Pc(200)", "Pc(200)", "Pc = exp(-n · λср · t)", f"={RESULTS_REF['t8_Pc_200']}", "—", RESULTS_REF['t8_start'], "0.0000"),
        ("Задача 8 (2)", "Та же задача, средняя наработка системы mtc", "mtc", "mtc = 1 / (n · λср)", f"={RESULTS_REF['t8_mtc']}", "ч", RESULTS_REF['t8_start'], "0.00"),
        ("Задача 9", "ЭВМ из 200000 элементов, λср=0.2·10⁻⁶ ч⁻¹, сутки (t=24 ч)", "Pc(24)", "Pc = exp(-n · λср · 24)", f"={RESULTS_REF['t9_Pc_24']}", "—", RESULTS_REF['t9_start'], "0.0000"),
        ("Задача 9 (2)", "Та же задача, средняя наработка ЭВМ mtc", "mtc", "mtc = 1 / (n · λср)", f"={RESULTS_REF['t9_mtc']}", "ч", RESULTS_REF['t9_start'], "0.00"),
        ("Задача 10", "n=6000, λср=0.16·10⁻⁶ ч⁻¹, t=50 ч. Найти Pc(50)", "Pc(50)", "Pc = exp(-n · λср · t)", f"={RESULTS_REF['t10_Pc_50']}", "—", RESULTS_REF['t10_start'], "0.0000"),
        ("Задача 10 (2)", "Та же задача, средняя наработка системы mtc", "mtc", "mtc = 1 / (n · λср)", f"={RESULTS_REF['t10_mtc']}", "ч", RESULTS_REF['t10_start'], "0.00"),
        ("Задача 11", "Прибор из 5 узлов (Pi = {0.98, 0.99, 0.998, 0.975, 0.985})", "Pc (точная)", "Pc = Π Pi", f"={RESULTS_REF['t11_Pc_exact']}", "—", RESULTS_REF['t11_start'], "0.0000"),
        ("Задача 11 (2)", "Та же задача, приближенное значение по формуле (3.18)", "Pc (прибл.)", "Pc ≈ 1 - Σ(1 - Pi)", f"={RESULTS_REF['t11_Pc_approx']}", "—", RESULTS_REF['t11_start'], "0.0000"),
        ("Задача 12", "Система из 5 приборов (mti = {83, 220, 280, 400, 700} ч)", "λc", "λc = Σ(1/mti)", f"={RESULTS_REF['t12_lc']}", "1/ч", RESULTS_REF['t12_start'], "0.000000"),
        ("Задача 12 (2)", "Та же задача, средняя наработка системы mtc", "mtc", "mtc = 1 / λc", f"={RESULTS_REF['t12_mtc']}", "ч", RESULTS_REF['t12_start'], "0.00"),
        ("Задача 13", "5 блоков с Pi={0.98..0.985} за 50 ч. Найти mtc прибора", "mtc (точная)", "mtc = -t / ln(Pc)", f"={RESULTS_REF['t13_mtc']}", "ч", RESULTS_REF['t13_start'], "0.00"),
        ("Задача 13 (2)", "Та же задача, приближенное значение mtc по (3.18)", "mtc (прибл.)", "mtc ≈ t / Σ qi", f"={RESULTS_REF['t13_mtc_appr']}", "ч", RESULTS_REF['t13_start'], "0.00")
    ]
    
    row = 7
    for item in summary_data_gainanov:
        row += 1
        ws_sum.row_dimensions[row].height = 20
        bg = fill_zebra if row % 2 == 0 else fill_white
        
        c1 = ws_sum.cell(row, 2, item[0]); c1.font = f_bold; c1.alignment = al_c; c1.border = box_cell; c1.fill = bg
        c2 = ws_sum.cell(row, 3, item[1]); c2.font = f_reg; c2.alignment = al_l; c2.border = box_cell; c2.fill = bg
        c3 = ws_sum.cell(row, 4, item[2]); c3.font = f_bold; c3.alignment = al_c; c3.border = box_cell; c3.fill = bg
        c4 = ws_sum.cell(row, 5, item[3]); c4.font = f_math; c4.alignment = al_l; c4.border = box_cell; c4.fill = bg
        c5 = ws_sum.cell(row, 6, item[4]); c5.font = f_res; c5.fill = fill_res; c5.alignment = al_r; c5.border = box_res; c5.number_format = item[7]
        c6 = ws_sum.cell(row, 7, item[5]); c6.font = f_reg; c6.alignment = al_c; c6.border = box_cell; c6.fill = bg
        c7 = ws_sum.cell(row, 8, "Смотреть"); c7.hyperlink = f"#{item[6]}"; c7.font = Font(name=FONT, size=9, underline='single', color='1F4E78'); c7.alignment = al_c; c7.border = box_cell; c7.fill = bg
        c8 = ws_sum.cell(row, 9, "Решено"); c8.font = f_stat; c8.alignment = al_c; c8.border = box_cell; c8.fill = bg

    # Таблица типовых задач
    row += 3
    ws_sum.cell(row, 2, "2. РЕЗУЛЬТАТЫ РЕШЕНИЯ ТИПОВЫХ ЗАДАЧ ИЗ ПРЕЗЕНТАЦИИ (№ 1 – 7)").font = f_sec
    ws_sum.row_dimensions[row].height = 24
    row += 1
    typ_heads = ["№ задачи", "Тема / Объект расчета", "Ключевые исходные данные", "Искомый параметр", "Результат вычислений", "Ед. изм.", "Сверка со слайдом", "Ссылка"]
    for ci, h in enumerate(typ_heads, start=2):
        cell = ws_sum.cell(row, ci, h)
        cell.font = f_th; cell.fill = fill_th; cell.alignment = al_c_wrap; cell.border = box_cell
    ws_sum.row_dimensions[row].height = 24
    
    typ_sum_data = [
        ("Задача 1", "3 устройства (разные λ(t))", "λ1=const, λ2=k2·t, λ3=k3·t^2.6, t=100 ч", "Pc(100)", f"={RESULTS_REF['typ1_Pc_100']}", "—", "Pc=0.8543 (на слайде 0.33)", RESULTS_REF['typ1_start']),
        ("Задача 2", "3 блока (mt1=160, mt2=320, mt3=600 ч)", "Экспоненциальный закон надежности", "mtc", f"={RESULTS_REF['typ2_mtc']}", "ч", "mtc=90.57 ч ≈ 91 ч (Совпадает)", RESULTS_REF['typ2_start']),
        ("Задача 3", "Система из 12600 элементов", "λср=0.32·10⁻⁶ ч⁻¹, t=50 ч", "Pc(50), mtc", f"={RESULTS_REF['typ3_Pc_50']}", "—", "Pc=0.8174; mtc=248 ч (Совпадает)", RESULTS_REF['typ3_start']),
        ("Задача 4", "2 устройства (P1=0.95, P2=0.97)", "t=100 ч, экспоненциальный закон", "mtc", f"={RESULTS_REF['typ4_mtc']}", "ч", "mtc=1223 ч ≈ 1200 ч (Совпадает)", RESULTS_REF['typ4_start']),
        ("Задача 5", "Система из 100 элементов (P=0.9997)", "Приближенный расчет по формуле (3.18)", "Pc(t)", f"={RESULTS_REF['typ5_Pc_appr']}", "—", "Pc=0.9700 (Совпадает)", RESULTS_REF['typ5_start']),
        ("Задача 6", "Определение Pi(t) по Pc=0.95 (n=120)", "Приближенный расчет по формуле (3.18)", "Pi(t)", f"={RESULTS_REF['typ6_P_el_appr']}", "—", "Pi=0.9996 (Совпадает)", RESULTS_REF['typ6_start']),
        ("Задача 7", "Система из 12600 элементов", "λср=0.32·10⁻⁶ ч⁻¹, t=50 ч", "Pc(50)", f"={RESULTS_REF['typ7_Pc_50']}", "—", "Pc=0.8174 ≈ 0.82 (Совпадает)", RESULTS_REF['typ7_start'])
    ]
    for item in typ_sum_data:
        row += 1
        ws_sum.row_dimensions[row].height = 20
        bg = fill_zebra if row % 2 == 0 else fill_white
        c1 = ws_sum.cell(row, 2, item[0]); c1.font = f_bold; c1.alignment = al_c; c1.border = box_cell; c1.fill = bg
        c2 = ws_sum.cell(row, 3, item[1]); c2.font = f_reg; c2.alignment = al_l; c2.border = box_cell; c2.fill = bg
        c3 = ws_sum.cell(row, 4, item[2]); c3.font = f_reg; c3.alignment = al_l; c3.border = box_cell; c3.fill = bg
        c4 = ws_sum.cell(row, 5, item[3]); c4.font = f_bold; c4.alignment = al_c; c4.border = box_cell; c4.fill = bg
        c5 = ws_sum.cell(row, 6, item[4]); c5.font = f_res; c5.fill = fill_res; c5.alignment = al_r; c5.border = box_res; c5.number_format = '0.0000'
        c6 = ws_sum.cell(row, 7, item[5]); c6.font = f_reg; c6.alignment = al_c; c6.border = box_cell; c6.fill = bg
        c7 = ws_sum.cell(row, 8, item[6]); c7.font = f_stat; c7.alignment = al_l; c7.border = box_cell; c7.fill = bg
        c8 = ws_sum.cell(row, 9, "Смотреть"); c8.hyperlink = f"#{item[7]}"; c8.font = Font(name=FONT, size=9, underline='single', color='1F4E78'); c8.alignment = al_c; c8.border = box_cell; c8.fill = bg

    # -------------------------------------------------------------------------
    # ЛИСТ 4: ТЕОРИЯ И ФОРМУЛЫ
    # -------------------------------------------------------------------------
    ws_th = wb.create_sheet(title="Теория и формулы")
    ws_th.views.sheetView[0].showGridLines = True
    make_header(
        ws_th,
        "ТЕОРЕТИЧЕСКИЙ СПРАВОЧНИК — ПОСЛЕДОВАТЕЛЬНЫЕ СИСТЕМЫ",
        "Математические модели надежности систем при последовательном соединении элементов (слайды 12–13)",
        "Определения, расчетные соотношения, формулы приближенных вычислений",
        max_c=7
    )
    
    ws_th.cell(6, 2, "ОСНОВНЫЕ РАСЧЕТНЫЕ СООТНОШЕНИЯ").font = f_sec
    ws_th.row_dimensions[6].height = 24
    
    r = 7
    for ci, h in enumerate(["Показатель надежности", "Общая формула", "Экспоненциальный закон", "Формула через наработки", "Ед. изм.", "Пояснение"], start=2):
        cell = ws_th.cell(r, ci, h)
        cell.font = f_th; cell.fill = fill_th; cell.alignment = al_c_wrap; cell.border = box_cell
    ws_th.row_dimensions[r].height = 24
    
    th_rows = [
        ("Вероятность безотказной работы Pc(t)", "Pc(t) = Π Pi(t) = exp(-∫λc(t)dt)", "Pc(t) = exp(-λc · t)", "Pc(t) = exp(-t · Σ(1/mti))", "—", "Отказ любого элемента вызывает отказ всей системы"),
        ("Интенсивность отказов системы λc(t)", "λc(t) = Σ λi(t)", "λc = Σ λi = const", "λc = Σ (1 / mti)", "1/ч", "Теорема сложения интенсивностей отказов элементов"),
        ("Вероятность отказа системы qc(t)", "qc(t) = 1 - Pc(t)", "qc(t) = 1 - exp(-λc · t)", "qc(t) = 1 - exp(-t / mtc)", "—", "Вероятность хотя бы одного отказа за время t"),
        ("Частота отказов системы fc(t)", "fc(t) = -dPc(t)/dt", "fc(t) = λc · exp(-λc · t)", "fc(t) = λc · Pc(t)", "1/ч", "Плотность вероятности первого отказа системы"),
        ("Среднее время безотказной работы mtc", "mtc = ∫ Pc(t) dt", "mtc = 1 / λc", "mtc = 1 / (Σ 1 / mti)", "ч", "Средняя наработка системы всегда меньше наработки наименее надежного элемента!")
    ]
    for row_data in th_rows:
        r += 1
        ws_th.row_dimensions[r].height = 26
        bg = fill_zebra if r % 2 == 0 else fill_white
        c1 = ws_th.cell(r, 2, row_data[0]); c1.font = f_bold; c1.alignment = al_l; c1.border = box_cell; c1.fill = bg
        c2 = ws_th.cell(r, 3, row_data[1]); c2.font = f_math; c2.alignment = al_l; c2.border = box_cell; c2.fill = bg
        c3 = ws_th.cell(r, 4, row_data[2]); c3.font = f_bold; c3.alignment = al_l; c3.border = box_cell; c3.fill = bg
        c4 = ws_th.cell(r, 5, row_data[3]); c4.font = f_math; c4.alignment = al_l; c4.border = box_cell; c4.fill = bg
        c5 = ws_th.cell(r, 6, row_data[4]); c5.font = f_reg; c5.alignment = al_c; c5.border = box_cell; c5.fill = bg
        c6 = ws_th.cell(r, 7, row_data[5]); c6.font = f_cond; c6.alignment = al_l_wrap; c6.border = box_cell; c6.fill = bg

    r += 2
    ws_th.cell(r, 2, "ПРИБЛИЖЕННЫЕ ФОРМУЛЫ ПРИ ВЫСОКОЙ НАДЕЖНОСТИ (Pi(t) ≈ 1)").font = f_sec
    ws_th.row_dimensions[r].height = 24
    
    approx_rows = [
        ("Произведение вероятностей разнонадежных узлов:", "P1 · P2 · ... · Pn ≈ 1 - Σ qi(t) = 1 - Σ (1 - Pi(t)). Погрешность мала, если Σ qi << 1."),
        ("Степень вероятности для n одинаковых элементов:", "[P(t)]ⁿ = [1 - q(t)]ⁿ ≈ 1 - n · q(t). Удобно для систем с тысячами элементов."),
        ("Определение надежности элемента по системе:", "P(t) = [Pc(t)]^(1/n) ≈ 1 - qc(t) / n. Позволяет задавать требования к элементам системы.")
    ]
    for a_title, a_desc in approx_rows:
        r += 1
        ws_th.row_dimensions[r].height = 24
        ws_th.cell(r, 2, a_title).font = f_bold; ws_th.cell(r, 2).alignment = al_l
        ws_th.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        c_desc = ws_th.cell(r, 3, a_desc)
        c_desc.font = f_reg; c_desc.alignment = al_l_wrap
        apply_style(ws_th, r, 2, 7, fill=fill_card, border=box_cell)

    # Ширина колонок
    col_w = {
        "Сводка и Ответы": {"A": 3, "B": 15, "C": 38, "D": 22, "E": 28, "F": 18, "G": 10, "H": 18, "I": 12},
        "Задачи 8-13": {"A": 3, "B": 36, "C": 16, "D": 30, "E": 22, "F": 12, "G": 44},
        "Задачи 1-7": {"A": 3, "B": 36, "C": 16, "D": 30, "E": 22, "F": 12, "G": 44},
        "Теория и формулы": {"A": 3, "B": 34, "C": 28, "D": 24, "E": 24, "F": 10, "G": 46}
    }
    for sheet_name, widths in col_w.items():
        if sheet_name in wb.sheetnames:
            ws_cur = wb[sheet_name]
            for col_letter, width in widths.items():
                ws_cur.column_dimensions[col_letter].width = width

    fname = "НАС_ПР3_ГайнановДИ.xlsx"
    wb.save(fname)
    print(f"Created: {fname}")


# =============================================================================
# ФАЙЛ ДЛЯ ГАРИФУЛЛИНА (ПР3): Изумрудно-зеленый стиль, 3 листа
# =============================================================================

def build_pr3_garifullin():
    wb = openpyxl.Workbook()
    FONT = 'Calibri'
    
    # Стили шрифтов (Хвоя / Лес)
    f_title = Font(name=FONT, size=13, bold=True, color='FFFFFF')
    f_subtitle = Font(name=FONT, size=10, bold=False, color='E8F5E9')
    f_meta = Font(name=FONT, size=9, bold=False, color='C8E6C9')
    f_sec = Font(name=FONT, size=11, bold=True, color='2D5A27')
    f_th = Font(name=FONT, size=9, bold=True, color='FFFFFF')
    f_reg = Font(name=FONT, size=9, bold=False, color='202020')
    f_bold = Font(name=FONT, size=9, bold=True, color='202020')
    f_math = Font(name=FONT, size=9, italic=True, color='2D5A27')
    f_res = Font(name=FONT, size=10, bold=True, color='1B5E20')
    f_cond = Font(name=FONT, size=9, italic=True, color='4A5568')
    f_ans = Font(name=FONT, size=9, bold=True, color='1B5E20')
    
    # Цвета (Зеленая гамма)
    fill_forest = PatternFill(start_color='2D5A27', end_color='2D5A27', fill_type='solid')
    fill_sub = PatternFill(start_color='3E7A36', end_color='3E7A36', fill_type='solid')
    fill_th = PatternFill(start_color='4E8A46', end_color='4E8A46', fill_type='solid')
    fill_card = PatternFill(start_color='F2F7F2', end_color='F2F7F2', fill_type='solid')
    fill_res = PatternFill(start_color='D9EED9', end_color='D9EED9', fill_type='solid')
    fill_zebra = PatternFill(start_color='F9FCF9', end_color='F9FCF9', fill_type='solid')
    fill_ans = PatternFill(start_color='E8F5E9', end_color='E8F5E9', fill_type='solid')
    fill_white = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    
    bd_gray = Side(style='thin', color='C2D6C2')
    bd_green = Side(style='double', color='2D5A27')
    bd_res = Side(style='thin', color='4E8A46')
    bd_med_green = Side(style='medium', color='2D5A27')
    
    box_cell = Border(left=bd_gray, right=bd_gray, top=bd_gray, bottom=bd_gray)
    box_total = Border(left=bd_gray, right=bd_gray, top=bd_gray, bottom=bd_green)
    box_res = Border(left=bd_res, right=bd_res, top=bd_res, bottom=bd_res)
    
    al_c = Alignment(horizontal='center', vertical='center')
    al_c_wrap = Alignment(horizontal='center', vertical='center', wrap_text=True)
    al_l = Alignment(horizontal='left', vertical='center')
    al_l_wrap = Alignment(horizontal='left', vertical='center', wrap_text=True)
    al_r = Alignment(horizontal='right', vertical='center')

    def apply_style(ws, r, c1, c2, font=None, fill=None, border=None, al=None):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            if font: cell.font = font
            if fill: cell.fill = fill
            if border: cell.border = border
            if al: cell.alignment = al

    def make_header_gar(ws, title, sub, meta, max_c=7):
        ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=max_c)
        ws.cell(2, 2, title).font = f_title; ws.cell(2, 2).alignment = al_c
        apply_style(ws, 2, 2, max_c, fill=fill_forest)
        ws.row_dimensions[2].height = 24
        
        ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=max_c)
        ws.cell(3, 2, sub).font = f_subtitle; ws.cell(3, 2).alignment = al_c
        apply_style(ws, 3, 2, max_c, fill=fill_sub)
        ws.row_dimensions[3].height = 18
        
        ws.merge_cells(start_row=4, start_column=2, end_row=4, end_column=max_c)
        ws.cell(4, 2, meta).font = f_meta; ws.cell(4, 2).alignment = al_c
        apply_style(ws, 4, 2, max_c, fill=fill_sub)
        ws.row_dimensions[4].height = 18

    # ЛИСТ 2: РАСЧЕТНАЯ ЧАСТЬ (ЗАДАНИЯ 8-13)
    ws_sol = wb.active
    ws_sol.title = "Расчетная часть (№8-13)"
    ws_sol.views.sheetView[0].showGridLines = True
    
    make_header_gar(
        ws_sol,
        "РАСЧЕТНАЯ ЧАСТЬ ПРАКТИЧЕСКОЙ РАБОТЫ № 3 (ЗАДАНИЯ 8 – 13)",
        "Последовательное соединение элементов в систему  |  Студент: Гарифуллин",
        "Расчет показателей надежности формулами Excel со ссылками на исходные ячейки",
        max_c=6
    )
    
    cur_r = 6
    results_map = {}
    
    def add_card(title, cond, inputs, calcs, note, task_id):
        nonlocal cur_r
        r = cur_r
        
        ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws_sol.cell(r, 2, title).font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
        ws_sol.cell(r, 2).alignment = al_l
        apply_style(ws_sol, r, 2, 6, fill=fill_sub)
        ws_sol.row_dimensions[r].height = 20
        
        r += 1
        ws_sol.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=6)
        c_c = ws_sol.cell(r, 2, f"Условие задания: {cond}")
        c_c.font = f_cond; c_c.alignment = al_l_wrap
        apply_style(ws_sol, r, 2, 6, fill=fill_card)
        apply_style(ws_sol, r+1, 2, 6, fill=fill_card)
        ws_sol.row_dimensions[r].height = 16
        ws_sol.row_dimensions[r+1].height = 16
        r += 1
        
        # Исходные данные
        r += 1
        ws_sol.cell(r, 2, "1. Исходные данные:").font = f_bold
        r += 1
        for ci, h in enumerate(["Параметр", "Обозначение", "Значение", "Размерность"], start=2):
            cell = ws_sol.cell(r, ci, h)
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
            
            c1 = ws_sol.cell(r, 2, p_name); c1.font = f_reg; c1.border = box_cell
            c2 = ws_sol.cell(r, 3, p_sym); c2.font = f_bold; c2.alignment = al_c; c2.border = box_cell
            c3 = ws_sol.cell(r, 4, p_val); c3.font = f_bold; c3.alignment = al_r; c3.border = box_cell
            if isinstance(p_val, (int, float)):
                if abs(p_val) < 0.0001 and p_val != 0:
                    c3.number_format = '0.00E+00'
                elif isinstance(p_val, float):
                    c3.number_format = '#,##0.####'
                else:
                    c3.number_format = '#,##0'
            c4 = ws_sol.cell(r, 5, p_unit); c4.font = f_reg; c4.alignment = al_c; c4.border = box_cell
            
            inp_cells[p_sym] = f"D{r}"
            inp_cells[p_sym.replace('(', '_').replace(')', '')] = f"D{r}"
            
        # Расчет
        r += 2
        ws_sol.cell(r, 2, "2. Формульный расчет:").font = f_bold
        r += 1
        for ci, h in enumerate(["Искомый показатель", "Матем. формула", "Значение в Excel", "Размерность"], start=2):
            cell = ws_sol.cell(r, ci, h)
            cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
            cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
        ws_sol.row_dimensions[r].height = 18
        
        calc_cells = {}
        for row_c in calcs:
            r += 1
            ws_sol.row_dimensions[r].height = 19
            c_name, c_math, c_template, c_unit, c_fmt, c_key = row_c
            formula = c_template.format(**inp_cells, **calc_cells)
            
            c1 = ws_sol.cell(r, 2, c_name); c1.font = f_reg; c1.border = box_cell
            c2 = ws_sol.cell(r, 3, c_math); c2.font = f_math; c2.alignment = al_c; c2.border = box_cell
            c3 = ws_sol.cell(r, 4, formula); c3.font = f_res; c3.fill = fill_res; c3.alignment = al_r; c3.border = box_res; c3.number_format = c_fmt
            c4 = ws_sol.cell(r, 5, c_unit); c4.font = f_reg; c4.alignment = al_c; c4.border = box_cell
            
            calc_cells[c_key] = f"D{r}"
            results_map[f"{task_id}_{c_key}"] = f"D{r}"
            
        r += 1
        ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        ans_cell = ws_sol.cell(r, 2, f"Ответ: {note}")
        ans_cell.font = f_ans
        apply_style(ws_sol, r, 2, 5, fill=fill_ans, border=box_cell, al=al_l)
        ws_sol.row_dimensions[r].height = 20
        cur_r = r + 2

    add_card(
        "Задание 8. Аппаратура связи (2000 элементов, t = 200 ч)",
        "Аппаратура связи: n = 2000 элементов, λср = 0,33∙10⁻⁵ ч⁻¹. Определить Pc(200) и mtc.",
        [("Число элементов", "n", 2000, "шт."), ("Интенсивность λср", "l_elem", 0.0000033, "1/ч"), ("Время t", "t", 200, "ч")],
        [
            ("Интенсивность системы λc", "n · λср", "={n}*{l_elem}", "1/ч", "0.000000", "lc"),
            ("Вероятность безотказной работы", "exp(-λc·t)", "=EXP(-{lc}*{t})", "—", "0.0000", "pc"),
            ("Среднее время работы системы", "1 / λc", "=1/{lc}", "ч", "0.00", "mtc")
        ],
        "Pc(200) = 0,2671 (26,71%); mtc = 151,52 ч.",
        "t8"
    )

    add_card(
        "Задание 9. ЭВМ из 200000 элементов (сутки, t = 24 ч)",
        "Невосстанавливаемая машина: n = 200000, λср = 0,2∙10⁻⁶ ч⁻¹. Найти Pc(24) и mtc.",
        [("Число элементов", "n", 200000, "шт."), ("Интенсивность λср", "l_elem", 0.0000002, "1/ч"), ("Сутки t", "t", 24, "ч")],
        [
            ("Интенсивность системы λc", "n · λср", "={n}*{l_elem}", "1/ч", "0.000000", "lc"),
            ("Вероятность безотказной работы", "exp(-λc·t)", "=EXP(-{lc}*{t})", "—", "0.0000", "pc"),
            ("Среднее время работы ЭВМ", "1 / λc", "=1/{lc}", "ч", "0.00", "mtc")
        ],
        "Pc(24) = 0,3829 (38,29%); mtc = 25,00 ч.",
        "t9"
    )

    add_card(
        "Задание 10. Система управления (6000 элементов, t = 50 ч)",
        "Система управления: n = 6000, λср = 0,16∙10⁻⁶ ч⁻¹. Определить Pc(50) и mtc.",
        [("Число элементов", "n", 6000, "шт."), ("Интенсивность λср", "l_elem", 0.00000016, "1/ч"), ("Время t", "t", 50, "ч")],
        [
            ("Интенсивность системы λc", "n · λср", "={n}*{l_elem}", "1/ч", "0.000000", "lc"),
            ("Вероятность безотказной работы", "exp(-λc·t)", "=EXP(-{lc}*{t})", "—", "0.0000", "pc"),
            ("Средняя наработка системы", "1 / λc", "=1/{lc}", "ч", "0.00", "mtc")
        ],
        "Pc(50) = 0,9531 (95,31%); mtc = 1041,67 ч.",
        "t10"
    )

    add_card(
        "Задание 11. Надежность прибора из 5 узлов",
        "P1 = 0,98; P2 = 0,99; P3 = 0,998; P4 = 0,975; P5 = 0,985. Определить надежность прибора.",
        [
            ("Узел 1", "P1", 0.98, "—"), ("Узел 2", "P2", 0.99, "—"), ("Узел 3", "P3", 0.998, "—"),
            ("Узел 4", "P4", 0.975, "—"), ("Узел 5", "P5", 0.985, "—")
        ],
        [
            ("Точная вероятность Pc", "Π Pi", "={P1}*{P2}*{P3}*{P4}*{P5}", "—", "0.0000", "pc"),
            ("Приближенная вероятность по (3.18)", "1 - Σ(1 - Pi)", "=1-((1-{P1})+(1-{P2})+(1-{P3})+(1-{P4})+(1-{P5}))", "—", "0.0000", "pc_appr")
        ],
        "Pc = 0,9301 (приближенно 0,9280).",
        "t11"
    )

    # Задание 12
    r = cur_r
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws_sol.cell(r, 2, "Задание 12. Система из 5 приборов").font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    apply_style(ws_sol, r, 2, 6, fill=fill_sub, al=al_l)
    ws_sol.row_dimensions[r].height = 20
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=6)
    ws_sol.cell(r, 2, "Условие: mt1 = 83 ч, mt2 = 220 ч, mt3 = 280 ч, mt4 = 400 ч, mt5 = 700 ч. Экспоненциальный закон надежности. Найти mtc.").font = f_cond
    ws_sol.cell(r, 2).alignment = al_l_wrap
    apply_style(ws_sol, r, 2, 6, fill=fill_card)
    apply_style(ws_sol, r+1, 2, 6, fill=fill_card)
    ws_sol.row_dimensions[r].height = 16
    ws_sol.row_dimensions[r+1].height = 16
    r += 1
    
    r += 1
    ws_sol.cell(r, 2, "1. Наработки и интенсивности приборов:").font = f_bold
    r += 1
    for ci, h in enumerate(["№", "Обозначение", "mti (ч)", "λi = 1/mti (1/ч)", "Примечание"], start=2):
        cell = ws_sol.cell(r, ci, h)
        cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
        cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws_sol.row_dimensions[r].height = 18
    
    devs_gar = [("1", "mt1", 83, "Прибор 1"), ("2", "mt2", 220, "Прибор 2"), ("3", "mt3", 280, "Прибор 3"), ("4", "mt4", 400, "Прибор 4"), ("5", "mt5", 700, "Прибор 5")]
    st_dg = r + 1
    for dg in devs_gar:
        r += 1
        ws_sol.row_dimensions[r].height = 18
        ws_sol.cell(r, 2, dg[0]).font = f_reg; ws_sol.cell(r, 2).alignment = al_c; ws_sol.cell(r, 2).border = box_cell
        ws_sol.cell(r, 3, dg[1]).font = f_bold; ws_sol.cell(r, 3).alignment = al_c; ws_sol.cell(r, 3).border = box_cell
        ws_sol.cell(r, 4, dg[2]).font = f_bold; ws_sol.cell(r, 4).alignment = al_r; ws_sol.cell(r, 4).border = box_cell; ws_sol.cell(r, 4).number_format = '#,##0'
        c_lg = ws_sol.cell(r, 5, f"=1/D{r}"); c_lg.font = f_bold; c_lg.alignment = al_r; c_lg.border = box_cell; c_lg.number_format = '0.000000'
        ws_sol.cell(r, 6, dg[3]).font = f_cond; ws_sol.cell(r, 6).border = box_cell
    end_dg = r
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws_sol.cell(r, 2, "Интенсивность отказов системы λc = Σ λi :").font = f_bold
    ws_sol.cell(r, 2).alignment = al_r
    apply_style(ws_sol, r, 2, 4, fill=fill_card, border=box_total)
    c_slcg = ws_sol.cell(r, 5, f"=SUM(E{st_dg}:E{end_dg})")
    c_slcg.font = f_res; c_slcg.alignment = al_r; c_slcg.fill = fill_res; c_slcg.border = box_total; c_slcg.number_format = '0.000000'
    ws_sol.cell(r, 6, "1/ч").font = f_bold; ws_sol.cell(r, 6).alignment = al_c; ws_sol.cell(r, 6).border = box_total
    ref_lc_gar = f"E{r}"
    results_map["t12_lc"] = ref_lc_gar
    
    r += 2
    ws_sol.cell(r, 2, "2. Итоговое время безотказной работы системы:").font = f_bold
    r += 1
    for ci, h in enumerate(["Искомый показатель", "Матем. формула", "Значение в Excel", "Размерность"], start=2):
        cell = ws_sol.cell(r, ci, h)
        cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
        cell.fill = fill_th; cell.alignment = al_c; cell.border = box_cell
    ws_sol.row_dimensions[r].height = 18
    
    r += 1
    ws_sol.row_dimensions[r].height = 19
    ws_sol.cell(r, 2, "Среднее время работы системы mtc").font = f_reg; ws_sol.cell(r, 2).border = box_cell
    ws_sol.cell(r, 3, "1 / λc").font = f_math; ws_sol.cell(r, 3).alignment = al_c; ws_sol.cell(r, 3).border = box_cell
    c_fgar = ws_sol.cell(r, 4, f"=1/{ref_lc_gar}")
    c_fgar.font = f_res; c_fgar.fill = fill_res; c_fgar.alignment = al_r; c_fgar.border = box_res; c_fgar.number_format = '0.00'
    ws_sol.cell(r, 5, "ч").font = f_reg; ws_sol.cell(r, 5).alignment = al_c; ws_sol.cell(r, 5).border = box_cell
    results_map["t12_mtc"] = f"D{r}"
    
    r += 1
    ws_sol.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    ws_sol.cell(r, 2, "Ответ: λc = 0,024094 1/ч; mtc = 41,50 ч.").font = f_ans
    apply_style(ws_sol, r, 2, 5, fill=fill_ans, border=box_cell, al=al_l)
    ws_sol.row_dimensions[r].height = 20
    cur_r = r + 2

    add_card(
        "Задание 13. Средняя наработка прибора из 5 блоков (t = 50 ч)",
        "P1 = 0,98; P2 = 0,99; P3 = 0,998; P4 = 0,975; P5 = 0,985 при t = 50 ч. Экспоненциальный закон. Найти mtc.",
        [
            ("Блок 1", "P1", 0.98, "—"), ("Блок 2", "P2", 0.99, "—"), ("Блок 3", "P3", 0.998, "—"),
            ("Блок 4", "P4", 0.975, "—"), ("Блок 5", "P5", 0.985, "—"), ("Время t", "t", 50, "ч")
        ],
        [
            ("Вероятность прибора Pc(50)", "Π Pi", "={P1}*{P2}*{P3}*{P4}*{P5}", "—", "0.0000", "pc"),
            ("Интенсивность отказов λc", "-LN(Pc) / t", "=-LN({pc})/{t}", "1/ч", "0.000000", "lc"),
            ("Среднее время работы прибора mtc", "1 / λc", "=1/{lc}", "ч", "0.00", "mtc")
        ],
        "Pc(50) = 0,9301; λc = 0,001450 1/ч; mtc = 689,57 ч.",
        "t13"
    )

    col_w_sol = {"A": 3, "B": 28, "C": 18, "D": 22, "E": 14, "F": 18}
    for col, width in col_w_sol.items():
        ws_sol.column_dimensions[col].width = width

    # ЛИСТ 1: ИТОГОВАЯ ВЕДОМОСТЬ (СВОДКА)
    ws_sum_g = wb.create_sheet(title="Итоговая ведомость", index=0)
    ws_sum_g.views.sheetView[0].showGridLines = True
    make_header_gar(
        ws_sum_g,
        "ИТОГОВАЯ ВЕДОМОСТЬ — ПРАКТИЧЕСКАЯ РАБОТА № 3",
        "Студент: Гарифуллин  |  Вариант / Задания № 8 – 13",
        "Сводная таблица результатов аналитического расчета надежности системы",
        max_c=7
    )
    
    ws_sum_g['B6'] = "Сводные числовые результаты расчета последовательных систем:"
    ws_sum_g['B6'].font = f_sec
    ws_sum_g.row_dimensions[6].height = 20
    
    sum_heads_gar = ["№ задания", "Объект расчета", "Показатель", "Формула", "Результат", "Ед. изм."]
    for i, h in enumerate(sum_heads_gar, start=2):
        cell = ws_sum_g.cell(7, i, h)
        cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
        cell.fill = fill_forest; cell.alignment = al_c; cell.border = box_cell
    ws_sum_g.row_dimensions[7].height = 22

    sum_rows_gar = [
        ("Задание 8", "Аппаратура связи (2000 элементов, t=200 ч)", "Pc(200)", "Pc = exp(-λc·t)", f"='Расчетная часть (№8-13)'!{results_map['t8_pc']}", "—", "0.0000"),
        ("Задание 8", "Аппаратура связи (2000 элементов, t=200 ч)", "mtc", "mtc = 1 / λc", f"='Расчетная часть (№8-13)'!{results_map['t8_mtc']}", "ч", "0.00"),
        ("Задание 9", "ЭВМ из 200000 элементов (сутки)", "Pc(24)", "Pc = exp(-λc·t)", f"='Расчетная часть (№8-13)'!{results_map['t9_pc']}", "—", "0.0000"),
        ("Задание 9", "ЭВМ из 200000 элементов (сутки)", "mtc", "mtc = 1 / λc", f"='Расчетная часть (№8-13)'!{results_map['t9_mtc']}", "ч", "0.00"),
        ("Задание 10", "Система управления (6000 элементов, t=50 ч)", "Pc(50)", "Pc = exp(-λc·t)", f"='Расчетная часть (№8-13)'!{results_map['t10_pc']}", "—", "0.0000"),
        ("Задание 10", "Система управления (6000 элементов, t=50 ч)", "mtc", "mtc = 1 / λc", f"='Расчетная часть (№8-13)'!{results_map['t10_mtc']}", "ч", "0.00"),
        ("Задание 11", "Прибор из 5 узлов", "Pc (точная)", "Pc = Π Pi", f"='Расчетная часть (№8-13)'!{results_map['t11_pc']}", "—", "0.0000"),
        ("Задание 11", "Прибор из 5 узлов", "Pc (прибл.)", "Pc ≈ 1 - Σ qi", f"='Расчетная часть (№8-13)'!{results_map['t11_pc_appr']}", "—", "0.0000"),
        ("Задание 12", "Система из 5 приборов (mti = {83, 220, 280, 400, 700})", "λc", "λc = Σ(1/mti)", f"='Расчетная часть (№8-13)'!{results_map['t12_lc']}", "1/ч", "0.000000"),
        ("Задание 12", "Система из 5 приборов (mti = {83, 220, 280, 400, 700})", "mtc", "mtc = 1 / λc", f"='Расчетная часть (№8-13)'!{results_map['t12_mtc']}", "ч", "0.00"),
        ("Задание 13", "5 блоков с Pi={0.98..0.985} за 50 ч", "mtc", "mtc = -t / ln(Pc)", f"='Расчетная часть (№8-13)'!{results_map['t13_mtc']}", "ч", "0.00")
    ]
    
    for idx, s_row in enumerate(sum_rows_gar, start=8):
        ws_sum_g.row_dimensions[idx].height = 19
        bg = fill_zebra if idx % 2 == 0 else PatternFill(fill_type=None)
        
        c1 = ws_sum_g.cell(idx, 2, s_row[0]); c1.font = f_bold; c1.alignment = al_c; c1.border = box_cell; c1.fill = bg
        c2 = ws_sum_g.cell(idx, 3, s_row[1]); c2.font = f_reg; c2.alignment = al_l; c2.border = box_cell; c2.fill = bg
        c3 = ws_sum_g.cell(idx, 4, s_row[2]); c3.font = f_bold; c3.alignment = al_c; c3.border = box_cell; c3.fill = bg
        c4 = ws_sum_g.cell(idx, 5, s_row[3]); c4.font = f_math; c4.alignment = al_l; c4.border = box_cell; c4.fill = bg
        c5 = ws_sum_g.cell(idx, 6, s_row[4]); c5.font = f_res; c5.fill = fill_res; c5.alignment = al_r; c5.border = box_res; c5.number_format = s_row[6]
        c6 = ws_sum_g.cell(idx, 7, s_row[5]); c6.font = f_reg; c6.alignment = al_c; c6.border = box_cell; c6.fill = bg

    col_w_sum_g = {"A": 3, "B": 14, "C": 40, "D": 14, "E": 24, "F": 16, "G": 12}
    for col, width in col_w_sum_g.items():
        ws_sum_g.column_dimensions[col].width = width

    fname = "Практика_3_Гарифуллин.xlsx"
    wb.save(fname)
    print(f"Created: {fname}")


if __name__ == '__main__':
    build_pr3_gainanov()
    build_pr3_garifullin()

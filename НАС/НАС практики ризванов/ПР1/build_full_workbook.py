import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def generate_full_practice_1():
    wb = openpyxl.Workbook()
    
    FONT_NAME = 'Segoe UI'
    
    # Fonts
    font_main_title = Font(name=FONT_NAME, size=15, bold=True, color='FFFFFF')
    font_subtitle = Font(name=FONT_NAME, size=11, bold=False, italic=True, color='FFFFFF')
    font_meta = Font(name=FONT_NAME, size=10, bold=False, color='E0E0E0')
    font_sec_header = Font(name=FONT_NAME, size=12, bold=True, color='1B365D')
    font_card_title = Font(name=FONT_NAME, size=11, bold=True, color='1B365D')
    font_tbl_header = Font(name=FONT_NAME, size=10, bold=True, color='FFFFFF')
    font_regular = Font(name=FONT_NAME, size=10, bold=False, color='262626')
    font_bold = Font(name=FONT_NAME, size=10, bold=True, color='262626')
    font_italic = Font(name=FONT_NAME, size=9, italic=True, color='595959')
    font_math = Font(name=FONT_NAME, size=10, italic=True, color='1F4E78')
    font_result = Font(name=FONT_NAME, size=11, bold=True, color='1E4620')
    font_status_ok = Font(name=FONT_NAME, size=10, bold=True, color='1E4620')
    
    # Fills
    fill_navy_header = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
    fill_sub_header = PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid')
    fill_tbl_header = PatternFill(start_color='335E99', end_color='335E99', fill_type='solid')
    fill_card_header = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    fill_condition = PatternFill(start_color='F2F4F8', end_color='F2F4F8', fill_type='solid')
    fill_zebra = PatternFill(start_color='F9FAFC', end_color='F9FAFC', fill_type='solid')
    fill_result = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
    fill_white = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    fill_accent_row = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    
    # Borders
    border_thin_gray = Side(style='thin', color='D9D9D9')
    border_double_bottom = Side(style='double', color='1B365D')
    border_green_med = Side(style='medium', color='385723')
    border_green_thin = Side(style='thin', color='70AD47')
    
    box_cell_border = Border(left=border_thin_gray, right=border_thin_gray, top=border_thin_gray, bottom=border_thin_gray)
    box_total_border = Border(left=border_thin_gray, right=border_thin_gray, top=border_thin_gray, bottom=border_double_bottom)
    box_result_border = Border(left=border_green_thin, right=border_green_thin, top=border_green_thin, bottom=border_green_thin)
    
    # Alignments
    align_center = Alignment(horizontal='center', vertical='center')
    align_center_wrap = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center')
    align_left_wrap = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_right = Alignment(horizontal='right', vertical='center')
    
    def apply_row_style(ws, row_idx, start_col, end_col, font=None, fill=None, border=None, alignment=None, number_format=None):
        for col_idx in range(start_col, end_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if font: cell.font = font
            if fill: cell.fill = fill
            if border: cell.border = border
            if alignment: cell.alignment = alignment
            if number_format: cell.number_format = number_format

    def make_title_banner(ws, title_text, subtitle_text, meta_text, max_col=7):
        ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=max_col)
        c2 = ws.cell(row=2, column=2, value=title_text)
        c2.font = font_main_title
        c2.alignment = align_center
        apply_row_style(ws, 2, 2, max_col, fill=fill_navy_header)
        ws.row_dimensions[2].height = 30

        ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=max_col)
        c3 = ws.cell(row=3, column=2, value=subtitle_text)
        c3.font = font_subtitle
        c3.alignment = align_center
        apply_row_style(ws, 3, 2, max_col, fill=fill_sub_header)
        ws.row_dimensions[3].height = 24

        ws.merge_cells(start_row=4, start_column=2, end_row=4, end_column=max_col)
        c4 = ws.cell(row=4, column=2, value=meta_text)
        c4.font = font_meta
        c4.alignment = align_center
        apply_row_style(ws, 4, 2, max_col, fill=fill_sub_header)
        ws.row_dimensions[4].height = 20

    # Dictionary to collect all target cell references for the summary sheet
    RESULTS_REF = {}

    # =========================================================================
    # CREATE SHEET: ЗАДАЧИ 6-15 (САМОСТОЯТЕЛЬНЫЕ)
    # =========================================================================
    ws_main = wb.active
    ws_main.title = "Задачи 6-15"
    ws_main.views.sheetView[0].showGridLines = True
    
    make_title_banner(
        ws_main,
        "ЗАДАЧИ ДЛЯ САМОСТОЯТЕЛЬНОГО РЕШЕНИЯ (№ 6 – 15)",
        "Практическая работа № 1. Определение количественных характеристик надежности по статистическим данным об отказах",
        "Все расчеты выполнены формулами Excel со ссылками на ячейки исходных данных",
        max_col=7
    )
    
    cur_r = 6
    
    calc_col_headers = [
        "Определяемый показатель", "Обозначение", "Математическая формула", 
        "Значение (формула Excel)", "Ед. изм.", "Пояснение к результату"
    ]
    
    def render_dynamic_task(ws, task_title, task_condition, inputs, calculations, final_note, start_r, task_key):
        r = start_r
        # 1. Header Card
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        c_title = ws.cell(row=r, column=2, value=task_title)
        c_title.font = Font(name=FONT_NAME, size=11, bold=True, color='FFFFFF')
        c_title.alignment = align_left
        apply_row_style(ws, r, 2, 7, fill=fill_navy_header)
        ws.row_dimensions[r].height = 24
        task_start_row = r
        
        # 2. Condition Box
        r += 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=7)
        c_cond = ws.cell(row=r, column=2, value=f"Условие: {task_condition}")
        c_cond.font = font_italic
        c_cond.alignment = align_left_wrap
        apply_row_style(ws, r, 2, 7, fill=fill_condition)
        apply_row_style(ws, r+1, 2, 7, fill=fill_condition)
        ws.row_dimensions[r].height = 18
        ws.row_dimensions[r+1].height = 18
        r += 1
        
        # 3. Inputs Header
        r += 1
        ws.cell(row=r, column=2, value="Таблица 1. Исходные данные задачи").font = font_card_title
        ws.row_dimensions[r].height = 20
        
        r += 1
        sub_headers_inp = ["Наименование параметра", "Обозначение", "Значение", "Ед. изм.", "Пояснение"]
        for ci, sh in enumerate(sub_headers_inp, start=2):
            c = ws.cell(row=r, column=ci, value=sh)
            c.font = font_tbl_header
            c.fill = fill_sub_header
            c.alignment = align_center
            c.border = box_cell_border
        ws.row_dimensions[r].height = 22
        
        inp_map = {}
        for inp_row in inputs:
            r += 1
            ws.row_dimensions[r].height = 20
            p_name, p_sym, p_val, p_unit, p_desc = inp_row
            
            # If p_val is a dynamic formula template, format it with current inp_map
            if isinstance(p_val, str) and '{' in p_val:
                p_val = p_val.format(**inp_map)
            
            c1 = ws.cell(row=r, column=2, value=p_name)
            c1.font = font_regular
            c1.alignment = align_left
            c1.border = box_cell_border
            
            c2 = ws.cell(row=r, column=3, value=p_sym)
            c2.font = font_bold
            c2.alignment = align_center
            c2.border = box_cell_border
            
            c3 = ws.cell(row=r, column=4, value=p_val)
            c3.font = font_bold
            c3.alignment = align_right
            c3.border = box_cell_border
            if isinstance(p_val, (int, float)):
                c3.number_format = '#,##0.####'
            
            c4 = ws.cell(row=r, column=5, value=p_unit)
            c4.font = font_regular
            c4.alignment = align_center
            c4.border = box_cell_border
            
            c5 = ws.cell(row=r, column=6, value=p_desc)
            c5.font = font_italic
            c5.alignment = align_left
            c5.border = box_cell_border
            
            # Map symbol to cell coordinate
            inp_map[p_sym] = f"D{r}"
            # Also clean symbol for formatting (e.g. without parens)
            clean_sym = p_sym.replace('(', '_').replace(')', '').replace(' ', '')
            inp_map[clean_sym] = f"D{r}"
        
        # 4. Calculations Header
        r += 2
        ws.cell(row=r, column=2, value="Таблица 2. Расчет показателей надежности").font = font_card_title
        ws.row_dimensions[r].height = 20
        
        r += 1
        for ci, ch in enumerate(calc_col_headers, start=2):
            c = ws.cell(row=r, column=ci, value=ch)
            c.font = font_tbl_header
            c.fill = fill_tbl_header
            c.alignment = align_center
            c.border = box_cell_border
        ws.row_dimensions[r].height = 24
        
        calc_map = {}
        for calc_row in calculations:
            r += 1
            ws.row_dimensions[r].height = 22
            c_name, c_sym, c_math, c_fmt_template, c_unit, c_note, c_num_fmt = calc_row
            
            # Resolve formula template
            resolved_formula = c_fmt_template.format(**inp_map, **calc_map)
            
            c1 = ws.cell(row=r, column=2, value=c_name)
            c1.font = font_regular
            c1.alignment = align_left
            c1.border = box_cell_border
            
            c2 = ws.cell(row=r, column=3, value=c_sym)
            c2.font = font_bold
            c2.alignment = align_center
            c2.border = box_cell_border
            
            c3 = ws.cell(row=r, column=4, value=c_math)
            c3.font = font_math
            c3.alignment = align_left
            c3.border = box_cell_border
            
            c4 = ws.cell(row=r, column=5, value=resolved_formula)
            c4.font = font_result
            c4.alignment = align_right
            c4.fill = fill_result
            c4.border = box_result_border
            c4.number_format = c_num_fmt
            
            # Map calculated symbol
            calc_map[c_sym] = f"E{r}"
            clean_csym = c_sym.replace('(', '_').replace(')', '').replace(' ', '')
            calc_map[clean_csym] = f"E{r}"
            RESULTS_REF[f"{task_key}_{clean_csym}"] = f"'{ws.title}'!E{r}"
            
            c5 = ws.cell(row=r, column=6, value=c_unit)
            c5.font = font_regular
            c5.alignment = align_center
            c5.border = box_cell_border
            
            c6 = ws.cell(row=r, column=7, value=c_note)
            c6.font = font_italic
            c6.alignment = align_left
            c6.border = box_cell_border
            
        # 5. Final Output / Answer Box
        r += 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        c_ans = ws.cell(row=r, column=2, value=f"ВЫВОД / ОТВЕТ: {final_note}")
        c_ans.font = font_bold
        c_ans.alignment = align_left
        apply_row_style(ws, r, 2, 7, fill=fill_accent_row, border=Border(top=border_green_med, bottom=border_green_med))
        ws.row_dimensions[r].height = 24
        
        RESULTS_REF[f"{task_key}_start"] = f"'{ws.title}'!A{task_start_row}"
        
        r += 3
        return r

    # --- ЗАДАЧА 6 ---
    cur_r = render_dynamic_task(
        ws_main,
        "ЗАДАЧА № 6. Расчет частоты отказов f(t) и интенсивности отказов λ(t) при t = 4000 ч",
        "На испытание поставлено 100 однотипных изделий. За 4000 часов отказало 50 изделий. За интервал времени 4000 – 4100 часов отказало еще 20 изделий. Требуется определить f(t), λ(t) при t = 4000 ч.",
        [
            ("Общее число изделий на испытании", "N", 100, "шт.", "Количество испытываемых изделий"),
            ("Время наработки к началу интервала", "t", 4000, "ч", "Момент времени t = 4000 ч"),
            ("Число отказавших изделий к моменту t", "n_отк", 50, "шт.", "Число отказов за первые 4000 ч"),
            ("Время наработки к концу интервала", "t_след", 4100, "ч", "Конец интервала t + Δt"),
            ("Длина интервала времени", "dt", "={t_след}-{t}", "ч", "Интервал Δt = 4100 - 4000 = 100 ч"),
            ("Число отказов в интервале (t, t+Δt)", "dn", 20, "шт.", "Отказы за интервал 4000-4100 ч")
        ],
        [
            ("Число работоспособных к моменту t", "nt", "N - n_отк(t)", "={N}-{n_отк}", "шт.", "n(4000) = 100 - 50 = 50 изделий", "0"),
            ("Частота отказов", "f_4000", "Δn(t) / (N · Δt)", "={dn}/({N}*{dt})", "1/ч", "f(4000) = 20 / (100 · 100) = 0.002 ч⁻¹", "0.00000"),
            ("Интенсивность отказов", "l_4000", "Δn(t) / (n(t) · Δt)", "={dn}/({nt}*{dt})", "1/ч", "λ(4000) = 20 / (50 · 100) = 0.004 ч⁻¹", "0.00000")
        ],
        "При t = 4000 ч частота отказов f(4000) = 0,002 1/ч (2,0·10⁻³ ч⁻¹), интенсивность отказов λ(4000) = 0,004 1/ч (4,0·10⁻³ ч⁻¹).",
        cur_r,
        "t6"
    )

    # --- ЗАДАЧА 7 ---
    cur_r = render_dynamic_task(
        ws_main,
        "ЗАДАЧА № 7. Расчет вероятности безотказной работы p(t) и вероятности отказа q(t)",
        "На испытание поставлено 100 однотипных изделий. За 4000 часов отказало 50 изделий. Требуется определить p(t) и q(t) при t = 4000 ч.",
        [
            ("Общее число изделий на испытании", "N", 100, "шт.", "Число изделий, поставленных на испытания"),
            ("Время наработки", "t", 4000, "ч", "Момент времени t = 4000 ч"),
            ("Число отказавших изделий к моменту t", "n_отк", 50, "шт.", "Число изделий, отказавших за 4000 ч")
        ],
        [
            ("Число не отказавших изделий", "nt", "N - n_отк(t)", "={N}-{n_отк}", "шт.", "n(4000) = 100 - 50 = 50 изделий", "0"),
            ("Вероятность безотказной работы", "p_4000", "n(t) / N", "={nt}/{N}", "—", "p(4000) = 50 / 100 = 0.50 (50%)", "0.0000"),
            ("Вероятность отказа изделия", "q_4000", "(N - n(t)) / N = 1 - p(t)", "=1-{p_4000}", "—", "q(4000) = 50 / 100 = 0.50 (50%)", "0.0000")
        ],
        "При наработке t = 4000 ч вероятность безотказной работы p(4000) = 0,50 (50%), вероятность отказа q(4000) = 0,50 (50%).",
        cur_r,
        "t7"
    )

    # --- ЗАДАЧА 8 ---
    cur_r = render_dynamic_task(
        ws_main,
        "ЗАДАЧА № 8. Расчет показателей надежности гироскопов: f(t) и λ(t) при t = 1000 ч",
        "В течение 1000 часов из 10 гироскопов отказало 2. За интервал времени 1000 – 1100 часов отказал еще один гироскоп. Требуется определить f(t), λ(t) при t = 1000 ч.",
        [
            ("Общее число гироскопов", "N", 10, "шт.", "Число гироскопов на испытании"),
            ("Время наработки", "t", 1000, "ч", "Момент времени t = 1000 ч"),
            ("Число отказавших к моменту t", "n_отк", 2, "шт.", "Отказы за первые 1000 ч"),
            ("Время наработки к концу интервала", "t_след", 1100, "ч", "Конец интервала"),
            ("Интервал времени", "dt", "={t_след}-{t}", "ч", "Δt = 1100 - 1000 = 100 ч"),
            ("Число отказов за интервал Δt", "dn", 1, "шт.", "Отказ за интервал 1000-1100 ч")
        ],
        [
            ("Число работоспособных гироскопов", "nt", "N - n_отк(t)", "={N}-{n_отк}", "шт.", "n(1000) = 10 - 2 = 8 гироскопов", "0"),
            ("Частота отказов гироскопов", "f_1000", "Δn(t) / (N · Δt)", "={dn}/({N}*{dt})", "1/ч", "f(1000) = 1 / (10 · 100) = 0.001 ч⁻¹", "0.00000"),
            ("Интенсивность отказов гироскопов", "l_1000", "Δn(t) / (n(t) · Δt)", "={dn}/({nt}*{dt})", "1/ч", "λ(1000) = 1 / (8 · 100) = 0.00125 ч⁻¹", "0.00000")
        ],
        "При t = 1000 ч: частота отказов f(1000) = 0,001 1/ч (1,0·10⁻³ ч⁻¹); интенсивность отказов λ(1000) = 0,00125 1/ч (1,25·10⁻³ ч⁻¹).",
        cur_r,
        "t8"
    )

    # --- ЗАДАЧА 9 ---
    cur_r = render_dynamic_task(
        ws_main,
        "ЗАДАЧА № 9. Расчет характеристик электронных ламп: p(t) и q(t) при t = 4000 ч",
        "На испытание поставлено 1000 однотипных электронных ламп. За первые 3000 часов отказало 80 ламп. За интервал времени 3000 – 4000 часов отказало еще 50 ламп. Требуется определить p(t) и q(t) при t = 4000 ч.",
        [
            ("Общее число ламп на испытании", "N", 1000, "шт.", "Партия испытуемых ламп"),
            ("Число отказавших за первые 3000 ч", "n_отк1", 80, "шт.", "Отказы за интервал 0 - 3000 ч"),
            ("Число отказавших в интервале 3000-4000 ч", "dn", 50, "шт.", "Дополнительные отказы"),
            ("Время наработки", "t", 4000, "ч", "Контрольный момент времени")
        ],
        [
            ("Суммарное число отказов к t=4000 ч", "n_tot_fail", "n_отк(3000) + Δn", "={n_отк1}+{dn}", "шт.", "Всего отказало к 4000 ч: 80 + 50 = 130", "0"),
            ("Число ламп, не отказавших к 4000 ч", "nt", "N - Σn_отк", "={N}-{n_tot_fail}", "шт.", "Работоспособно к 4000 ч: 1000 - 130 = 870", "0"),
            ("Вероятность безотказной работы", "p_4000", "n(4000) / N", "={nt}/{N}", "—", "p(4000) = 870 / 1000 = 0.87 (87%)", "0.0000"),
            ("Вероятность отказа ламп", "q_4000", "Σn_отк / N = 1 - p(4000)", "=1-{p_4000}", "—", "q(4000) = 130 / 1000 = 0.13 (13%)", "0.0000")
        ],
        "К моменту времени t = 4000 ч вероятность безотказной работы p(4000) = 0,87 (87%), вероятность отказа ламп q(4000) = 0,13 (13%).",
        cur_r,
        "t9"
    )

    # --- ЗАДАЧА 10 ---
    cur_r = render_dynamic_task(
        ws_main,
        "ЗАДАЧА № 10. Определение p(1300), p(1400), f(1300), λ(1300) для партии 1000 изделий",
        "На испытание поставлено 1000 изделий. За время t = 1300 ч. вышло из строя 288 изделий. За последующий интервал времени 1300 – 1400 часов вышло из строя еще 13 изделий. Необходимо вычислить p(t) при t = 1300 ч. и t = 1400 ч.; f(t), λ(t) при t = 1300 ч.",
        [
            ("Общее число изделий", "N", 1000, "шт.", "Число изделий на испытании"),
            ("Время наработки t1", "t1", 1300, "ч", "Момент первого замера"),
            ("Число отказов к моменту t1", "n_отк1", 288, "шт.", "Отказы за первые 1300 ч"),
            ("Время наработки t2", "t2", 1400, "ч", "Момент второго замера"),
            ("Интервал времени", "dt", "={t2}-{t1}", "ч", "Δt = 1400 - 1300 = 100 ч"),
            ("Число отказов в интервале (t1, t2)", "dn", 13, "шт.", "Отказы за интервал 1300-1400 ч")
        ],
        [
            ("Число работоспособных при t=1300 ч", "nt1", "N - n_отк(1300)", "={N}-{n_отк1}", "шт.", "n(1300) = 1000 - 288 = 712 шт.", "0"),
            ("Число работоспособных при t=1400 ч", "nt2", "n(1300) - Δn", "={nt1}-{dn}", "шт.", "n(1400) = 712 - 13 = 699 шт.", "0"),
            ("Вероятность безотказной работы p(1300)", "p_1300", "n(1300) / N", "={nt1}/{N}", "—", "p(1300) = 712 / 1000 = 0.712 (71.2%)", "0.0000"),
            ("Вероятность безотказной работы p(1400)", "p_1400", "n(1400) / N", "={nt2}/{N}", "—", "p(1400) = 699 / 1000 = 0.699 (69.9%)", "0.0000"),
            ("Частота отказов f(1300)", "f_1300", "Δn / (N · Δt)", "={dn}/({N}*{dt})", "1/ч", "f(1300) = 13 / (1000 · 100) = 0.00013 ч⁻¹", "0.000000"),
            ("Интенсивность отказов λ(1300)", "l_1300", "Δn / (n(1300) · Δt)", "={dn}/({nt1}*{dt})", "1/ч", "λ(1300) = 13 / (712 · 100) ≈ 0.00018258 ч⁻¹", "0.000000")
        ],
        "Результаты: p(1300) = 0,7120; p(1400) = 0,6990; f(1300) = 1,30·10⁻⁴ 1/ч (0,000130 ч⁻¹); λ(1300) = 1,83·10⁻⁴ 1/ч (0,000183 ч⁻¹).",
        cur_r,
        "t10"
    )

    # --- ЗАДАЧА 11 ---
    cur_r = render_dynamic_task(
        ws_main,
        "ЗАДАЧА № 11. Определение p(60), p(65), f(60), λ(60) для выборки из 45 изделий",
        "На испытание поставлено 45 изделий. За время t = 60 ч. вышло из строя 35 изделий. За последующий интервал времени 60 – 65 часов вышло из строя еще 3 изделия. Необходимо вычислить p(t) при t = 60 ч. и t = 65 ч.; f(t), λ(t) при t = 60 ч.",
        [
            ("Общее число изделий", "N", 45, "шт.", "Испытуемая группа изделий"),
            ("Время наработки t1", "t1", 60, "ч", "Начало интервала"),
            ("Число отказов к моменту t1", "n_отк1", 35, "шт.", "Отказы за первые 60 ч"),
            ("Время наработки t2", "t2", 65, "ч", "Конец интервала"),
            ("Интервал времени", "dt", "={t2}-{t1}", "ч", "Δt = 65 - 60 = 5 ч"),
            ("Число отказов в интервале (60-65 ч)", "dn", 3, "шт.", "Отказы за 5 ч")
        ],
        [
            ("Число работоспособных при t=60 ч", "nt1", "N - n_отк(60)", "={N}-{n_отк1}", "шт.", "n(60) = 45 - 35 = 10 шт.", "0"),
            ("Число работоспособных при t=65 ч", "nt2", "n(60) - Δn", "={nt1}-{dn}", "шт.", "n(65) = 10 - 3 = 7 шт.", "0"),
            ("Вероятность безотказной работы p(60)", "p_60", "n(60) / N", "={nt1}/{N}", "—", "p(60) = 10 / 45 ≈ 0.2222 (22.2%)", "0.0000"),
            ("Вероятность безотказной работы p(65)", "p_65", "n(65) / N", "={nt2}/{N}", "—", "p(65) = 7 / 45 ≈ 0.1556 (15.6%)", "0.0000"),
            ("Частота отказов f(60)", "f_60", "Δn / (N · Δt)", "={dn}/({N}*{dt})", "1/ч", "f(60) = 3 / (45 · 5) = 1/75 ≈ 0.013333 ч⁻¹", "0.000000"),
            ("Интенсивность отказов λ(60)", "l_60", "Δn / (n(60) · Δt)", "={dn}/({nt1}*{dt})", "1/ч", "λ(60) = 3 / (10 · 5) = 0.060000 ч⁻¹", "0.000000")
        ],
        "Результаты: p(60) = 0,2222; p(65) = 0,1556; f(60) = 0,013333 1/ч; λ(60) = 0,060000 1/ч (6,0·10⁻² ч⁻¹).",
        cur_r,
        "t11"
    )

    # --- ЗАДАЧА 12 (ИНТЕРВАЛЬНАЯ ТАБЛИЦА) ---
    r = cur_r
    task_start_row = r
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_main.cell(row=r, column=2, value="ЗАДАЧА № 12. Расчет статистической оценки среднего времени безотказной работы mt* по сгруппированным данным").font = Font(name=FONT_NAME, size=11, bold=True, color='FFFFFF')
    apply_row_style(ws_main, r, 2, 7, fill=fill_navy_header)
    ws_main.row_dimensions[r].height = 24
    
    r += 1
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=7)
    c_cond12 = ws_main.cell(row=r, column=2, value="Условие: В результате наблюдения за 45 образцами радиоэлектронного оборудования, которые прошли предварительную 80-часовую приработку, получены данные до первого отказа всех 45 образцов, сведенные в таблицу. Необходимо определить mt*.")
    c_cond12.font = font_italic
    c_cond12.alignment = align_left_wrap
    apply_row_style(ws_main, r, 2, 7, fill=fill_condition)
    apply_row_style(ws_main, r+1, 2, 7, fill=fill_condition)
    ws_main.row_dimensions[r].height = 18
    ws_main.row_dimensions[r+1].height = 18
    r += 1
    
    r += 1
    ws_main.cell(row=r, column=2, value="Таблица 1. Расчет характеристик по интервалам времени (Формула 1.6: mt* ≈ Σ(ni · tср.i) / N)").font = font_card_title
    ws_main.row_dimensions[r].height = 20
    
    r += 1
    t12_headers = ["№ интервала (i)", "Интервал времени Δti, ч", "Середина интервала tср.i, ч", "Число отказов ni, шт.", "Произведение (ni · tср.i), ч·шт", "Пояснение"]
    for ci, th in enumerate(t12_headers, start=2):
        cell = ws_main.cell(row=r, column=ci, value=th)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center_wrap
        cell.border = box_cell_border
    ws_main.row_dimensions[r].height = 25
    
    interval_data_12 = [
        (1, "0 - 10", 0, 10, 19),
        (2, "10 - 20", 10, 20, 13),
        (3, "20 - 30", 20, 30, 8),
        (4, "30 - 40", 30, 40, 3),
        (5, "40 - 50", 40, 50, 0),
        (6, "50 - 60", 50, 60, 1),
        (7, "60 - 70", 60, 70, 1)
    ]
    
    start_table_row = r + 1
    for row_data in interval_data_12:
        r += 1
        ws_main.row_dimensions[r].height = 20
        idx, int_label, t_left, t_right, n_val = row_data
        
        c1 = ws_main.cell(row=r, column=2, value=idx)
        c1.font = font_regular
        c1.alignment = align_center
        c1.border = box_cell_border
        
        c2 = ws_main.cell(row=r, column=3, value=int_label)
        c2.font = font_regular
        c2.alignment = align_center
        c2.border = box_cell_border
        
        mid_val = (t_left + t_right) / 2
        c3 = ws_main.cell(row=r, column=4, value=mid_val)
        c3.font = font_regular
        c3.alignment = align_right
        c3.border = box_cell_border
        c3.number_format = '0.0'
        
        c4 = ws_main.cell(row=r, column=5, value=n_val)
        c4.font = font_bold
        c4.alignment = align_right
        c4.border = box_cell_border
        c4.number_format = '#,##0'
        
        c5 = ws_main.cell(row=r, column=6, value=f"=D{r}*E{r}")
        c5.font = font_bold
        c5.alignment = align_right
        c5.border = box_cell_border
        c5.number_format = '#,##0.0'
        
        c6 = ws_main.cell(row=r, column=7, value=f"tср = ({t_left}+{t_right})/2 = {mid_val} ч")
        c6.font = font_italic
        c6.alignment = align_left
        c6.border = box_cell_border
    
    end_table_row = r
    
    # Total row
    r += 1
    ws_main.row_dimensions[r].height = 22
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    c_tot_label = ws_main.cell(row=r, column=2, value="ИТОГО (Сумма):")
    c_tot_label.font = font_bold
    c_tot_label.alignment = align_right
    apply_row_style(ws_main, r, 2, 4, fill=fill_card_header, border=box_total_border)
    
    c_sum_n = ws_main.cell(row=r, column=5, value=f"=SUM(E{start_table_row}:E{end_table_row})")
    c_sum_n.font = font_bold
    c_sum_n.alignment = align_right
    c_sum_n.fill = fill_card_header
    c_sum_n.border = box_total_border
    c_sum_n.number_format = '#,##0'
    sum_n_cell = f"E{r}"
    
    c_sum_prod = ws_main.cell(row=r, column=6, value=f"=SUM(F{start_table_row}:F{end_table_row})")
    c_sum_prod.font = font_bold
    c_sum_prod.alignment = align_right
    c_sum_prod.fill = fill_card_header
    c_sum_prod.border = box_total_border
    c_sum_prod.number_format = '#,##0.0'
    sum_prod_cell = f"F{r}"
    
    c_tot_note = ws_main.cell(row=r, column=7, value="N = 45 образцов; Σ(ni · tср.i) = 715 ч·шт")
    c_tot_note.font = font_italic
    c_tot_note.alignment = align_left
    c_tot_note.fill = fill_card_header
    c_tot_note.border = box_total_border
    
    r += 2
    ws_main.cell(row=r, column=2, value="Таблица 2. Итоговый расчет среднего времени безотказной работы").font = font_card_title
    ws_main.row_dimensions[r].height = 20
    
    r += 1
    for ci, ch in enumerate(calc_col_headers, start=2):
        c = ws_main.cell(row=r, column=ci, value=ch)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_center
        c.border = box_cell_border
    ws_main.row_dimensions[r].height = 24
    
    calc_rows_12 = [
        ("Суммарное время наработки выборки", "Σ(ni·tср.i)", f"={sum_prod_cell}", f"={sum_prod_cell}", "ч·шт", "Числитель формулы (1.6)", '#,##0.0'),
        ("Общее число образцов (объем выборки)", "N", f"={sum_n_cell}", f"={sum_n_cell}", "шт.", "Знаменатель формулы (1.6)", '#,##0'),
        ("Среднее время безотказной работы (после приработки)", "mt*", "Σ(ni·tср.i) / N", f"={sum_prod_cell}/{sum_n_cell}", "ч", "mt* = 715 / 45 ≈ 15.8889 ч", '0.00'),
        ("Полное среднее время с учетом приработки (80 ч)", "T_полн*", "t_прир + mt*", f"=80+{sum_prod_cell}/{sum_n_cell}", "ч", "T = 80 + 15.89 = 95.89 ч", '0.00')
    ]
    
    for row_calc in calc_rows_12:
        r += 1
        ws_main.row_dimensions[r].height = 22
        c_name, c_sym, c_math, c_formula, c_unit, c_note, c_fmt = row_calc
        
        c1 = ws_main.cell(row=r, column=2, value=c_name)
        c1.font = font_regular
        c1.alignment = align_left
        c1.border = box_cell_border
        
        c2 = ws_main.cell(row=r, column=3, value=c_sym)
        c2.font = font_bold
        c2.alignment = align_center
        c2.border = box_cell_border
        
        c3 = ws_main.cell(row=r, column=4, value=c_math)
        c3.font = font_math
        c3.alignment = align_left
        c3.border = box_cell_border
        
        c4 = ws_main.cell(row=r, column=5, value=c_formula)
        c4.font = font_result
        c4.alignment = align_right
        c4.fill = fill_result
        c4.border = box_result_border
        c4.number_format = c_fmt
        
        if c_sym == "mt*":
            RESULTS_REF["t12_mt"] = f"'Задачи 6-15'!E{r}"
        
        c5 = ws_main.cell(row=r, column=6, value=c_unit)
        c5.font = font_regular
        c5.alignment = align_center
        c5.border = box_cell_border
        
        c6 = ws_main.cell(row=r, column=7, value=c_note)
        c6.font = font_italic
        c6.alignment = align_left
        c6.border = box_cell_border
        
    r += 1
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    c_ans12 = ws_main.cell(row=r, column=2, value="ВЫВОД / ОТВЕТ: Статистическая оценка среднего времени безотказной работы образцов после приработки составляет mt* = 15,89 ч (суммарное с учетом 80 ч приработки — 95,89 ч).")
    c_ans12.font = font_bold
    c_ans12.alignment = align_left
    apply_row_style(ws_main, r, 2, 7, fill=fill_accent_row, border=Border(top=border_green_med, bottom=border_green_med))
    ws_main.row_dimensions[r].height = 24
    
    RESULTS_REF["t12_start"] = f"'Задачи 6-15'!A{task_start_row}"
    cur_r = r + 3

    # --- ЗАДАЧА 13 (ВЫБОРКА 8 ИЗДЕЛИЙ: СРЕДНЕЕ, ДИСПЕРСИЯ, СКО) ---
    r = cur_r
    task_start_row = r
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_main.cell(row=r, column=2, value="ЗАДАЧА № 13. Статистическая оценка среднего времени безотказной работы mt* и дисперсии Dt* по индивидуальным наработкам").font = Font(name=FONT_NAME, size=11, bold=True, color='FFFFFF')
    apply_row_style(ws_main, r, 2, 7, fill=fill_navy_header)
    ws_main.row_dimensions[r].height = 24
    
    r += 1
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=7)
    c_cond13 = ws_main.cell(row=r, column=2, value="Условие: На испытание поставлено 8 однотипных изделий. Получены следующие значения (ti – время безотказной работы i-го изделия): t1 = 560 ч; t2 = 700 ч; t3 = 800 ч; t4 = 650 ч; t5 = 580 ч; t6 = 760 ч; t7 = 920 ч; t8 = 850 ч. Определить статистическую оценку среднего времени безотказной работы изделия.")
    c_cond13.font = font_italic
    c_cond13.alignment = align_left_wrap
    apply_row_style(ws_main, r, 2, 7, fill=fill_condition)
    apply_row_style(ws_main, r+1, 2, 7, fill=fill_condition)
    ws_main.row_dimensions[r].height = 18
    ws_main.row_dimensions[r+1].height = 18
    r += 1
    
    r += 1
    ws_main.cell(row=r, column=2, value="Таблица 1. Индивидуальные наработки изделий и квадраты отклонений").font = font_card_title
    ws_main.row_dimensions[r].height = 20
    
    r += 1
    t13_headers = ["№ изделия (i)", "Обозначение", "Время безотказной работы ti, ч", "Отклонение (ti - mt*), ч", "Квадрат отклонения (ti - mt*)^2, ч^2", "Пояснение"]
    for ci, th in enumerate(t13_headers, start=2):
        cell = ws_main.cell(row=r, column=ci, value=th)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center_wrap
        cell.border = box_cell_border
    ws_main.row_dimensions[r].height = 25
    
    t13_vals = [560, 700, 800, 650, 580, 760, 920, 850]
    start_t13_r = r + 1
    
    # We will reference the mt* row which is at:
    # 8 data rows + 1 total row + 2 gap rows + 1 table title + 1 header + 2nd row in Table 2
    # Let's write the rows first:
    temp_rows_to_update = []
    for idx_i, val_i in enumerate(t13_vals, start=1):
        r += 1
        ws_main.row_dimensions[r].height = 20
        c1 = ws_main.cell(row=r, column=2, value=idx_i)
        c1.font = font_regular
        c1.alignment = align_center
        c1.border = box_cell_border
        
        c2 = ws_main.cell(row=r, column=3, value=f"t{idx_i}")
        c2.font = font_bold
        c2.alignment = align_center
        c2.border = box_cell_border
        
        c3 = ws_main.cell(row=r, column=4, value=val_i)
        c3.font = font_bold
        c3.alignment = align_right
        c3.border = box_cell_border
        c3.number_format = '#,##0.0'
        
        temp_rows_to_update.append(r)
        
        c6 = ws_main.cell(row=r, column=7, value=f"Изделие № {idx_i}")
        c6.font = font_italic
        c6.alignment = align_left
        c6.border = box_cell_border
        
    end_t13_r = r
    
    # Total row
    r += 1
    ws_main.row_dimensions[r].height = 22
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c_tot13 = ws_main.cell(row=r, column=2, value="ИТОГО (Сумма Σti):")
    c_tot13.font = font_bold
    c_tot13.alignment = align_right
    apply_row_style(ws_main, r, 2, 3, fill=fill_card_header, border=box_total_border)
    
    c_sum_t13 = ws_main.cell(row=r, column=4, value=f"=SUM(D{start_t13_r}:D{end_t13_r})")
    c_sum_t13.font = font_bold
    c_sum_t13.alignment = align_right
    c_sum_t13.fill = fill_card_header
    c_sum_t13.border = box_total_border
    c_sum_t13.number_format = '#,##0.0'
    
    c_sum_dev = ws_main.cell(row=r, column=5, value=f"=SUM(E{start_t13_r}:E{end_t13_r})")
    c_sum_dev.font = font_italic
    c_sum_dev.alignment = align_right
    c_sum_dev.fill = fill_card_header
    c_sum_dev.border = box_total_border
    c_sum_dev.number_format = '0.00'
    
    c_sum_sq = ws_main.cell(row=r, column=6, value=f"=SUM(F{start_t13_r}:F{end_t13_r})")
    c_sum_sq.font = font_bold
    c_sum_sq.alignment = align_right
    c_sum_sq.fill = fill_card_header
    c_sum_sq.border = box_total_border
    c_sum_sq.number_format = '#,##0.00'
    
    c_t13_note = ws_main.cell(row=r, column=7, value="Σti = 5820 ч; сумма отклонений = 0")
    c_t13_note.font = font_italic
    c_t13_note.alignment = align_left
    c_t13_note.fill = fill_card_header
    c_t13_note.border = box_total_border
    
    # Table 2: Calculation Results
    r += 2
    ws_main.cell(row=r, column=2, value="Таблица 2. Статистические показатели надежности").font = font_card_title
    ws_main.row_dimensions[r].height = 20
    
    r += 1
    for ci, ch in enumerate(calc_col_headers, start=2):
        c = ws_main.cell(row=r, column=ci, value=ch)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_center
        c.border = box_cell_border
    ws_main.row_dimensions[r].height = 24
    
    r_n = r + 1
    r_mean = r + 2
    r_var = r + 3
    r_sd = r + 4
    r_v = r + 5
    
    # Fill Table 1 formulas now that r_mean is known!
    mean_ref_abs = f"$E${r_mean}"
    for row_idx in temp_rows_to_update:
        c4 = ws_main.cell(row=row_idx, column=5, value=f"=D{row_idx}-{mean_ref_abs}")
        c4.font = font_regular
        c4.alignment = align_right
        c4.border = box_cell_border
        c4.number_format = '0.00'
        
        c5 = ws_main.cell(row=row_idx, column=6, value=f"=E{row_idx}^2")
        c5.font = font_regular
        c5.alignment = align_right
        c5.border = box_cell_border
        c5.number_format = '#,##0.00'
        
    calc_rows_13 = [
        ("Общее число испытанных изделий", "N", f"=COUNT(D{start_t13_r}:D{end_t13_r})", f"=COUNT(D{start_t13_r}:D{end_t13_r})", "шт.", "Объем выборки N = 8", '0'),
        ("Среднее время безотказной работы (Формула 1.5)", "mt*", f"Σti / N = AVERAGE(D{start_t13_r}:D{end_t13_r})", f"=AVERAGE(D{start_t13_r}:D{end_t13_r})", "ч", "mt* = 5820 / 8 = 727.50 ч", '0.00'),
        ("Дисперсия наработки (Формула 1.7)", "Dt*", f"Σ(ti-mt)^2 / (N-1) = VAR.S(...)", f"=VAR.S(D{start_t13_r}:D{end_t13_r})", "ч^2", "Выборочная дисперсия времени работы", '0.00'),
        ("Среднеквадратическое отклонение (СКО)", "σt", "SQRT(Dt*) = STDEV.S(...)", f"=STDEV.S(D{start_t13_r}:D{end_t13_r})", "ч", "Стандартное отклонение наработки", '0.00'),
        ("Коэффициент вариации", "V", "σt / mt*", f"=E{r_sd}/E{r_mean}", "—", "V = 127.70 / 727.50 ≈ 0.1755 (17.55%)", '0.0000')
    ]
    
    for row_calc in calc_rows_13:
        r += 1
        ws_main.row_dimensions[r].height = 22
        c_name, c_sym, c_math, c_formula, c_unit, c_note, c_fmt = row_calc
        
        c1 = ws_main.cell(row=r, column=2, value=c_name)
        c1.font = font_regular
        c1.alignment = align_left
        c1.border = box_cell_border
        
        c2 = ws_main.cell(row=r, column=3, value=c_sym)
        c2.font = font_bold
        c2.alignment = align_center
        c2.border = box_cell_border
        
        c3 = ws_main.cell(row=r, column=4, value=c_math)
        c3.font = font_math
        c3.alignment = align_left
        c3.border = box_cell_border
        
        c4 = ws_main.cell(row=r, column=5, value=c_formula)
        c4.font = font_result
        c4.alignment = align_right
        c4.fill = fill_result
        c4.border = box_result_border
        c4.number_format = c_fmt
        
        if c_sym == "mt*":
            RESULTS_REF["t13_mt"] = f"'Задачи 6-15'!E{r}"
        
        c5 = ws_main.cell(row=r, column=6, value=c_unit)
        c5.font = font_regular
        c5.alignment = align_center
        c5.border = box_cell_border
        
        c6 = ws_main.cell(row=r, column=7, value=c_note)
        c6.font = font_italic
        c6.alignment = align_left
        c6.border = box_cell_border
        
    r += 1
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    c_ans13 = ws_main.cell(row=r, column=2, value="ВЫВОД / ОТВЕТ: Статистическая оценка среднего времени безотказной работы mt* = 727,50 ч. Выборочная дисперсия Dt* = 16307,14 ч², СКО σt = 127,70 ч.")
    c_ans13.font = font_bold
    c_ans13.alignment = align_left
    apply_row_style(ws_main, r, 2, 7, fill=fill_accent_row, border=Border(top=border_green_med, bottom=border_green_med))
    ws_main.row_dimensions[r].height = 24
    
    RESULTS_REF["t13_start"] = f"'Задачи 6-15'!A{task_start_row}"
    cur_r = r + 3

    # --- ЗАДАЧА 14 (СРЕДНЕЕ ВРЕМЯ ВОССТАНОВЛЕНИЯ) ---
    r = cur_r
    task_start_row = r
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_main.cell(row=r, column=2, value="ЗАДАЧА № 14. Расчет среднего времени восстановления аппаратуры mt (m_в)").font = Font(name=FONT_NAME, size=11, bold=True, color='FFFFFF')
    apply_row_style(ws_main, r, 2, 7, fill=fill_navy_header)
    ws_main.row_dimensions[r].height = 24
    
    r += 1
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=7)
    c_cond14 = ws_main.cell(row=r, column=2, value="Условие: За наблюдаемый период эксплуатации в аппаратуре было зарегистрировано 6 отказов. Время восстановления составило: t1 = 15 мин.; t2 = 20 мин.; t3 = 10 мин.; t4 = 28 мин.; t5 = 22 мин.; t6 = 30 мин. Требуется определить среднее время восстановления аппаратуры mt.")
    c_cond14.font = font_italic
    c_cond14.alignment = align_left_wrap
    apply_row_style(ws_main, r, 2, 7, fill=fill_condition)
    apply_row_style(ws_main, r+1, 2, 7, fill=fill_condition)
    ws_main.row_dimensions[r].height = 18
    ws_main.row_dimensions[r+1].height = 18
    r += 1
    
    r += 1
    ws_main.cell(row=r, column=2, value="Таблица 1. Данные по времени восстановления при отказах").font = font_card_title
    ws_main.row_dimensions[r].height = 20
    
    r += 1
    t14_headers = ["№ отказа (i)", "Обозначение", "Время восстановления tв.i, мин", "Время восстановления, ч", "Пояснение"]
    for ci, th in enumerate(t14_headers, start=2):
        cell = ws_main.cell(row=r, column=ci, value=th)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center_wrap
        cell.border = box_cell_border
    ws_main.row_dimensions[r].height = 25
    
    t14_vals = [15, 20, 10, 28, 22, 30]
    start_t14_r = r + 1
    for idx_i, val_i in enumerate(t14_vals, start=1):
        r += 1
        ws_main.row_dimensions[r].height = 20
        c1 = ws_main.cell(row=r, column=2, value=idx_i)
        c1.font = font_regular
        c1.alignment = align_center
        c1.border = box_cell_border
        
        c2 = ws_main.cell(row=r, column=3, value=f"t{idx_i}")
        c2.font = font_bold
        c2.alignment = align_center
        c2.border = box_cell_border
        
        c3 = ws_main.cell(row=r, column=4, value=val_i)
        c3.font = font_bold
        c3.alignment = align_right
        c3.border = box_cell_border
        c3.number_format = '0'
        
        c4 = ws_main.cell(row=r, column=5, value=f"=D{r}/60")
        c4.font = font_regular
        c4.alignment = align_right
        c4.border = box_cell_border
        c4.number_format = '0.00'
        
        c5 = ws_main.cell(row=r, column=6, value=f"Отказ № {idx_i}")
        c5.font = font_italic
        c5.alignment = align_left
        c5.border = box_cell_border
        
    end_t14_r = r
    
    # Total row
    r += 1
    ws_main.row_dimensions[r].height = 22
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c_tot14 = ws_main.cell(row=r, column=2, value="ИТОГО (Суммарное время):")
    c_tot14.font = font_bold
    c_tot14.alignment = align_right
    apply_row_style(ws_main, r, 2, 3, fill=fill_card_header, border=box_total_border)
    
    c_sum_t14 = ws_main.cell(row=r, column=4, value=f"=SUM(D{start_t14_r}:D{end_t14_r})")
    c_sum_t14.font = font_bold
    c_sum_t14.alignment = align_right
    c_sum_t14.fill = fill_card_header
    c_sum_t14.border = box_total_border
    c_sum_t14.number_format = '0'
    
    c_sum_t14_h = ws_main.cell(row=r, column=5, value=f"=SUM(E{start_t14_r}:E{end_t14_r})")
    c_sum_t14_h.font = font_bold
    c_sum_t14_h.alignment = align_right
    c_sum_t14_h.fill = fill_card_header
    c_sum_t14_h.border = box_total_border
    c_sum_t14_h.number_format = '0.00'
    
    c_t14_note = ws_main.cell(row=r, column=6, value="Сумма времени восстановления: 125 мин (2.08 ч)")
    c_t14_note.font = font_italic
    c_t14_note.alignment = align_left
    c_t14_note.fill = fill_card_header
    c_t14_note.border = box_total_border
    
    # Table 2: Calculation Results
    r += 2
    ws_main.cell(row=r, column=2, value="Таблица 2. Расчет среднего времени восстановления").font = font_card_title
    ws_main.row_dimensions[r].height = 20
    
    r += 1
    for ci, ch in enumerate(calc_col_headers, start=2):
        c = ws_main.cell(row=r, column=ci, value=ch)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_center
        c.border = box_cell_border
    ws_main.row_dimensions[r].height = 24
    
    calc_rows_14 = [
        ("Число зафиксированных отказов", "N", f"=COUNT(D{start_t14_r}:D{end_t14_r})", f"=COUNT(D{start_t14_r}:D{end_t14_r})", "отказов", "Число ремонтов N = 6", '0'),
        ("Среднее время восстановления (минуты)", "mt_мин", f"Σtв.i / N = AVERAGE(D{start_t14_r}:D{end_t14_r})", f"=AVERAGE(D{start_t14_r}:D{end_t14_r})", "мин", "m_в = 125 / 6 = 20.83 мин", '0.00'),
        ("Среднее время восстановления (часы)", "mt_ч", "m_в(мин) / 60", f"=AVERAGE(E{start_t14_r}:E{end_t14_r})", "ч", "m_в = 20.83 / 60 ≈ 0.347 ч", '0.0000')
    ]
    
    for row_calc in calc_rows_14:
        r += 1
        ws_main.row_dimensions[r].height = 22
        c_name, c_sym, c_math, c_formula, c_unit, c_note, c_fmt = row_calc
        
        c1 = ws_main.cell(row=r, column=2, value=c_name)
        c1.font = font_regular
        c1.alignment = align_left
        c1.border = box_cell_border
        
        c2 = ws_main.cell(row=r, column=3, value=c_sym)
        c2.font = font_bold
        c2.alignment = align_center
        c2.border = box_cell_border
        
        c3 = ws_main.cell(row=r, column=4, value=c_math)
        c3.font = font_math
        c3.alignment = align_left
        c3.border = box_cell_border
        
        c4 = ws_main.cell(row=r, column=5, value=c_formula)
        c4.font = font_result
        c4.alignment = align_right
        c4.fill = fill_result
        c4.border = box_result_border
        c4.number_format = c_fmt
        
        if c_sym == "mt_мин":
            RESULTS_REF["t14_mt"] = f"'Задачи 6-15'!E{r}"
        
        c5 = ws_main.cell(row=r, column=6, value=c_unit)
        c5.font = font_regular
        c5.alignment = align_center
        c5.border = box_cell_border
        
        c6 = ws_main.cell(row=r, column=7, value=c_note)
        c6.font = font_italic
        c6.alignment = align_left
        c6.border = box_cell_border
        
    r += 1
    ws_main.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    c_ans14 = ws_main.cell(row=r, column=2, value="ВЫВОД / ОТВЕТ: Среднее время восстановления аппаратуры составляет mt (m_в) = 20,83 мин (или 0,347 ч).")
    c_ans14.font = font_bold
    c_ans14.alignment = align_left
    apply_row_style(ws_main, r, 2, 7, fill=fill_accent_row, border=Border(top=border_green_med, bottom=border_green_med))
    ws_main.row_dimensions[r].height = 24
    
    RESULTS_REF["t14_start"] = f"'Задачи 6-15'!A{task_start_row}"
    cur_r = r + 3

    # --- ЗАДАЧА 15 ---
    cur_r = render_dynamic_task(
        ws_main,
        "ЗАДАЧА № 15. Определение p(11000), p(12000), f(11000), λ(11000) для партии 1000 изделий",
        "На испытание поставлено 1000 изделий. За время t = 11000 ч. вышло из строя 410 изделий. За последующий интервал времени 11000 – 12000 часов вышло из строя еще 40 изделий. Необходимо вычислить p(t) при t = 11000 ч. и t = 12000 ч., а также f(t), λ(t) при t = 11000 ч.",
        [
            ("Общее число изделий", "N", 1000, "шт.", "Испытуемая партия изделий"),
            ("Время наработки t1", "t1", 11000, "ч", "Первая контрольная точка"),
            ("Число отказов к моменту t1", "n_отк1", 410, "шт.", "Отказы за 11000 ч"),
            ("Время наработки t2", "t2", 12000, "ч", "Вторая контрольная точка"),
            ("Интервал времени", "dt", "={t2}-{t1}", "ч", "Δt = 12000 - 11000 = 1000 ч"),
            ("Число отказов в интервале (t1, t2)", "dn", 40, "шт.", "Отказы за интервал 11000-12000 ч")
        ],
        [
            ("Число работоспособных при t=11000 ч", "nt1", "N - n_отк(11000)", "={N}-{n_отк1}", "шт.", "n(11000) = 1000 - 410 = 590 шт.", "0"),
            ("Число работоспособных при t=12000 ч", "nt2", "n(11000) - Δn", "={nt1}-{dn}", "шт.", "n(12000) = 590 - 40 = 550 шт.", "0"),
            ("Вероятность безотказной работы p(11000)", "p_11000", "n(11000) / N", "={nt1}/{N}", "—", "p(11000) = 590 / 1000 = 0.5900 (59.0%)", "0.0000"),
            ("Вероятность безотказной работы p(12000)", "p_12000", "n(12000) / N", "={nt2}/{N}", "—", "p(12000) = 550 / 1000 = 0.5500 (55.0%)", "0.0000"),
            ("Частота отказов f(11000)", "f_11000", "Δn / (N · Δt)", "={dn}/({N}*{dt})", "1/ч", "f(11000) = 40 / (1000 · 1000) = 0.000040 ч⁻¹", "0.000000"),
            ("Интенсивность отказов λ(11000)", "l_11000", "Δn / (n(11000) · Δt)", "={dn}/({nt1}*{dt})", "1/ч", "λ(11000) = 40 / (590 · 1000) ≈ 0.00006780 ч⁻¹", "0.000000")
        ],
        "Результаты: p(11000) = 0,5900; p(12000) = 0,5500; f(11000) = 4,00·10⁻⁵ 1/ч (0,000040 ч⁻¹); λ(11000) = 6,78·10⁻⁵ 1/ч (0,000068 ч⁻¹).",
        cur_r,
        "t15"
    )


    # =========================================================================
    # CREATE SHEET: ЗАДАЧИ 1-5 (ТИПОВЫЕ ПРИМЕРЫ ИЗ ПРЕЗЕНТАЦИИ)
    # =========================================================================
    ws_typ = wb.create_sheet(title="Задачи 1-5")
    ws_typ.views.sheetView[0].showGridLines = True
    
    make_title_banner(
        ws_typ,
        "ТИПОВЫЕ ЗАДАЧИ С РЕШЕНИЕМ ИЗ ПРЕЗЕНТАЦИИ (№ 1 – 5)",
        "Практическая работа № 1. Примеры расчета показателей надежности с формулами Excel",
        "Полное соответствие слайдам 2-4 презентации лекций",
        max_col=7
    )
    
    cur_r_typ = 6
    
    # --- ТИПОВАЯ ЗАДАЧА 1 ---
    cur_r_typ = render_dynamic_task(
        ws_typ,
        "ЗАДАЧА № 1. Определение среднего времени безотказной работы системы из трех блоков",
        "Система состоит из трех блоков, среднее время безотказной работы которых равно: mt1 = 160 ч.; mt2 = 320 ч.; mt3 = 600 ч. Для блоков справедлив экспоненциальный закон надежности. Требуется определить среднее время безотказной работы системы mtc.",
        [
            ("Среднее время безотказной работы блока 1", "mt1", 160, "ч", "Блок 1"),
            ("Среднее время безотказной работы блока 2", "mt2", 320, "ч", "Блок 2"),
            ("Среднее время безотказной работы блока 3", "mt3", 600, "ч", "Блок 3")
        ],
        [
            ("Интенсивность отказов блока 1", "l1", "1 / mt1", "=1/{mt1}", "1/ч", "λ1 = 1 / 160 = 0.00625 ч⁻¹", "0.000000"),
            ("Интенсивность отказов блока 2", "l2", "1 / mt2", "=1/{mt2}", "1/ч", "λ2 = 1 / 320 = 0.003125 ч⁻¹", "0.000000"),
            ("Интенсивность отказов блока 3", "l3", "1 / mt3", "=1/{mt3}", "1/ч", "λ3 = 1 / 600 ≈ 0.001667 ч⁻¹", "0.000000"),
            ("Интенсивность отказов системы", "lc", "λ1 + λ2 + λ3", "={l1}+{l2}+{l3}", "1/ч", "λc = 0.00625 + 0.003125 + 0.001667 ≈ 0.011042 ч⁻¹", "0.000000"),
            ("Среднее время безотказной работы системы", "mtc", "1 / λc", "=1/{lc}", "ч", "mtc = 1 / 0.011042 ≈ 90.57 ч (в слайде округлено до 91 ч)", "0.00")
        ],
        "Среднее время безотказной работы системы: mtc = 90,57 ч ≈ 91 ч (интенсивность отказов системы λc = 0,011 1/ч).",
        cur_r_typ,
        "typ1"
    )
    
    # --- ТИПОВАЯ ЗАДАЧА 2 ---
    cur_r_typ = render_dynamic_task(
        ws_typ,
        "ЗАДАЧА № 2. Определение p(t) и q(t) для 1000 подшипников качения при t = 3000 ч",
        "На испытание поставлено 1000 однотипных подшипников качения; за 3000 ч отказало 80 подшипников. Требуется определить p(t), q(t) при t = 3000 ч.",
        [
            ("Число подшипников на испытании", "N", 1000, "шт.", "Объем партии"),
            ("Время испытания", "t", 3000, "ч", "Момент времени 3000 ч"),
            ("Число отказавших подшипников", "n_отк", 80, "шт.", "Отказы за 3000 ч")
        ],
        [
            ("Число не отказавших подшипников", "nt", "N - n_отк(t)", "={N}-{n_отк}", "шт.", "n(3000) = 1000 - 80 = 920 шт.", "0"),
            ("Вероятность безотказной работы", "p_3000", "n(t) / N", "={nt}/{N}", "—", "p(3000) = 920 / 1000 = 0.92 (92%)", "0.0000"),
            ("Вероятность отказа подшипников", "q_3000", "(N - n(t)) / N = 1 - p", "=1-{p_3000}", "—", "q(3000) = 80 / 1000 = 0.08 (8%)", "0.0000")
        ],
        "При t = 3000 ч вероятность безотказной работы p(3000) = 0,92 (92%), вероятность отказа q(3000) = 0,08 (8%).",
        cur_r_typ,
        "typ2"
    )
    
    # --- ТИПОВАЯ ЗАДАЧА 3 ---
    cur_r_typ = render_dynamic_task(
        ws_typ,
        "ЗАДАЧА № 3. Определение p(3000), p(3100), f(3000), λ(3000) для 500 изделий",
        "На испытание поставлено 500 изделий. За время 3000 ч отказало 300 изделий, т.е. n(t) = 500 - 300 = 200. За интервал времени (t, t+Δt), где Δt = 100 ч, отказало еще 100 изделий, т.е. Δn(t) = 100. Требуется определить p(3000), p(3100), f(3000), λ(3000).",
        [
            ("Число изделий на испытании", "N", 500, "шт.", "Партия изделий"),
            ("Время первой наработки", "t", 3000, "ч", "Момент t = 3000 ч"),
            ("Число отказавших к моменту t", "n_отк1", 300, "шт.", "Отказы за первые 3000 ч"),
            ("Интервал времени", "dt", 100, "ч", "Δt = 100 ч (до 3100 ч)"),
            ("Число отказов в интервале Δt", "dn", 100, "шт.", "Отказы за 100 ч")
        ],
        [
            ("Число работоспособных к 3000 ч", "nt1", "N - n_отк(3000)", "={N}-{n_отк1}", "шт.", "n(3000) = 500 - 300 = 200 шт.", "0"),
            ("Число работоспособных к 3100 ч", "nt2", "n(3000) - Δn", "={nt1}-{dn}", "шт.", "n(3100) = 200 - 100 = 100 шт.", "0"),
            ("Вероятность безотказной работы p(3000)", "p_3000", "n(3000) / N", "={nt1}/{N}", "—", "p(3000) = 200 / 500 = 0.40 (40%)", "0.0000"),
            ("Вероятность безотказной работы p(3100)", "p_3100", "n(3100) / N", "={nt2}/{N}", "—", "p(3100) = 100 / 500 = 0.20 (20%)", "0.0000"),
            ("Частота отказов f(3000)", "f_3000", "Δn / (N · Δt)", "={dn}/({N}*{dt})", "1/ч", "f(3000) = 100 / (500 · 100) = 0.002 ч⁻¹", "0.000000"),
            ("Интенсивность отказов λ(3000)", "l_3000", "Δn / (n(3000) · Δt)", "={dn}/({nt1}*{dt})", "1/ч", "λ(3000) = 100 / (200 · 100) = 0.005 ч⁻¹", "0.000000")
        ],
        "Результаты: p(3000) = 0,40; p(3100) = 0,20; f(3000) = 0,002 1/ч; λ(3000) = 0,005 1/ч.",
        cur_r_typ,
        "typ3"
    )
    
    # --- ТИПОВАЯ ЗАДАЧА 4 ---
    r = cur_r_typ
    task_start_row = r
    ws_typ.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_typ.cell(row=r, column=2, value="ЗАДАЧА № 4. Статистическая оценка среднего времени безотказной работы mt* по выборке из 6 изделий").font = Font(name=FONT_NAME, size=11, bold=True, color='FFFFFF')
    apply_row_style(ws_typ, r, 2, 7, fill=fill_navy_header)
    ws_typ.row_dimensions[r].height = 24
    
    r += 1
    ws_typ.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=7)
    c_cond_typ4 = ws_typ.cell(row=r, column=2, value="Условие: На испытание поставлено шесть однотипных изделий. Получены следующие значения ti (время безотказной работы i-го изделия): t1 = 280 ч; t2 = 350 ч; t3 = 400 ч; t4 = 320 ч; t5 = 380 ч; t6 = 330 ч. Определить статистическую оценку среднего времени безотказной работы изделия.")
    c_cond_typ4.font = font_italic
    c_cond_typ4.alignment = align_left_wrap
    apply_row_style(ws_typ, r, 2, 7, fill=fill_condition)
    apply_row_style(ws_typ, r+1, 2, 7, fill=fill_condition)
    ws_typ.row_dimensions[r].height = 18
    ws_typ.row_dimensions[r+1].height = 18
    r += 1
    
    r += 1
    ws_typ.cell(row=r, column=2, value="Таблица 1. Значения наработок изделий до отказа").font = font_card_title
    ws_typ.row_dimensions[r].height = 20
    
    r += 1
    for ci, th in enumerate(["№ изделия (i)", "Обозначение", "Время работы ti, ч", "Пояснение"], start=2):
        c = ws_typ.cell(row=r, column=ci, value=th)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_center
        c.border = box_cell_border
    ws_typ.row_dimensions[r].height = 22
    
    t4_vals = [280, 350, 400, 320, 380, 330]
    st_t4 = r + 1
    for idx_i, val_i in enumerate(t4_vals, start=1):
        r += 1
        ws_typ.row_dimensions[r].height = 20
        c1 = ws_typ.cell(row=r, column=2, value=idx_i)
        c1.font = font_regular
        c1.alignment = align_center
        c1.border = box_cell_border
        
        c2 = ws_typ.cell(row=r, column=3, value=f"t{idx_i}")
        c2.font = font_bold
        c2.alignment = align_center
        c2.border = box_cell_border
        
        c3 = ws_typ.cell(row=r, column=4, value=val_i)
        c3.font = font_bold
        c3.alignment = align_right
        c3.border = box_cell_border
        c3.number_format = '#,##0.0'
        
        c4 = ws_typ.cell(row=r, column=5, value=f"Изделие № {idx_i}")
        c4.font = font_italic
        c4.alignment = align_left
        c4.border = box_cell_border
    end_t4 = r
    
    r += 1
    ws_typ.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws_typ.cell(row=r, column=2, value="ИТОГО (Сумма Σti):").font = font_bold
    ws_typ.cell(row=r, column=2).alignment = align_right
    apply_row_style(ws_typ, r, 2, 3, fill=fill_card_header, border=box_total_border)
    
    c_sum_t4 = ws_typ.cell(row=r, column=4, value=f"=SUM(D{st_t4}:D{end_t4})")
    c_sum_t4.font = font_bold
    c_sum_t4.alignment = align_right
    c_sum_t4.fill = fill_card_header
    c_sum_t4.border = box_total_border
    c_sum_t4.number_format = '#,##0.0'
    
    c_note_t4 = ws_typ.cell(row=r, column=5, value="Суммарная наработка 6 изделий = 2060 ч")
    c_note_t4.font = font_italic
    c_note_t4.alignment = align_left
    c_note_t4.fill = fill_card_header
    c_note_t4.border = box_total_border
    
    r += 2
    ws_typ.cell(row=r, column=2, value="Таблица 2. Расчет среднего времени mt* по формуле (1.5)").font = font_card_title
    ws_typ.row_dimensions[r].height = 20
    
    r += 1
    for ci, ch in enumerate(calc_col_headers, start=2):
        c = ws_typ.cell(row=r, column=ci, value=ch)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_center
        c.border = box_cell_border
    ws_typ.row_dimensions[r].height = 24
    
    calc_rows_t4 = [
        ("Число изделий на испытании", "N", f"=COUNT(D{st_t4}:D{end_t4})", f"=COUNT(D{st_t4}:D{end_t4})", "шт.", "Объем выборки N = 6", '0'),
        ("Суммарная наработка всех изделий", "Σti", f"=SUM(D{st_t4}:D{end_t4})", f"=SUM(D{st_t4}:D{end_t4})", "ч", "Σti = 2060 ч", '#,##0.0'),
        ("Среднее время безотказной работы", "mt*", f"Σti / N = AVERAGE(D{st_t4}:D{end_t4})", f"=AVERAGE(D{st_t4}:D{end_t4})", "ч", "mt* = 2060 / 6 ≈ 343.33 ч", '0.00')
    ]
    for row_calc in calc_rows_t4:
        r += 1
        ws_typ.row_dimensions[r].height = 22
        c_name, c_sym, c_math, c_formula, c_unit, c_note, c_fmt = row_calc
        
        c1 = ws_typ.cell(row=r, column=2, value=c_name)
        c1.font = font_regular
        c1.alignment = align_left
        c1.border = box_cell_border
        
        c2 = ws_typ.cell(row=r, column=3, value=c_sym)
        c2.font = font_bold
        c2.alignment = align_center
        c2.border = box_cell_border
        
        c3 = ws_typ.cell(row=r, column=4, value=c_math)
        c3.font = font_math
        c3.alignment = align_left
        c3.border = box_cell_border
        
        c4 = ws_typ.cell(row=r, column=5, value=c_formula)
        c4.font = font_result
        c4.alignment = align_right
        c4.fill = fill_result
        c4.border = box_result_border
        c4.number_format = c_fmt
        
        if c_sym == "mt*":
            RESULTS_REF["typ4_mt"] = f"'Задачи 1-5'!E{r}"
        
        c5 = ws_typ.cell(row=r, column=6, value=c_unit)
        c5.font = font_regular
        c5.alignment = align_center
        c5.border = box_cell_border
        
        c6 = ws_typ.cell(row=r, column=7, value=c_note)
        c6.font = font_italic
        c6.alignment = align_left
        c6.border = box_cell_border
        
    r += 1
    ws_typ.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    c_ans_t4 = ws_typ.cell(row=r, column=2, value="ВЫВОД / ОТВЕТ: Статистическая оценка среднего времени безотказной работы изделия: mt* = 343,33 ч.")
    c_ans_t4.font = font_bold
    c_ans_t4.alignment = align_left
    apply_row_style(ws_typ, r, 2, 7, fill=fill_accent_row, border=Border(top=border_green_med, bottom=border_green_med))
    ws_typ.row_dimensions[r].height = 24
    
    RESULTS_REF["typ4_start"] = f"'Задачи 1-5'!A{task_start_row}"
    cur_r_typ = r + 3
    
    # --- ТИПОВАЯ ЗАДАЧА 5 (ИНТЕРВАЛЬНЫЙ РЯД 16 ИНТЕРВАЛОВ) ---
    r = cur_r_typ
    task_start_row = r
    ws_typ.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws_typ.cell(row=r, column=2, value="ЗАДАЧА № 5. Расчет mt* по данным испытаний 45 образцов радиоэлектронного оборудования (16 интервалов)").font = Font(name=FONT_NAME, size=11, bold=True, color='FFFFFF')
    apply_row_style(ws_typ, r, 2, 7, fill=fill_navy_header)
    ws_typ.row_dimensions[r].height = 24
    
    r += 1
    ws_typ.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=7)
    c_cond_typ5 = ws_typ.cell(row=r, column=2, value="Условие: В результате наблюдения за 45 образцами радиоэлектронного оборудования получены данные до первого отказа всех 45 образцов, сведенные в таблицу. Требуется определить mt*.")
    c_cond_typ5.font = font_italic
    c_cond_typ5.alignment = align_left_wrap
    apply_row_style(ws_typ, r, 2, 7, fill=fill_condition)
    apply_row_style(ws_typ, r+1, 2, 7, fill=fill_condition)
    ws_typ.row_dimensions[r].height = 18
    ws_typ.row_dimensions[r+1].height = 18
    r += 1
    
    r += 1
    ws_typ.cell(row=r, column=2, value="Таблица 1. Интервальный вариационный ряд наработок до первого отказа").font = font_card_title
    ws_typ.row_dimensions[r].height = 20
    
    r += 1
    for ci, th in enumerate(t12_headers, start=2):
        cell = ws_typ.cell(row=r, column=ci, value=th)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center_wrap
        cell.border = box_cell_border
    ws_typ.row_dimensions[r].height = 25
    
    intervals_5 = [
        (1, "0 - 5", 0, 5, 1),
        (2, "5 - 10", 5, 10, 5),
        (3, "10 - 15", 10, 15, 8),
        (4, "15 - 20", 15, 20, 2),
        (5, "20 - 25", 20, 25, 5),
        (6, "25 - 30", 25, 30, 6),
        (7, "30 - 35", 30, 35, 4),
        (8, "35 - 40", 35, 40, 3),
        (9, "40 - 45", 40, 45, 0),
        (10, "45 - 50", 45, 50, 1),
        (11, "50 - 55", 50, 55, 0),
        (12, "55 - 60", 55, 60, 0),
        (13, "60 - 65", 60, 65, 3),
        (14, "65 - 70", 65, 70, 3),
        (15, "70 - 75", 70, 75, 3),
        (16, "75 - 80", 75, 80, 1)
    ]
    
    st_t5 = r + 1
    for row_data in intervals_5:
        r += 1
        ws_typ.row_dimensions[r].height = 20
        idx, int_label, t_left, t_right, n_val = row_data
        
        c1 = ws_typ.cell(row=r, column=2, value=idx)
        c1.font = font_regular
        c1.alignment = align_center
        c1.border = box_cell_border
        
        c2 = ws_typ.cell(row=r, column=3, value=int_label)
        c2.font = font_regular
        c2.alignment = align_center
        c2.border = box_cell_border
        
        mid_val = (t_left + t_right) / 2
        c3 = ws_typ.cell(row=r, column=4, value=mid_val)
        c3.font = font_regular
        c3.alignment = align_right
        c3.border = box_cell_border
        c3.number_format = '0.0'
        
        c4 = ws_typ.cell(row=r, column=5, value=n_val)
        c4.font = font_bold
        c4.alignment = align_right
        c4.border = box_cell_border
        c4.number_format = '#,##0'
        
        c5 = ws_typ.cell(row=r, column=6, value=f"=D{r}*E{r}")
        c5.font = font_bold
        c5.alignment = align_right
        c5.border = box_cell_border
        c5.number_format = '#,##0.0'
        
        c6 = ws_typ.cell(row=r, column=7, value=f"tср = ({t_left}+{t_right})/2 = {mid_val} ч")
        c6.font = font_italic
        c6.alignment = align_left
        c6.border = box_cell_border
    end_t5 = r
    
    # Total row
    r += 1
    ws_typ.row_dimensions[r].height = 22
    ws_typ.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws_typ.cell(row=r, column=2, value="ИТОГО (Сумма):").font = font_bold
    ws_typ.cell(row=r, column=2).alignment = align_right
    apply_row_style(ws_typ, r, 2, 4, fill=fill_card_header, border=box_total_border)
    
    c_sum_n5 = ws_typ.cell(row=r, column=5, value=f"=SUM(E{st_t5}:E{end_t5})")
    c_sum_n5.font = font_bold
    c_sum_n5.alignment = align_right
    c_sum_n5.fill = fill_card_header
    c_sum_n5.border = box_total_border
    c_sum_n5.number_format = '#,##0'
    sum_n5_ref = f"E{r}"
    
    c_sum_prod5 = ws_typ.cell(row=r, column=6, value=f"=SUM(F{st_t5}:F{end_t5})")
    c_sum_prod5.font = font_bold
    c_sum_prod5.alignment = align_right
    c_sum_prod5.fill = fill_card_header
    c_sum_prod5.border = box_total_border
    c_sum_prod5.number_format = '#,##0.0'
    sum_prod5_ref = f"F{r}"
    
    c_note_t5 = ws_typ.cell(row=r, column=7, value="N = 45; Σ(ni · tср.i) = 1427.5 ч·шт")
    c_note_t5.font = font_italic
    c_note_t5.alignment = align_left
    c_note_t5.fill = fill_card_header
    c_note_t5.border = box_total_border
    
    r += 2
    ws_typ.cell(row=r, column=2, value="Таблица 2. Расчет среднего времени безотказной работы по формуле (1.6)").font = font_card_title
    ws_typ.row_dimensions[r].height = 20
    
    r += 1
    for ci, ch in enumerate(calc_col_headers, start=2):
        c = ws_typ.cell(row=r, column=ci, value=ch)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_center
        c.border = box_cell_border
    ws_typ.row_dimensions[r].height = 24
    
    calc_rows_t5 = [
        ("Сумма произведений Σ(ni · tср.i)", "Σ(ni·tср.i)", f"={sum_prod5_ref}", f"={sum_prod5_ref}", "ч·шт", "Суммарная наработка выборки", '#,##0.0'),
        ("Объем выборки изделий", "N", f"={sum_n5_ref}", f"={sum_n5_ref}", "шт.", "Количество вышедших из строя изделий", '#,##0'),
        ("Среднее время безотказной работы", "mt*", f"Σ(ni·tср.i) / N", f"={sum_prod5_ref}/{sum_n5_ref}", "ч", "mt* = 1427.5 / 45 ≈ 31.72 ч (в слайде 31.7 ч)", '0.00')
    ]
    
    for row_calc in calc_rows_t5:
        r += 1
        ws_typ.row_dimensions[r].height = 22
        c_name, c_sym, c_math, c_formula, c_unit, c_note, c_fmt = row_calc
        
        c1 = ws_typ.cell(row=r, column=2, value=c_name)
        c1.font = font_regular
        c1.alignment = align_left
        c1.border = box_cell_border
        
        c2 = ws_typ.cell(row=r, column=3, value=c_sym)
        c2.font = font_bold
        c2.alignment = align_center
        c2.border = box_cell_border
        
        c3 = ws_typ.cell(row=r, column=4, value=c_math)
        c3.font = font_math
        c3.alignment = align_left
        c3.border = box_cell_border
        
        c4 = ws_typ.cell(row=r, column=5, value=c_formula)
        c4.font = font_result
        c4.alignment = align_right
        c4.fill = fill_result
        c4.border = box_result_border
        c4.number_format = c_fmt
        
        if c_sym == "mt*":
            RESULTS_REF["typ5_mt"] = f"'Задачи 1-5'!E{r}"
        
        c5 = ws_typ.cell(row=r, column=6, value=c_unit)
        c5.font = font_regular
        c5.alignment = align_center
        c5.border = box_cell_border
        
        c6 = ws_typ.cell(row=r, column=7, value=c_note)
        c6.font = font_italic
        c6.alignment = align_left
        c6.border = box_cell_border
        
    r += 1
    ws_typ.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    c_ans_t5 = ws_typ.cell(row=r, column=2, value="ВЫВОД / ОТВЕТ: Статистическая оценка среднего времени безотказной работы образцов: mt* = 31,72 ч ≈ 31,7 ч.")
    c_ans_t5.font = font_bold
    c_ans_t5.alignment = align_left
    apply_row_style(ws_typ, r, 2, 7, fill=fill_accent_row, border=Border(top=border_green_med, bottom=border_green_med))
    ws_typ.row_dimensions[r].height = 24
    
    RESULTS_REF["typ5_start"] = f"'Задачи 1-5'!A{task_start_row}"


    # =========================================================================
    # CREATE SHEET: СВОДКА И ОТВЕТЫ (INSERT AT INDEX 0)
    # =========================================================================
    ws_sum = wb.create_sheet(title="Сводка и Ответы", index=0)
    ws_sum.views.sheetView[0].showGridLines = True
    
    make_title_banner(
        ws_sum,
        "ПРАКТИЧЕСКАЯ РАБОТА № 1 — СВОДНЫЕ РЕЗУЛЬТАТЫ",
        "Определение количественных характеристик надежности по статистическим данным об отказах изделия",
        "Дисциплина: Надежность автоматизированных систем (НАС)  |  Преподаватель: Ризванов  |  Выполнил: Студент",
        max_col=8
    )
    
    ws_sum.cell(row=6, column=2, value="1. РЕЗУЛЬТАТЫ РЕШЕНИЯ ЗАДАЧ ДЛЯ САМОСТОЯТЕЛЬНОГО РЕШЕНИЯ (№ 6 – 15)").font = font_sec_header
    ws_sum.row_dimensions[6].height = 25
    
    sum_headers = [
        "№ задачи", "Краткое условие задачи", "Определяемые параметры", 
        "Расчетная формула (математическая)", "Результат вычислений", "Ед. изм.", "Ссылка на лист решения", "Статус"
    ]
    
    row = 7
    for c_idx, h in enumerate(sum_headers, start=2):
        cell = ws_sum.cell(row=row, column=c_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center_wrap
        cell.border = box_cell_border
    ws_sum.row_dimensions[row].height = 28
    
    summary_data = [
        ("Задача 6", "N=100, за 4000 ч отказало 50, за 4000-4100 ч отказало 20", "f(4000)", "f = Δn/(N·Δt)", f"={RESULTS_REF['t6_f_4000']}", "1/ч", RESULTS_REF['t6_start'], "Решено", "0.00000"),
        ("Задача 6 (доп.)", "Та же задача, интенсивность отказов λ(4000)", "λ(4000)", "λ = Δn/(n(t)·Δt)", f"={RESULTS_REF['t6_l_4000']}", "1/ч", RESULTS_REF['t6_start'], "Решено", "0.00000"),
        ("Задача 7", "N=100, за 4000 ч отказало 50. Найти p(4000), q(4000)", "p(4000)", "p = n(t)/N", f"={RESULTS_REF['t7_p_4000']}", "—", RESULTS_REF['t7_start'], "Решено", "0.0000"),
        ("Задача 7 (доп.)", "Та же задача, вероятность отказа q(4000)", "q(4000)", "q = (N - n(t))/N = 1 - p", f"={RESULTS_REF['t7_q_4000']}", "—", RESULTS_REF['t7_start'], "Решено", "0.0000"),
        ("Задача 8", "N=10 гироскопов, за 1000 ч отказало 2, за 1000-1100 ч еще 1", "f(1000)", "f = Δn/(N·Δt)", f"={RESULTS_REF['t8_f_1000']}", "1/ч", RESULTS_REF['t8_start'], "Решено", "0.00000"),
        ("Задача 8 (доп.)", "Та же задача, интенсивность отказов λ(1000)", "λ(1000)", "λ = Δn/(n(t)·Δt)", f"={RESULTS_REF['t8_l_1000']}", "1/ч", RESULTS_REF['t8_start'], "Решено", "0.00000"),
        ("Задача 9", "N=1000 ламп, за 3000 ч отказало 80, за 3000-4000 ч еще 50", "p(4000)", "p = (N - Σn_отк)/N", f"={RESULTS_REF['t9_p_4000']}", "—", RESULTS_REF['t9_start'], "Решено", "0.0000"),
        ("Задача 9 (доп.)", "Та же задача, вероятность отказа q(4000)", "q(4000)", "q = Σn_отк/N = 1 - p", f"={RESULTS_REF['t9_q_4000']}", "—", RESULTS_REF['t9_start'], "Решено", "0.0000"),
        ("Задача 10", "N=1000, t=1300: отк. 288; за 1300-1400: отк. 13", "p(1300)", "p(1300) = n(1300)/N", f"={RESULTS_REF['t10_p_1300']}", "—", RESULTS_REF['t10_start'], "Решено", "0.0000"),
        ("Задача 10 (2)", "Та же задача, вероятность безотказной работы p(1400)", "p(1400)", "p(1400) = (n(1300)-Δn)/N", f"={RESULTS_REF['t10_p_1400']}", "—", RESULTS_REF['t10_start'], "Решено", "0.0000"),
        ("Задача 10 (3)", "Та же задача, частота отказов f(1300)", "f(1300)", "f = Δn/(N·Δt)", f"={RESULTS_REF['t10_f_1300']}", "1/ч", RESULTS_REF['t10_start'], "Решено", "0.000000"),
        ("Задача 10 (4)", "Та же задача, интенсивность отказов λ(1300)", "λ(1300)", "λ = Δn/(n(1300)·Δt)", f"={RESULTS_REF['t10_l_1300']}", "1/ч", RESULTS_REF['t10_start'], "Решено", "0.000000"),
        ("Задача 11", "N=45, t=60: отк. 35; за 60-65: отк. 3", "p(60)", "p(60) = n(60)/N", f"={RESULTS_REF['t11_p_60']}", "—", RESULTS_REF['t11_start'], "Решено", "0.0000"),
        ("Задача 11 (2)", "Та же задача, вероятность безотказной работы p(65)", "p(65)", "p(65) = (n(60)-Δn)/N", f"={RESULTS_REF['t11_p_65']}", "—", RESULTS_REF['t11_start'], "Решено", "0.0000"),
        ("Задача 11 (3)", "Та же задача, частота отказов f(60)", "f(60)", "f = Δn/(N·Δt)", f"={RESULTS_REF['t11_f_60']}", "1/ч", RESULTS_REF['t11_start'], "Решено", "0.000000"),
        ("Задача 11 (4)", "Та же задача, интенсивность отказов λ(60)", "λ(60)", "λ = Δn/(n(60)·Δt)", f"={RESULTS_REF['t11_l_60']}", "1/ч", RESULTS_REF['t11_start'], "Решено", "0.000000"),
        ("Задача 12", "N=45 образцов РЭО (после 80 ч приработки), интервальные данные", "mt* (ср. наработка)", "mt* = Σ(ni·tср.i) / N", f"={RESULTS_REF['t12_mt']}", "ч", RESULTS_REF['t12_start'], "Решено", "0.00"),
        ("Задача 13", "N=8 изделий: ti = {560, 700, 800, 650, 580, 760, 920, 850}", "mt* (ср. наработка)", "mt* = Σti / N", f"={RESULTS_REF['t13_mt']}", "ч", RESULTS_REF['t13_start'], "Решено", "0.00"),
        ("Задача 14", "N=6 отказов аппаратуры: ti = {15, 20, 10, 28, 22, 30} мин", "m_в (ср. время восст.)", "m_в = Σtв,i / N", f"={RESULTS_REF['t14_mt']}", "мин", RESULTS_REF['t14_start'], "Решено", "0.00"),
        ("Задача 15", "N=1000, t=11000: отк. 410; за 11000-12000: отк. 40", "p(11000)", "p(11000) = n(11000)/N", f"={RESULTS_REF['t15_p_11000']}", "—", RESULTS_REF['t15_start'], "Решено", "0.0000"),
        ("Задача 15 (2)", "Та же задача, вероятность безотказной работы p(12000)", "p(12000)", "p(12000) = (n(11000)-Δn)/N", f"={RESULTS_REF['t15_p_12000']}", "—", RESULTS_REF['t15_start'], "Решено", "0.0000"),
        ("Задача 15 (3)", "Та же задача, частота отказов f(11000)", "f(11000)", "f = Δn/(N·Δt)", f"={RESULTS_REF['t15_f_11000']}", "1/ч", RESULTS_REF['t15_start'], "Решено", "0.000000"),
        ("Задача 15 (4)", "Та же задача, интенсивность отказов λ(11000)", "λ(11000)", "λ = Δn/(n(11000)·Δt)", f"={RESULTS_REF['t15_l_11000']}", "1/ч", RESULTS_REF['t15_start'], "Решено", "0.000000"),
    ]
    
    for item in summary_data:
        row += 1
        ws_sum.row_dimensions[row].height = 21
        is_even = (row % 2 == 0)
        curr_fill = fill_zebra if is_even else fill_white
        
        c_num = ws_sum.cell(row=row, column=2, value=item[0])
        c_num.font = font_bold
        c_num.alignment = align_center
        c_num.fill = curr_fill
        c_num.border = box_cell_border
        
        c_cond = ws_sum.cell(row=row, column=3, value=item[1])
        c_cond.font = font_regular
        c_cond.alignment = align_left
        c_cond.fill = curr_fill
        c_cond.border = box_cell_border
        
        c_param = ws_sum.cell(row=row, column=4, value=item[2])
        c_param.font = font_bold
        c_param.alignment = align_center
        c_param.fill = curr_fill
        c_param.border = box_cell_border
        
        c_form = ws_sum.cell(row=row, column=5, value=item[3])
        c_form.font = font_math
        c_form.alignment = align_left
        c_form.fill = curr_fill
        c_form.border = box_cell_border
        
        c_val = ws_sum.cell(row=row, column=6, value=item[4])
        c_val.font = font_result
        c_val.alignment = align_right
        c_val.fill = fill_result
        c_val.border = box_result_border
        c_val.number_format = item[8]
        
        c_unit = ws_sum.cell(row=row, column=7, value=item[5])
        c_unit.font = font_regular
        c_unit.alignment = align_center
        c_unit.fill = curr_fill
        c_unit.border = box_cell_border
        
        c_link = ws_sum.cell(row=row, column=8, value="Перейти к решению")
        c_link.hyperlink = f"#{item[6]}"
        c_link.font = Font(name=FONT_NAME, size=9, underline='single', color='1F4E78')
        c_link.alignment = align_center
        c_link.fill = curr_fill
        c_link.border = box_cell_border
        
        c_stat = ws_sum.cell(row=row, column=9, value=item[7])
        c_stat.font = font_status_ok
        c_stat.alignment = align_center
        c_stat.fill = curr_fill
        c_stat.border = box_cell_border

    # Table 2: Typical Examples Summary
    row += 3
    ws_sum.cell(row=row, column=2, value="2. РЕЗУЛЬТАТЫ РЕШЕНИЯ ТИПОВЫХ ЗАДАЧ ИЗ ПРЕЗЕНТАЦИИ (№ 1 – 5)").font = font_sec_header
    ws_sum.row_dimensions[row].height = 25
    
    row += 1
    typ_headers = [
        "№ задачи", "Тема / Объект расчета", "Ключевые исходные данные",
        "Искомый параметр", "Результат вычислений", "Ед. изм.", "Сверка со слайдом", "Ссылка"
    ]
    for c_idx, h in enumerate(typ_headers, start=2):
        cell = ws_sum.cell(row=row, column=c_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center_wrap
        cell.border = box_cell_border
    ws_sum.row_dimensions[row].height = 26
    
    typ_summary_data = [
        ("Задача 1", "Система из 3 блоков (экспоненц. закон)", "mt1=160, mt2=320, mt3=600 ч", "mtc (ср. время безотказной работы)", f"={RESULTS_REF['typ1_mtc']}", "ч", "90.57 ч ≈ 91 ч (Совпадает)", RESULTS_REF['typ1_start']),
        ("Задача 2", "Подшипники качения (N=1000)", "N=1000, t=3000 ч, отказов=80", "p(3000)", f"={RESULTS_REF['typ2_p_3000']}", "—", "p=0.92, q=0.08 (Совпадает)", RESULTS_REF['typ2_start']),
        ("Задача 3", "Изделия (N=500)", "N=500, t=3000 ч, отк.=300; Δt=100, Δn=100", "p(3000)", f"={RESULTS_REF['typ3_p_3000']}", "—", "p=0.4; 0.2; f=0.002; λ=0.005 (Совпадает)", RESULTS_REF['typ3_start']),
        ("Задача 4", "Индивидуальные наработки 6 изделий", "ti = {280, 350, 400, 320, 380, 330} ч", "mt* (выборочное среднее)", f"={RESULTS_REF['typ4_mt']}", "ч", "343.33 ч (Совпадает)", RESULTS_REF['typ4_start']),
        ("Задача 5", "Группированные данные 45 образцов РЭО", "16 интервалов по 5 ч, N=45", "mt* (по интервальному ряду)", f"={RESULTS_REF['typ5_mt']}", "ч", "31.72 ч ≈ 31.7 ч (Совпадает)", RESULTS_REF['typ5_start']),
    ]
    
    for item in typ_summary_data:
        row += 1
        ws_sum.row_dimensions[row].height = 21
        is_even = (row % 2 == 0)
        curr_fill = fill_zebra if is_even else fill_white
        
        c_num = ws_sum.cell(row=row, column=2, value=item[0])
        c_num.font = font_bold
        c_num.alignment = align_center
        c_num.fill = curr_fill
        c_num.border = box_cell_border
        
        c_th = ws_sum.cell(row=row, column=3, value=item[1])
        c_th.font = font_regular
        c_th.alignment = align_left
        c_th.fill = curr_fill
        c_th.border = box_cell_border
        
        c_in = ws_sum.cell(row=row, column=4, value=item[2])
        c_in.font = font_regular
        c_in.alignment = align_left
        c_in.fill = curr_fill
        c_in.border = box_cell_border
        
        c_p = ws_sum.cell(row=row, column=5, value=item[3])
        c_p.font = font_bold
        c_p.alignment = align_center
        c_p.fill = curr_fill
        c_p.border = box_cell_border
        
        c_v = ws_sum.cell(row=row, column=6, value=item[4])
        c_v.font = font_result
        c_v.alignment = align_right
        c_v.fill = fill_result
        c_v.border = box_result_border
        c_v.number_format = '0.00'
        
        c_u = ws_sum.cell(row=row, column=7, value=item[5])
        c_u.font = font_regular
        c_u.alignment = align_center
        c_u.fill = curr_fill
        c_u.border = box_cell_border
        
        c_chk = ws_sum.cell(row=row, column=8, value=item[6])
        c_chk.font = font_status_ok
        c_chk.alignment = align_center
        c_chk.fill = curr_fill
        c_chk.border = box_cell_border
        
        c_l = ws_sum.cell(row=row, column=9, value="Смотреть")
        c_l.hyperlink = f"#{item[7]}"
        c_l.font = Font(name=FONT_NAME, size=9, underline='single', color='1F4E78')
        c_l.alignment = align_center
        c_l.fill = curr_fill
        c_l.border = box_cell_border


    # =========================================================================
    # CREATE SHEET: ТЕОРИЯ И ФОРМУЛЫ
    # =========================================================================
    ws_th = wb.create_sheet(title="Теория и формулы")
    ws_th.views.sheetView[0].showGridLines = True
    
    make_title_banner(
        ws_th,
        "ТЕОРЕТИЧЕСКИЙ СПРАВОЧНИК — ХАРАКТЕРИСТИКИ НАДЕЖНОСТИ",
        "Основные формулы и определения статистической теории надежности (по ГОСТ 27.002 и материалам темы 1)",
        "Определения, обозначения, математические формулы и физический смысл",
        max_col=7
    )
    
    r = 6
    ws_th.cell(row=r, column=2, value="СВОДНАЯ ТАБЛИЦА СТАТИСТИЧЕСКИХ ОЦЕНОК НАДЕЖНОСТИ").font = font_sec_header
    ws_th.row_dimensions[r].height = 25
    
    r += 1
    th_headers = ["Показатель надежности", "Обозначение", "Математическая формула", "Номер формулы", "Ед. изм.", "Физический смысл и примечание"]
    for ci, th in enumerate(th_headers, start=2):
        cell = ws_th.cell(row=r, column=ci, value=th)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center_wrap
        cell.border = box_cell_border
    ws_th.row_dimensions[r].height = 26
    
    theory_rows = [
        ("Вероятность безотказной работы", "p(t)", "p(t) = n(t) / N", "(1.1)", "— (безразм.)", "Вероятность того, что в пределах заданной наработки t отказ объекта не возникнет."),
        ("Вероятность отказа", "q(t)", "q(t) = [N - n(t)] / N = 1 - p(t)", "(1.2)", "— (безразм.)", "Вероятность того, что изделие откажет хотя бы один раз к моменту времени t."),
        ("Частота отказов (плотность распределения)", "f(t)", "f(t) = Δn(t) / (N · Δt)", "(1.3)", "1/ч (ч⁻¹)", "Отношение числа отказавших изделий на участке (t, t+Δt) к произведению N на Δt."),
        ("Интенсивность отказов", "λ(t)", "λ(t) = Δn(t) / [n(t) · Δt]", "(1.4)", "1/ч (ч⁻¹)", "Условная плотность вероятности возникновения отказа для объекта, дошедшего работоспособным до момента t."),
        ("Среднее время безотказной работы (индивидуальное)", "mt*", "mt* = (1 / N) · Σ ti", "(1.5)", "ч, мин, с", "Математическое ожидание наработки изделия до первого отказа при известных моментах отказа всех N изделий."),
        ("Среднее время безотказной работы (группированное)", "mt*", "mt* ≈ (1 / N) · Σ (ni · tср.i)", "(1.6)", "ч, мин, с", "Оценка наработки по интервальному ряду, где tср.i = (ti-1 + ti)/2 — середина i-го интервала."),
        ("Выборочная дисперсия времени безотказной работы", "Dt*", "Dt* = [1 / (N - 1)] · Σ (ti - mt*)^2", "(1.7)", "ч², с²", "Характеризует рассеяние (разброс) времени безотказной работы изделий относительно среднего mt*."),
        ("Среднее квадратическое отклонение (СКО)", "σt", "σt = SQRT(Dt*)", "—", "ч, мин, с", "Абсолютная мера разброса наработок до отказа в тех же единицах измерения, что и наработка."),
        ("Среднее время восстановления", "m_в (mt)", "m_в = (1 / N) · Σ tв.i", "—", "мин, ч", "Математическое ожидание времени восстановления работоспособного состояния объекта после отказа.")
    ]
    
    for row_data in theory_rows:
        r += 1
        ws_th.row_dimensions[r].height = 26
        is_even = (r % 2 == 0)
        curr_fill = fill_zebra if is_even else fill_white
        
        c1 = ws_th.cell(row=r, column=2, value=row_data[0])
        c1.font = font_bold
        c1.alignment = align_left
        c1.fill = curr_fill
        c1.border = box_cell_border
        
        c2 = ws_th.cell(row=r, column=3, value=row_data[1])
        c2.font = font_bold
        c2.alignment = align_center
        c2.fill = curr_fill
        c2.border = box_cell_border
        
        c3 = ws_th.cell(row=r, column=4, value=row_data[2])
        c3.font = font_math
        c3.alignment = align_left
        c3.fill = curr_fill
        c3.border = box_cell_border
        
        c4 = ws_th.cell(row=r, column=5, value=row_data[3])
        c4.font = font_bold
        c4.alignment = align_center
        c4.fill = curr_fill
        c4.border = box_cell_border
        
        c5 = ws_th.cell(row=r, column=6, value=row_data[4])
        c5.font = font_regular
        c5.alignment = align_center
        c5.fill = curr_fill
        c5.border = box_cell_border
        
        c6 = ws_th.cell(row=r, column=7, value=row_data[5])
        c6.font = font_italic
        c6.alignment = align_left_wrap
        c6.fill = curr_fill
        c6.border = box_cell_border
        
    r += 2
    ws_th.cell(row=r, column=2, value="ОСНОВНЫЕ СООТНОШЕНИЯ МЕЖДУ ПОКАЗАТЕЛЯМИ НАДЕЖНОСТИ").font = font_sec_header
    ws_th.row_dimensions[r].height = 25
    
    rel_rows = [
        ("Связь вероятности безотказной работы и вероятности отказа:", "p(t) + q(t) = 1  <=>  p(t) = 1 - q(t)  <=>  q(t) = 1 - p(t)"),
        ("Связь интенсивности отказов и частоты отказов:", "λ(t) = f(t) / p(t)  <=>  f(t) = λ(t) · p(t)"),
        ("Для экспоненциального закона распределения (λ = const):", "p(t) = exp(-λ · t);   f(t) = λ · exp(-λ · t);   mt = 1 / λ"),
        ("Для последовательного соединения независимых блоков по надежности:", "λ_сист = Σ λ_i;   mt_сист = 1 / λ_сист = 1 / (Σ (1 / mt_i))")
    ]
    for rel_title, rel_expr in rel_rows:
        r += 1
        ws_th.row_dimensions[r].height = 24
        ws_th.cell(row=r, column=2, value=rel_title).font = font_bold
        ws_th.cell(row=r, column=2).alignment = align_left
        ws_th.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        c_rel = ws_th.cell(row=r, column=3, value=rel_expr)
        c_rel.font = font_math
        c_rel.alignment = align_left
        apply_row_style(ws_th, r, 2, 7, fill=fill_condition, border=box_cell_border)


    # =========================================================================
    # COLUMN WIDTHS
    # =========================================================================
    column_widths = {
        "Сводка и Ответы": {
            "A": 4, "B": 18, "C": 44, "D": 26, "E": 36, "F": 22, "G": 10, "H": 22, "I": 12
        },
        "Задачи 6-15": {
            "A": 4, "B": 40, "C": 18, "D": 34, "E": 24, "F": 12, "G": 46
        },
        "Задачи 1-5": {
            "A": 4, "B": 40, "C": 18, "D": 34, "E": 24, "F": 12, "G": 46
        },
        "Теория и формулы": {
            "A": 4, "B": 38, "C": 16, "D": 36, "E": 14, "F": 14, "G": 52
        }
    }
    
    for sheet_name, widths in column_widths.items():
        if sheet_name in wb.sheetnames:
            ws_curr = wb[sheet_name]
            for col_letter, width in widths.items():
                ws_curr.column_dimensions[col_letter].width = width

    # Save to disk
    output_filename = "Практическая_работа_1_НАС.xlsx"
    wb.save(output_filename)
    print(f"Successfully generated: {output_filename}")
    
    # Also save with alternative name Практика_1_Решение.xlsx so user has easy access
    wb.save("Практика_1_Решение.xlsx")
    print("Also saved as: Практика_1_Решение.xlsx")

if __name__ == '__main__':
    generate_full_practice_1()

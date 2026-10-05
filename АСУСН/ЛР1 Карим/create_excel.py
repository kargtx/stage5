"""
Скрипт создания файлов Excel для Лабораторной работы №1
Файл: lab1_modeling.xlsx
Обычное строгое форматирование таблиц без цветов
Лист 1: Дискретная логика (Таблица истинности)
Лист 2: Аналоговое масштабирование (Прямое и обратное масштабирование)
"""

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side


def create_excel_workbook():
    wb = openpyxl.Workbook()

    # Стили оформления: строгие, черно-белые, без цветовых заливок
    font_title = Font(name='Times New Roman', size=13, bold=True, color='000000')
    font_subtitle = Font(name='Times New Roman', size=11, italic=True, color='000000')
    font_sec_header = Font(name='Times New Roman', size=12, bold=True, color='000000')
    font_tbl_header = Font(name='Times New Roman', size=11, bold=True, color='000000')
    font_body = Font(name='Times New Roman', size=11, color='000000')

    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    # ==========================================
    # SHEET 1: ДИСКРЕТНАЯ ЛОГИКА
    # ==========================================
    ws1 = wb.active
    ws1.title = 'Дискретная логика'
    ws1.views.sheetView[0].showGridLines = True

    ws1['A1'] = 'Таблица истинности логических операций'
    ws1['A1'].font = font_title
    ws1['A2'] = 'Лабораторная работа № 1 | Студент: Гарифуллин К.Р. (гр. ЭАС-514С)'
    ws1['A2'].font = font_subtitle

    # Таблица истинности (из методички)
    headers_t1 = ['A', 'B', 'A И B', 'A ИЛИ B', 'НЕ A', 'A XOR B']
    for col_idx, h in enumerate(headers_t1, start=1):
        cell = ws1.cell(row=4, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.alignment = align_center
        cell.border = thin_border

    truth_inputs = [
        (0, 0),
        (0, 1),
        (1, 0),
        (1, 1)
    ]

    for idx, (a, b) in enumerate(truth_inputs, start=5):
        ws1.cell(row=idx, column=1, value=a).alignment = align_center
        ws1.cell(row=idx, column=2, value=b).alignment = align_center

        # Формулы Excel: =AND(A5,B5), =OR(A5,B5), =NOT(A5), =XOR(A5,B5)
        ws1.cell(row=idx, column=3, value=f'=IF(AND(A{idx}=1,B{idx}=1),1,0)').alignment = align_center
        ws1.cell(row=idx, column=4, value=f'=IF(OR(A{idx}=1,B{idx}=1),1,0)').alignment = align_center
        ws1.cell(row=idx, column=5, value=f'=IF(NOT(A{idx}=1),1,0)').alignment = align_center
        ws1.cell(row=idx, column=6, value=f'=IF(XOR(A{idx}=1,B{idx}=1),1,0)').alignment = align_center

        for c in range(1, 7):
            ws1.cell(row=idx, column=c).border = thin_border
            ws1.cell(row=idx, column=c).font = font_body

    # Пояснение используемых формул
    ws1['A11'] = 'Формулы Excel:'
    ws1['A11'].font = font_sec_header
    ws1['A12'] = '• И: =AND(A2;B2)'
    ws1['A12'].font = font_body
    ws1['A13'] = '• ИЛИ: =OR(A2;B2)'
    ws1['A13'].font = font_body
    ws1['A14'] = '• НЕ: =NOT(A2)'
    ws1['A14'].font = font_body
    ws1['A15'] = '• XOR: =XOR(A2;B2)'
    ws1['A15'].font = font_body

    ws1.column_dimensions['A'].width = 12
    ws1.column_dimensions['B'].width = 12
    ws1.column_dimensions['C'].width = 16
    ws1.column_dimensions['D'].width = 16
    ws1.column_dimensions['E'].width = 16
    ws1.column_dimensions['F'].width = 16

    # ==========================================
    # SHEET 2: АНАЛОГОВОЕ МАСШТАБИРОВАНИЕ
    # ==========================================
    ws2 = wb.create_sheet(title='Аналоговое масштабирование')
    ws2.views.sheetView[0].showGridLines = True

    ws2['A1'] = 'Масштабирование аналогового сигнала'
    ws2['A1'].font = font_title
    ws2['A2'] = 'Преобразование сырого значения АЦП (0–27648) в физическую величину (0–100%)'
    ws2['A2'].font = font_subtitle

    # Таблица масштабирования (в точности по методичке: Таблица 4)
    headers_a1 = ['A (сырое)', 'B (формула)', 'C (результат)']

    for col_idx, h in enumerate(headers_a1, start=1):
        cell = ws2.cell(row=4, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.alignment = align_center
        cell.border = thin_border

    analog_data = [
        (0, '=A5/27648*100'),
        (6912, '=A6/27648*100'),
        (13824, '=A7/27648*100'),
        (20736, '=A8/27648*100'),
        (27648, '=A9/27648*100')
    ]

    for idx, (raw, formula_str) in enumerate(analog_data, start=5):
        ws2.cell(row=idx, column=1, value=raw).alignment = align_center
        ws2.cell(row=idx, column=2, value=formula_str).alignment = align_left
        ws2.cell(row=idx, column=3, value=f'=A{idx}/27648*100').alignment = align_right

        ws2.cell(row=idx, column=3).number_format = '0.0'

        for c in range(1, 4):
            ws2.cell(row=idx, column=c).border = thin_border
            ws2.cell(row=idx, column=c).font = font_body

    # Таблица обратного масштабирования (ЦАП)
    ws2['A12'] = 'Обратное преобразование для аналогового выхода (ЦАП)'
    ws2['A12'].font = font_sec_header

    headers_a2 = ['Физическая величина (%)', 'Формула ЦАП', 'Сырое значение ЦАП']
    for col_idx, h in enumerate(headers_a2, start=1):
        cell = ws2.cell(row=13, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.alignment = align_center
        cell.border = thin_border

    dac_data = [0.0, 25.0, 50.0, 75.0, 100.0]
    for idx, pct in enumerate(dac_data, start=14):
        ws2.cell(row=idx, column=1, value=pct).alignment = align_right
        ws2.cell(row=idx, column=2, value=f'=ROUND(A{idx}/100*27648; 0)').alignment = align_left
        ws2.cell(row=idx, column=3, value=f'=ROUND(A{idx}/100*27648, 0)').alignment = align_center

        ws2.cell(row=idx, column=1).number_format = '0.0'
        ws2.cell(row=idx, column=3).number_format = '0'

        for c in range(1, 4):
            ws2.cell(row=idx, column=c).border = thin_border
            ws2.cell(row=idx, column=c).font = font_body

    ws2.column_dimensions['A'].width = 24
    ws2.column_dimensions['B'].width = 28
    ws2.column_dimensions['C'].width = 24

    wb.save('lab1_modeling.xlsx')
    print("Файл 'lab1_modeling.xlsx' успешно обновлен без цветов!")


if __name__ == '__main__':
    create_excel_workbook()

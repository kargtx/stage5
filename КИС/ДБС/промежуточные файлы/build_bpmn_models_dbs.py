# -*- coding: utf-8 -*-
"""
Генератор динамических BPMN-моделей процесса чип-тюнинга для ДБС в MS Visio:
1. Подпроцесс 1: Оформление заявки клиента и наряд-заказ (студент Гайнанов Д.И.)
2. Подпроцесс 2: Исполнение наряд-заказа, калибровка с комментарием мастера и выдача акта (студент Гарифуллин К.Р.)
3. Общая сквозная динамическая модель процесса чип-тюнинга (BPMN 2.0, 5 дорожек)

Стиль строго по требованиям:
- Блоки действий: ТОЛЬКО СИНИЙ цвет (RGB 239, 246, 255 / обводка RGB 37, 99, 235)
- Форма документов: значок документа со скошенным уголком ([Документ])
- Подписи стрелок: СТРОГО ПРОЗРАЧНЫЙ ФОН (FillPattern = 0)
- Год: 2026
- Исполнители: Гайнанов Д.И. и Гарифуллин К.Р.
"""

import win32com.client
import os

print("=== Генерация динамических моделей BPMN (Подпроцесс 1, Подпроцесс 2, Общая модель) ===")

out_dir = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС"

visio = win32com.client.Dispatch("Visio.Application")
visio.Visible = False
visio.AlertResponse = 7

def set_font(shape, size_pt=9, bold=False, color="RGB(15, 23, 42)"):
    shape.Cells("Char.Size").FormulaU = f"{size_pt} pt"
    shape.Cells("Char.Color").FormulaU = color
    if bold:
        shape.Cells("Char.Style").FormulaU = "1"

def draw_task(page, x, y, w, h, text, prefix=""):
    shape = page.DrawRectangle(x - w/2, y - h/2, x + w/2, y + h/2)
    shape.Cells("Rounding").FormulaU = "0.12 in"
    shape.Cells("LineWeight").FormulaU = "1.2 pt"
    # Строгое требование: СИНИЙ стиль для действий
    shape.Cells("FillForegnd").FormulaU = "RGB(239, 246, 255)"
    shape.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
    shape.Cells("LeftMargin").FormulaU = "0.05 in"
    shape.Cells("RightMargin").FormulaU = "0.05 in"
    shape.Cells("TopMargin").FormulaU = "0.05 in"
    shape.Cells("BottomMargin").FormulaU = "0.05 in"
    shape.Text = (prefix + "\n" if prefix else "") + text
    set_font(shape, 8, bold=False, color="RGB(15, 23, 42)")
    return shape

def draw_document(page, x, y, w, h, title, subtext=""):
    f = 0.22 # размер среза уголка
    x1, y1 = x - w/2, y - h/2
    x2, y2 = x + w/2, y + h/2
    pts = [
        x1, y1,
        x2, y1,
        x2, y2 - f,
        x2 - f, y2,
        x1, y2,
        x1, y1
    ]
    shape = page.DrawPolyline(pts, 0)
    shape.Cells("FillForegnd").FormulaU = "RGB(255, 255, 255)"
    shape.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
    shape.Cells("LineWeight").FormulaU = "1.5 pt"
    shape.Cells("LeftMargin").FormulaU = "0.04 in"
    shape.Cells("RightMargin").FormulaU = "0.04 in"
    shape.Cells("TopMargin").FormulaU = "0.04 in"
    shape.Cells("BottomMargin").FormulaU = "0.04 in"
    full_t = f"[Документ]\n{title}"
    if subtext:
        full_t += f"\n({subtext})"
    shape.Text = full_t
    set_font(shape, 8, bold=True, color="RGB(30, 58, 138)")

    # Треугольник загнутого уголка
    fold_pts = [
        x2 - f, y2,
        x2 - f, y2 - f,
        x2, y2 - f,
        x2 - f, y2
    ]
    fold = page.DrawPolyline(fold_pts, 0)
    fold.Cells("FillForegnd").FormulaU = "RGB(219, 234, 254)"
    fold.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
    fold.Cells("LineWeight").FormulaU = "1.2 pt"
    return shape

def draw_event(page, x, y, r, text, ev_type="Start"):
    shape = page.DrawOval(x - r, y - r, x + r, y + r)
    if ev_type == "Start":
        shape.Cells("FillForegnd").FormulaU = "RGB(220, 252, 231)"
        shape.Cells("LineColor").FormulaU = "RGB(22, 163, 74)"
        shape.Cells("LineWeight").FormulaU = "2 pt"
    elif ev_type == "EndSuccess":
        shape.Cells("FillForegnd").FormulaU = "RGB(220, 252, 231)"
        shape.Cells("LineColor").FormulaU = "RGB(22, 163, 74)"
        shape.Cells("LineWeight").FormulaU = "3.5 pt"
    elif ev_type == "EndFail":
        shape.Cells("FillForegnd").FormulaU = "RGB(254, 226, 226)"
        shape.Cells("LineColor").FormulaU = "RGB(220, 38, 38)"
        shape.Cells("LineWeight").FormulaU = "3.5 pt"

    lbl = page.DrawRectangle(x - 0.95, y - r - 0.45, x + 0.95, y - r)
    lbl.Text = text
    lbl.Cells("FillPattern").FormulaU = "0" # Прозрачный фон
    lbl.Cells("LinePattern").FormulaU = "0"
    set_font(lbl, 7.5, bold=True, color="RGB(30, 41, 59)")
    return shape

def draw_gateway(page, x, y, size, question="", q_pos="above"):
    xy = [x - size, y,  x, y + size,  x + size, y,  x, y - size,  x - size, y]
    shape = page.DrawPolyline(xy, 0)
    shape.Cells("FillForegnd").FormulaU = "RGB(254, 249, 195)"
    shape.Cells("LineColor").FormulaU = "RGB(202, 138, 4)"
    shape.Cells("LineWeight").FormulaU = "1.5 pt"
    shape.Text = "X"
    set_font(shape, 12, bold=True, color="RGB(161, 98, 7)")

    if question:
        if q_pos == "above":
            qlbl = page.DrawRectangle(x - 1.0, y + size + 0.08, x + 1.0, y + size + 0.48)
        else:
            qlbl = page.DrawRectangle(x - 1.0, y - size - 0.48, x + 1.0, y - size - 0.08)
        qlbl.Text = question
        qlbl.Cells("FillPattern").FormulaU = "0" # Прозрачный фон
        qlbl.Cells("LinePattern").FormulaU = "0"
        set_font(qlbl, 8, bold=True, color="RGB(161, 98, 7)")
    return shape

def connect(page, x1, y1, x2, y2, label="", dashed=False, label_offset_y=0.15, label_offset_x=0.0):
    line = page.DrawLine(x1, y1, x2, y2)
    line.Cells("EndArrow").FormulaU = "13"
    line.Cells("LineWeight").FormulaU = "1.2 pt"
    line.Cells("LineColor").FormulaU = "RGB(51, 65, 85)"
    if dashed:
        line.Cells("LinePattern").FormulaU = "2"
        line.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"

    if label:
        cx = (x1 + x2) / 2 + label_offset_x
        cy = (y1 + y2) / 2 + label_offset_y
        lbl = page.DrawRectangle(cx - 0.8, cy - 0.15, cx + 0.8, cy + 0.15)
        lbl.Text = label
        lbl.Cells("FillPattern").FormulaU = "0" # СТРОГО ПРОЗРАЧНЫЙ ФОН
        lbl.Cells("LinePattern").FormulaU = "0"
        set_font(lbl, 7.5, bold=True, color="RGB(30, 41, 59)")
    return line

def connect_elbow(page, points, label="", dashed=False, label_seg=0, label_offset_y=0.15, label_offset_x=0.0):
    for i in range(len(points) - 1):
        p1 = points[i]
        p2 = points[i+1]
        is_last = (i == len(points) - 2)
        line = page.DrawLine(p1[0], p1[1], p2[0], p2[1])
        line.Cells("LineWeight").FormulaU = "1.2 pt"
        line.Cells("LineColor").FormulaU = "RGB(51, 65, 85)"
        if dashed:
            line.Cells("LinePattern").FormulaU = "2"
            line.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
        if is_last:
            line.Cells("EndArrow").FormulaU = "13"
    if label:
        p1, p2 = points[label_seg], points[label_seg+1]
        cx = (p1[0] + p2[0])/2 + label_offset_x
        cy = (p1[1] + p2[1])/2 + label_offset_y
        lbl = page.DrawRectangle(cx - 0.8, cy - 0.15, cx + 0.8, cy + 0.15)
        lbl.Text = label
        lbl.Cells("FillPattern").FormulaU = "0" # СТРОГО ПРОЗРАЧНЫЙ ФОН
        lbl.Cells("LinePattern").FormulaU = "0"
        set_font(lbl, 7.5, bold=True, color="RGB(30, 41, 59)")

def setup_base_pool(page, w, h, pool_title, lanes):
    page.PageSheet.Cells("PageWidth").FormulaU = f"{w} in"
    page.PageSheet.Cells("PageHeight").FormulaU = f"{h} in"

    # Верхний заголовочный баннер
    tb = page.DrawRectangle(0.5, h - 1.1, w - 0.5, h - 0.5)
    tb.Text = pool_title
    tb.Cells("FillForegnd").FormulaU = "RGB(15, 23, 42)"
    tb.Cells("LineColor").FormulaU = "RGB(15, 23, 42)"
    set_font(tb, 12, bold=True, color="RGB(255,255,255)")

    # Левый вертикальный заголовок пула
    ph = page.DrawRectangle(0.4, 1.2, 1.5, h - 1.3)
    ph.Text = "ПУЛ: АВТОСЕРВИС (BPMN 2.0)"
    ph.Cells("FillForegnd").FormulaU = "RGB(30, 41, 59)"
    ph.Cells("LineColor").FormulaU = "RGB(15, 23, 42)"
    ph.Cells("LineWeight").FormulaU = "1.5 pt"
    ph.Cells("LeftMargin").FormulaU = "0.02 in"
    ph.Cells("RightMargin").FormulaU = "0.02 in"
    set_font(ph, 10, bold=True, color="RGB(255,255,255)")
    ph.Cells("TxtAngle").FormulaU = "90 deg"

    # Дорожки (Lanes)
    for r_title, yb, yt, bg_col in lanes:
        lane = page.DrawRectangle(1.5, yb, w - 0.5, yt)
        lane.Cells("FillForegnd").FormulaU = bg_col
        lane.Cells("LineColor").FormulaU = "RGB(203, 213, 225)"
        lane.Cells("LineWeight").FormulaU = "1 pt"

        rh = page.DrawRectangle(1.5, yb, 3.1, yt)
        rh.Text = r_title
        rh.Cells("FillForegnd").FormulaU = "RGB(241, 245, 249)"
        rh.Cells("LineColor").FormulaU = "RGB(148, 163, 184)"
        rh.Cells("LineWeight").FormulaU = "1.2 pt"
        set_font(rh, 8.5, bold=True, color="RGB(51, 65, 85)")

    # Нижняя легенда
    lg = page.DrawRectangle(3.1, 0.35, w - 0.5, 1.05)
    lg.Cells("FillForegnd").FormulaU = "RGB(248, 250, 252)"
    lg.Cells("LineColor").FormulaU = "RGB(203, 213, 225)"
    lg.Cells("LineWeight").FormulaU = "1 pt"
    lg.Text = "СТАНДАРТ BPMN 2.0:  (O) Начальное / Конечное событие  |  [X] Шлюз исключающего выбора XOR\n" \
              "[ ] Действие (Задача исполнителя / КИС) — строго синий стиль  |  [ Документ ] Объект данных BPMN  |  Подписи стрелок: прозрачный фон"
    set_font(lg, 8.5, bold=False, color="RGB(51, 65, 85)")

# ==============================================================================
# 1. ПОДПРОЦЕСС 1: ОФОРМЛЕНИЕ ЗАЯВКИ КЛИЕНТА И НАРЯД-ЗАКАЗ (ГАЙНАНОВ Д.И.)
# ==============================================================================
def create_subprocess_1_danil():
    print("Создание схемы: Подпроцесс 1 (Гайнанов Д.И.)...")
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)
    W, H = 20.0, 12.0
    TW, TH = 1.85, 0.85

    lanes = [
        ("Клиент\n(Автовладелец)", 8.5, 10.7, "RGB(255, 253, 245)"),
        ("Мастер-приемщик\n(Гайнанов Д.И.)", 5.0, 8.5, "RGB(242, 253, 245)"),
        ("КИС «Автосервис»\n(CRM, Склад, Каталог)", 1.2, 5.0, "RGB(240, 247, 255)")
    ]

    setup_base_pool(page, W, H, "ПОДПРОЦЕСС 1: ОФОРМЛЕНИЕ ЗАЯВКИ КЛИЕНТА И ВЫПУСК НАРЯД-ЗАКАЗА (СТУДЕНТ: ГАЙНАНОВ Д.И.)", lanes)

    ev_start = draw_event(page, 3.6, 9.6, 0.25, "Старт: Обращение\nклиента", "Start")
    t_c1 = draw_task(page, 5.0, 9.6, TW, TH, "1.1 Озвучивание пожеланий\nи предоставление авто\n(VIN, модель, задача)")
    connect(page, 3.85, 9.6, 4.075, 9.6)

    t_m1 = draw_task(page, 5.0, 6.75, TW, TH, "1.2 Первичный опрос,\nосмотр авто и ввод\nданных в КИС")
    connect(page, 5.0, 9.175, 5.0, 7.175)

    t_k1 = draw_task(page, 7.0, 3.1, TW, TH, "1.3 Поиск в базе КИС:\nистория авто, тип ЭБУ\nи проверка допуска")
    connect_elbow(page, [(5.925, 6.75), (7.0, 6.75), (7.0, 3.525)])

    gw_1 = draw_gateway(page, 8.8, 6.75, 0.3, "Тюнинг возможен?", q_pos="below")
    connect_elbow(page, [(7.925, 3.1), (8.8, 3.1), (8.8, 6.45)])

    t_fail1 = draw_task(page, 8.8, 9.6, TW, TH, "1.4 Информирование\nклиента о технических\nограничениях")
    connect(page, 8.8, 7.05, 8.8, 9.175, label="Нет [Отказ]", label_offset_x=0.45, label_offset_y=0.0)
    ev_fail = draw_event(page, 10.4, 9.6, 0.25, "Конец: Заявка отклонена\n(несовместимость)", "EndFail")
    connect(page, 9.725, 9.6, 10.15, 9.6)

    t_m2 = draw_task(page, 10.8, 6.75, TW, TH, "1.5 Консультация по пакетам\n(Stage 1, Евро-2, DPF)\nи смете работ")
    connect(page, 9.1, 6.75, 9.875, 6.75, label="Да [Норма]")

    t_k2 = draw_task(page, 10.8, 3.1, TW, TH, "1.6 Автоматический расчет\nсметы в КИС по прайсу\nи фиксация скидок")
    connect(page, 10.8, 6.325, 10.8, 3.525)

    t_c2 = draw_task(page, 12.8, 9.6, TW, TH, "1.7 Рассмотрение сметы\nи условий тюнинга\nклиентом")
    connect_elbow(page, [(11.725, 3.1), (12.8, 3.1), (12.8, 9.175)], label="Смета готова", label_seg=1, label_offset_x=0.55)

    gw_2 = draw_gateway(page, 14.6, 9.6, 0.3, "Клиент согласен?", q_pos="above")
    connect(page, 13.725, 9.6, 14.3, 9.6)

    t_fail2 = draw_task(page, 16.5, 9.6, TW, TH, "1.8 Фиксация причины\nотказа в CRM-системе")
    connect(page, 14.9, 9.6, 15.575, 9.6, label="Нет [Отказ]")
    ev_fail2 = draw_event(page, 18.1, 9.6, 0.25, "Конец: Отказ клиента\nот услуги", "EndFail")
    connect(page, 17.425, 9.6, 17.85, 9.6)

    t_k3 = draw_task(page, 14.6, 3.1, TW, TH, "1.9 Формирование наряда:\nприсвоение ID заказа,\nпривязка к VIN и ТЗ")
    connect(page, 14.6, 9.3, 14.6, 3.525, label="Да [Согласовано]", label_offset_x=0.65, label_offset_y=0.0)

    # ОБЪЕКТ ДАННЫХ: НАРИД-ЗАКАЗ (ДОКУМЕНТ ДАНИЛА)
    doc_order = draw_document(page, 16.8, 4.3, 1.8, 1.0, "Наряд-заказ\n№ КИС-2026", "Ключевой документ")
    connect(page, 15.525, 3.1, 15.9, 4.3, dashed=True)

    t_m3 = draw_task(page, 16.8, 6.75, TW, TH, "1.10 Подписание наряда\nклиентом и мастером,\nприемка ключей и ТС")
    connect(page, 16.8, 4.8, 16.8, 6.325, dashed=True)
    connect_elbow(page, [(15.525, 3.1), (15.7, 3.1), (15.7, 6.75), (15.875, 6.75)])

    ev_ok = draw_event(page, 18.7, 6.75, 0.25, "Финал: Наряд-заказ\nготов и передан\nв цех тюнинга", "EndSuccess")
    connect(page, 17.725, 6.75, 18.45, 6.75)

    p_vsd = os.path.join(out_dir, "ДБС_Подпроцесс_1_Оформление_заявки_Гайнанов.vsd")
    p_vsdx = os.path.join(out_dir, "ДБС_Подпроцесс_1_Оформление_заявки_Гайнанов.vsdx")
    p_png = os.path.join(out_dir, "ДБС_Подпроцесс_1_Оформление_заявки_Гайнанов.png")
    for p in [p_vsd, p_vsdx, p_png]:
        if os.path.exists(p): os.remove(p)
    doc.SaveAs(p_vsd)
    doc.SaveAs(p_vsdx)
    page.Export(p_png)
    doc.Saved = True
    doc.Close()
    print("Подпроцесс 1 сохранен:", p_png)

# ==============================================================================
# 2. ПОДПРОЦЕСС 2: ИСПОЛНЕНИЕ НАРИД-ЗАКАЗА И ВЫДАЧА АКТА (ГАРИФУЛЛИН К.Р.)
# ==============================================================================
def create_subprocess_2_karim():
    print("Создание схемы: Подпроцесс 2 (Гарифуллин К.Р.)...")
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)
    W, H = 21.5, 13.0
    TW, TH = 1.85, 0.85

    lanes = [
        ("Мастер-приемщик\n(Диспетчеризация и выдача)", 9.5, 11.7, "RGB(255, 251, 240)"),
        ("Инженер-диагност\n(Входной/выходной контроль)", 6.7, 9.5, "RGB(242, 253, 245)"),
        ("Инженер-калибровщик\n(Гарифуллин К.Р.)", 3.8, 6.7, "RGB(240, 247, 255)"),
        ("КИС «Автосервис: Тюнинг»\n(Модуль калибровки и БД)", 1.2, 3.8, "RGB(252, 246, 255)")
    ]

    setup_base_pool(page, W, H, "ПОДПРОЦЕСС 2: ИСПОЛНЕНИЕ НАРИД-ЗАКАЗА И ВЫДАЧА АКТА (СТУДЕНТ: ГАРИФУЛЛИН К.Р.)", lanes)

    ev_start = draw_event(page, 3.6, 8.1, 0.25, "Старт: Наряд-заказ\nпоступил в цех", "Start")
    doc_in = draw_document(page, 5.2, 10.6, 1.8, 1.0, "Наряд-заказ", "Входной документ")
    
    t_d1 = draw_task(page, 5.2, 8.1, TW, TH, "2.1 Входная диагностика ЭСУД:\nпроверка параметров,\nчтение кодов DTC")
    connect(page, 3.85, 8.1, 4.275, 8.1)
    connect(page, 5.2, 10.1, 5.2, 8.525, dashed=True)

    gw_d = draw_gateway(page, 7.2, 8.1, 0.3, "ДВС исправен?", q_pos="below")
    connect(page, 6.125, 8.1, 6.9, 8.1)

    t_fail = draw_task(page, 7.2, 10.6, TW, TH, "2.2 Оформление акта\nдефектовки в КИС и отказ\nот прошивки")
    connect(page, 7.2, 8.4, 7.2, 10.175, label="Нет [Ошибки]", label_offset_x=0.45, label_offset_y=0.0)
    ev_fail = draw_event(page, 8.9, 10.6, 0.25, "Конец: Снятие с поста\n(неисправность)", "EndFail")
    connect(page, 8.125, 10.6, 8.65, 10.6)

    t_t1 = draw_task(page, 9.0, 5.25, TW, TH, "2.3 Подключение питания\nАКБ (14V), считывание\nсток-дампа ROM")
    connect_elbow(page, [(7.5, 8.1), (9.0, 8.1), (9.0, 5.675)], label="Да [Допуск]", label_seg=0, label_offset_y=0.15)

    t_k1 = draw_task(page, 9.0, 2.5, TW, TH, "2.4 Автосохранение стока\nв архиве КИС\n(хэш SHA-256)")
    connect(page, 9.0, 4.825, 9.0, 2.925)

    t_t2 = draw_task(page, 11.2, 5.25, TW, TH, "2.5 Модификация карт\n(Stage 1 / Евро-2) в WinOLS,\nпересчет CRC и запись")
    connect_elbow(page, [(9.925, 2.5), (11.2, 2.5), (11.2, 4.825)])

    gw_flash = draw_gateway(page, 12.9, 5.25, 0.3, "Запись OK?", q_pos="below")
    connect(page, 12.125, 5.25, 12.6, 5.25)

    t_rec = draw_task(page, 12.9, 6.35, TW, 0.75, "2.6 Recovery Mode:\nАварийный откат на сток\n(режим Bench / Boot)")
    connect(page, 12.9, 5.55, 12.9, 5.975, label="Сбой", label_offset_x=0.3, label_offset_y=0.0)
    connect_elbow(page, [(12.9, 6.725), (12.9, 6.95), (11.2, 6.95), (11.2, 5.675)], label="Повтор", label_seg=1, label_offset_y=0.15)

    # ОБЯЗАТЕЛЬНО: ВЫПОЛНЕНИЕ + ТЕХНИЧЕСКИЙ КОММЕНТАРИЙ КАРИМА
    t_t3 = draw_task(page, 14.8, 5.25, TW, TH, "2.7 Сброс адаптаций ЭБУ,\nочистка сервисных DTC\n+ ТЕХ. КОММЕНТАРИЙ")
    connect(page, 13.2, 5.25, 13.875, 5.25, label="Успех")

    t_d2 = draw_task(page, 14.8, 8.1, TW, TH, "2.8 Выходной тест-драйв:\nснятие логов параметров\n(наддув, детонация, AFR)")
    connect(page, 14.8, 5.675, 14.8, 7.675)

    gw_log = draw_gateway(page, 16.7, 8.1, 0.3, "Логи в норме?", q_pos="above")
    connect(page, 15.725, 8.1, 16.4, 8.1)

    connect_elbow(page, [(16.7, 8.4), (16.7, 9.0), (11.2, 9.0), (11.2, 5.675)], label="Отклонение [Доработка]", label_seg=1, label_offset_y=0.15)

    t_k2 = draw_task(page, 16.7, 2.5, TW, TH, "2.9 Архивирование прошивки\nв КИС, расчет гарантии\nи генерация Акта")
    connect(page, 16.7, 7.8, 16.7, 2.925, label="Норма [Да]", label_offset_x=0.45, label_offset_y=0.0)

    # ОБЪЕКТ ДАННЫХ: АКТ ВЫПОЛНЕННЫХ РАБОТ (ДОКУМЕНТ КАРИМА)
    doc_act = draw_document(page, 19.0, 4.3, 1.9, 1.05, "Акт выполненных\nработ + Отчет", "Выходной документ")
    connect(page, 17.625, 2.5, 19.0, 3.775, dashed=True)

    t_m_final = draw_task(page, 19.0, 10.6, TW, TH, "2.10 Подписание акта,\nвыдача авто и гарантийного\nсертификата клиенту")
    connect(page, 19.0, 4.825, 19.0, 10.175, dashed=True)
    connect_elbow(page, [(17.625, 2.5), (17.8, 2.5), (17.8, 10.6), (18.075, 10.6)])

    ev_ok = draw_event(page, 20.7, 10.6, 0.25, "Финал: Акт подписан,\nавтомобиль сдан\nклиенту", "EndSuccess")
    connect(page, 19.925, 10.6, 20.45, 10.6)

    p_vsd = os.path.join(out_dir, "ДБС_Подпроцесс_2_Исполнение_заказа_Гарифуллин.vsd")
    p_vsdx = os.path.join(out_dir, "ДБС_Подпроцесс_2_Исполнение_заказа_Гарифуллин.vsdx")
    p_png = os.path.join(out_dir, "ДБС_Подпроцесс_2_Исполнение_заказа_Гарифуллин.png")
    for p in [p_vsd, p_vsdx, p_png]:
        if os.path.exists(p): os.remove(p)
    doc.SaveAs(p_vsd)
    doc.SaveAs(p_vsdx)
    page.Export(p_png)
    doc.Saved = True
    doc.Close()
    print("Подпроцесс 2 сохранен:", p_png)

# ==============================================================================
# 3. ОБЩАЯ СКВОЗНАЯ ДИНАМИЧЕСКАЯ МОДЕЛЬ В ТОМ ЖЕ BPMN-СТИЛЕ (5 ДОРОЖЕК)
# ==============================================================================
def create_unified_bpmn_overview():
    print("Создание схемы: Общая сквозная динамическая модель BPMN (единый стиль)...")
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)
    W, H = 25.0, 14.5
    TW, TH = 1.85, 0.85

    lanes = [
        ("Клиент\n(Автовладелец)", 11.2, 13.2, "RGB(255, 253, 245)"),
        ("Мастер-приемщик\n(Гайнанов Д.И.)", 8.7, 11.2, "RGB(242, 253, 245)"),
        ("Инженер-диагност\n(Входной контроль)", 6.2, 8.7, "RGB(255, 250, 240)"),
        ("Инженер-калибровщик\n(Гарифуллин К.Р.)", 3.7, 6.2, "RGB(240, 247, 255)"),
        ("КИС «Автосервис»\n(База данных и АРМ)", 1.2, 3.7, "RGB(252, 246, 255)")
    ]

    setup_base_pool(page, W, H, "ОБЩАЯ СКВОЗНАЯ ДИНАМИЧЕСКАЯ МОДЕЛЬ ПРОЦЕССА ЧИП-ТЮНИНГА В КИС (BPMN 2.0)", lanes)

    # 1. Start Event (Client lane Y = 12.2)
    ev_start = draw_event(page, 3.6, 12.2, 0.25, "Старт: Обращение\nклиента", "Start")
    t1 = draw_task(page, 5.0, 12.2, TW, TH, "1. Предоставление ТС,\nозвучивание задач\nпо чип-тюнингу")
    connect(page, 3.85, 12.2, 4.075, 12.2)

    # 2. Master приемка (Y = 9.95)
    t2 = draw_task(page, 5.0, 9.95, TW, TH, "2. Осмотр авто,\nпервичный опрос и ввод\nреквизитов в КИС")
    connect(page, 5.0, 11.775, 5.0, 10.375)

    # 3. КИС: Поиск и допуск (Y = 2.45)
    t3 = draw_task(page, 6.9, 2.45, TW, TH, "3. Проверка истории,\nопределение ЭБУ и подбор\nпакета модификации")
    connect_elbow(page, [(5.925, 9.95), (6.9, 9.95), (6.9, 2.875)])

    # 4. Шлюз: Тюнинг возможен? (Master Y = 9.95)
    gw1 = draw_gateway(page, 8.6, 9.95, 0.3, "Тюнинг возможен?", q_pos="below")
    connect_elbow(page, [(7.825, 2.45), (8.6, 2.45), (8.6, 9.65)])

    # Отказ по несовместимости
    t_fail1 = draw_task(page, 8.6, 12.2, TW, TH, "4. Отказ клиенту по\nтехническим причинам\n(несовместимость)")
    connect(page, 8.6, 10.25, 8.6, 11.775, label="Нет [Отказ]", label_offset_x=0.45, label_offset_y=0.0)
    ev_fail1 = draw_event(page, 10.1, 12.2, 0.25, "Конец: Отказ в услуге\n(несовместимость)", "EndFail")
    connect(page, 9.525, 12.2, 9.85, 12.2)

    # Согласование сметы
    t4 = draw_task(page, 10.6, 9.95, TW, TH, "5. Консультация по пакетам\nStage 1/2 и согласование\nпредварительной сметы")
    connect(page, 8.9, 9.95, 9.675, 9.95, label="Да [Допуск]")

    # Клиент утверждает
    t5 = draw_task(page, 12.4, 12.2, TW, TH, "6. Рассмотрение сметы\nи условий гарантии\nклиентом")
    connect_elbow(page, [(11.525, 9.95), (12.4, 9.95), (12.4, 11.775)])

    gw2 = draw_gateway(page, 13.9, 12.2, 0.3, "Клиент согласен?", q_pos="above")
    connect(page, 13.325, 12.2, 13.6, 12.2)

    # Отказ клиента
    ev_fail2 = draw_event(page, 15.3, 12.2, 0.25, "Конец: Отказ клиента\nот условий", "EndFail")
    connect(page, 14.2, 12.2, 15.05, 12.2, label="Нет")

    # Формирование наряда в КИС (Y = 2.45)
    t6 = draw_task(page, 13.9, 2.45, TW, TH, "7. Автоформирование\nНаряд-заказа в КИС,\nфиксация нормо-часов")
    connect(page, 13.9, 11.9, 13.9, 2.875, label="Да [Согласовано]", label_offset_x=0.65, label_offset_y=0.0)

    # ДОКУМЕНТ 1: НАРИД-ЗАКАЗ (ПОДПРОЦЕСС ДАНИЛА)
    doc_order = draw_document(page, 15.6, 4.8, 1.8, 1.0, "Наряд-заказ\n№ КИС-2026", "Документ Данила")
    connect(page, 14.825, 2.45, 15.6, 4.3, dashed=True)

    # Подписание наряда мастером (Y = 9.95)
    t7 = draw_task(page, 15.6, 9.95, TW, TH, "8. Подписание наряда,\nприемка ключей и ТС,\nпередача в цех")
    connect(page, 15.6, 5.3, 15.6, 9.525, dashed=True)
    connect_elbow(page, [(14.825, 2.45), (15.0, 2.45), (15.0, 9.95), (14.675, 9.95)])

    # Входная диагностика (Диагност Y = 7.45)
    t8 = draw_task(page, 17.4, 7.45, TW, TH, "9. Входная диагностика:\nсканирование кодов DTC,\nпроверка турбины и ТНВД")
    connect_elbow(page, [(16.525, 9.95), (17.4, 9.95), (17.4, 7.875)])

    gw3 = draw_gateway(page, 18.9, 7.45, 0.3, "ДВС исправен?", q_pos="below")
    connect(page, 18.325, 7.45, 18.6, 7.45)

    # Брак ДВС
    t_fail3 = draw_task(page, 18.9, 9.95, TW, TH, "10. Акт дефектовки ДВС\nи направление авто\nна ремонт")
    connect(page, 18.9, 7.75, 18.9, 9.525, label="Нет [Брак]", label_offset_x=0.45, label_offset_y=0.0)
    ev_fail3 = draw_event(page, 20.3, 9.95, 0.25, "Конец: Снятие авто\nс тюнинга", "EndFail")
    connect(page, 19.825, 9.95, 20.05, 9.95)

    # Калибровщик (Гарифуллин К.Р. Y = 4.95): Вычитывание стока
    t9 = draw_task(page, 20.2, 4.95, TW, TH, "11. Считывание сток-дампа\nпрошивки ROM в режиме\nBench / OBD-II")
    connect_elbow(page, [(19.2, 7.45), (20.2, 7.45), (20.2, 5.375)], label="Да [Норма]")

    # КИС: автобэкап стока (Y = 2.45)
    t10 = draw_task(page, 20.2, 2.45, TW, TH, "12. Автоархивирование\nстока в защищенную БД\n(хэш SHA-256)")
    connect(page, 20.2, 4.525, 20.2, 2.875)

    # Калибровка и запись ПО
    t11 = draw_task(page, 22.0, 4.95, TW, TH, "13. Модификация карт в WinOLS,\nпересчет контрольной суммы\nCRC32 и запись в ЭБУ")
    connect_elbow(page, [(21.125, 2.45), (22.0, 2.45), (22.0, 4.525)])

    gw4 = draw_gateway(page, 23.4, 4.95, 0.3, "Запись OK?", q_pos="below")
    connect(page, 22.925, 4.95, 23.1, 4.95)

    # Сбой записи
    t_rec = draw_task(page, 23.4, 6.0, TW, 0.75, "14. Recovery Mode:\nАварийный откат на сток")
    connect(page, 23.4, 5.25, 23.4, 5.625, label="Сбой", label_offset_x=0.3, label_offset_y=0.0)
    connect_elbow(page, [(23.4, 6.375), (23.4, 6.55), (22.0, 6.55), (22.0, 5.375)], label="Повтор", label_seg=1, label_offset_y=0.15)

    # ОБЯЗАТЕЛЬНО: ВЫПОЛНЕНИЕ + ТЕХНИЧЕСКИЙ КОММЕНТАРИЙ МАСТЕРА (КАРИМ)
    t12 = draw_task(page, 24.8, 4.95, TW, TH, "15. Сброс адаптаций ЭБУ,\nочистка кодов DTC\n+ ТЕХ. КОММЕНТАРИЙ")
    connect(page, 23.7, 4.95, 23.875, 4.95, label="Успех")

    # Выходной тест-драйв (Диагност Y = 7.45)
    t13 = draw_task(page, 24.8, 7.45, TW, TH, "16. Выходной тест-драйв:\nзамер мощности на стенде,\nлогирование наддува")
    connect(page, 24.8, 5.375, 24.8, 7.025)

    # КИС: Генерация Акта (Y = 2.45)
    t14 = draw_task(page, 24.8, 2.45, TW, TH, "17. Фиксация гарантии 24 мес.,\nрасчет стоимости и выпуск\nАкта выполненных работ")
    connect_elbow(page, [(24.8, 7.875), (25.5, 7.875), (25.5, 2.45), (25.725, 2.45)])

    # ДОКУМЕНТ 2: АКТ ВЫПОЛНЕННЫХ РАБОТ (ПОДПРОЦЕСС КАРИМА)
    doc_act = draw_document(page, 27.0, 4.8, 1.8, 1.0, "Акт выполненных\nработ + Гарантия", "Документ Карима")
    connect(page, 25.725, 2.45, 27.0, 4.3, dashed=True)

    # Мастер: выдача авто (Y = 9.95)
    t15 = draw_task(page, 27.0, 9.95, TW, TH, "18. Подписание акта,\nкассовый расчет и выдача\nавтомобиля клиенту")
    connect(page, 27.0, 5.3, 27.0, 9.525, dashed=True)

    # Финал
    ev_ok = draw_event(page, 28.5, 9.95, 0.25, "Финал: Автомобиль\nвыдан клиенту,\nзаказ закрыт", "EndSuccess")
    connect(page, 27.925, 9.95, 28.25, 9.95)

    p_vsd = os.path.join(out_dir, "ДБС_Общая_динамическая_модель_BPMN.vsd")
    p_vsdx = os.path.join(out_dir, "ДБС_Общая_динамическая_модель_BPMN.vsdx")
    p_png = os.path.join(out_dir, "ДБС_Общая_динамическая_модель_BPMN.png")
    for p in [p_vsd, p_vsdx, p_png]:
        if os.path.exists(p): os.remove(p)
    doc.SaveAs(p_vsd)
    doc.SaveAs(p_vsdx)
    page.Export(p_png)
    doc.Saved = True
    doc.Close()
    print("Общая динамическая модель сохранена:", p_png)

try:
    create_subprocess_1_danil()
    create_subprocess_2_karim()
    create_unified_bpmn_overview()
    print("ВСЕ 3 МОДЕЛИ BPMN УСПЕШНО СФОРМИРОВАНЫ В ПАПКЕ ДБС!")
except Exception as e:
    print("Ошибка генерации:", e)
    import traceback
    traceback.print_exc()
finally:
    visio.Quit()

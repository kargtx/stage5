import win32com.client
import os

print("=== Generating BPMN 2.0 Dynamic Models: Danil, Karim, and Overview ===")

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
    # Strict requirement: ONLY BLUE for tasks
    shape.Cells("FillForegnd").FormulaU = "RGB(239, 246, 255)" # Soft clean blue
    shape.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"     # Royal blue border
    shape.Cells("LeftMargin").FormulaU = "0.05 in"
    shape.Cells("RightMargin").FormulaU = "0.05 in"
    shape.Cells("TopMargin").FormulaU = "0.05 in"
    shape.Cells("BottomMargin").FormulaU = "0.05 in"
    shape.Text = (prefix + "\n" if prefix else "") + text
    set_font(shape, 8, bold=False, color="RGB(15, 23, 42)")
    return shape

def draw_document(page, x, y, w, h, title, subtext=""):
    f = 0.22 # fold size
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

    # Corner fold triangle
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
    lbl.Cells("FillPattern").FormulaU = "0" # Transparent
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
        qlbl.Cells("FillPattern").FormulaU = "0" # Transparent
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
        lbl.Cells("FillPattern").FormulaU = "0" # STRICT REQUIREMENT: TRANSPARENT BACKGROUND!
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
        lbl.Cells("FillPattern").FormulaU = "0" # STRICT REQUIREMENT: TRANSPARENT BACKGROUND!
        lbl.Cells("LinePattern").FormulaU = "0"
        set_font(lbl, 7.5, bold=True, color="RGB(30, 41, 59)")

def setup_base_pool(page, w, h, pool_title, lanes):
    page.PageSheet.Cells("PageWidth").FormulaU = f"{w} in"
    page.PageSheet.Cells("PageHeight").FormulaU = f"{h} in"

    # Top banner
    tb = page.DrawRectangle(0.5, h - 1.1, w - 0.5, h - 0.5)
    tb.Text = pool_title
    tb.Cells("FillForegnd").FormulaU = "RGB(15, 23, 42)"
    tb.Cells("LineColor").FormulaU = "RGB(15, 23, 42)"
    set_font(tb, 12, bold=True, color="RGB(255,255,255)")

    # Left Pool vertical header
    ph = page.DrawRectangle(0.4, 1.2, 1.5, h - 1.3)
    ph.Text = "ПУЛ: АВТОСЕРВИС (BPMN 2.0)"
    ph.Cells("FillForegnd").FormulaU = "RGB(30, 41, 59)"
    ph.Cells("LineColor").FormulaU = "RGB(15, 23, 42)"
    ph.Cells("LineWeight").FormulaU = "1.5 pt"
    ph.Cells("LeftMargin").FormulaU = "0.02 in"
    ph.Cells("RightMargin").FormulaU = "0.02 in"
    set_font(ph, 10, bold=True, color="RGB(255,255,255)")
    ph.Cells("TxtAngle").FormulaU = "90 deg"

    # Lanes
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

    # Bottom legend
    lg = page.DrawRectangle(3.1, 0.35, w - 0.5, 1.05)
    lg.Cells("FillForegnd").FormulaU = "RGB(248, 250, 252)"
    lg.Cells("LineColor").FormulaU = "RGB(203, 213, 225)"
    lg.Cells("LineWeight").FormulaU = "1 pt"
    lg.Text = "СТАНДАРТ BPMN 2.0:  (O) Начальное / Конечное событие  |  [X] Шлюз исключающего выбора XOR\n" \
              "[ ] Действие (Задача исполнителя / КИС) — строго синий стиль  |  [ Документ ] Объект данных BPMN  |  Подписи стрелок: прозрачный фон"
    set_font(lg, 8.5, bold=False, color="RGB(51, 65, 85)")

# ==============================================================================
# 1. ПОДПРОЦЕСС ДАНИЛА: «ОФОРМЛЕНИЕ ЗАЯВКИ КЛИЕНТА -> НАРЯД-ЗАКАЗ»
# ==============================================================================
def create_danil_diagram():
    print("Building Danil's subprocess: Оформление заявки клиента...")
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)
    W, H = 20.0, 12.0
    TW, TH = 1.85, 0.85

    lanes = [
        ("Клиент\n(Автовладелец)", 8.5, 10.7, "RGB(255, 253, 245)"),
        ("Мастер-приемщик\n(Сервисный консультант)", 5.0, 8.5, "RGB(242, 253, 245)"),
        ("КИС «Автосервис»\n(CRM, Склад, Каталог)", 1.2, 5.0, "RGB(240, 247, 255)")
    ]

    setup_base_pool(page, W, H, "ПОДПРОЦЕСС 1: ОФОРМЛЕНИЕ ЗАЯВКИ КЛИЕНТА И ВЫПУСК НАРЯД-ЗАКАЗА (СТУДЕНТ: ДАНИЛ)", lanes)

    # 1. Lane Клиент (Y = 9.6)
    ev_start = draw_event(page, 3.6, 9.6, 0.25, "Старт: Обращение\nклиента", "Start")
    t_c1 = draw_task(page, 5.0, 9.6, TW, TH, "1.1 Озвучивание пожеланий\nи предоставление авто\n(VIN, модель, задача)")
    connect(page, 3.85, 9.6, 4.075, 9.6)

    # 2. Lane Мастер (Y = 6.75)
    t_m1 = draw_task(page, 5.0, 6.75, TW, TH, "1.2 Первичный опрос,\nосмотр авто и ввод\nданных в КИС")
    connect(page, 5.0, 9.175, 5.0, 7.175)

    # 3. Lane КИС (Y = 3.1)
    t_k1 = draw_task(page, 7.0, 3.1, TW, TH, "1.3 Поиск в базе КИС:\nистория авто, тип ЭБУ\nи проверка допуска")
    connect_elbow(page, [(5.925, 6.75), (7.0, 6.75), (7.0, 3.525)])

    # Gateway 1 in Master lane (X = 8.8, Y = 6.75): Возможность чип-тюнинга
    gw_1 = draw_gateway(page, 8.8, 6.75, 0.3, "Тюнинг возможен?", q_pos="below")
    connect_elbow(page, [(7.925, 3.1), (8.8, 3.1), (8.8, 6.45)])

    # Branch NO -> End Fail
    t_fail1 = draw_task(page, 8.8, 9.6, TW, TH, "1.4 Информирование\nклиента о технических\nограничениях")
    connect(page, 8.8, 7.05, 8.8, 9.175, label="Нет [Отказ]", label_offset_x=0.45, label_offset_y=0.0)
    ev_fail = draw_event(page, 10.4, 9.6, 0.25, "Конец: Заявка отклонена\n(несовместимость)", "EndFail")
    connect(page, 9.725, 9.6, 10.15, 9.6)

    # Branch YES -> Master consults & calculation in EIS
    t_m2 = draw_task(page, 10.8, 6.75, TW, TH, "1.5 Консультация по пакетам\n(Stage 1, Евро-2, DPF)\nи смете работ")
    connect(page, 9.1, 6.75, 9.875, 6.75, label="Да [Норма]")

    t_k2 = draw_task(page, 10.8, 3.1, TW, TH, "1.6 Автоматический расчет\nсметы в КИС по прайсу\nи фиксация скидок")
    connect(page, 10.8, 6.325, 10.8, 3.525)

    # Client reviews estimate
    t_c2 = draw_task(page, 12.8, 9.6, TW, TH, "1.7 Рассмотрение сметы\nи условий тюнинга\nклиентом")
    connect_elbow(page, [(11.725, 3.1), (12.8, 3.1), (12.8, 9.175)], label="Смета готова", label_seg=1, label_offset_x=0.55)

    # Gateway 2: Клиент согласен? (X = 14.6, Y = 9.6)
    gw_2 = draw_gateway(page, 14.6, 9.6, 0.3, "Клиент согласен?", q_pos="above")
    connect(page, 13.725, 9.6, 14.3, 9.6)

    # Branch NO
    t_fail2 = draw_task(page, 16.5, 9.6, TW, TH, "1.8 Фиксация причины\nотказа в CRM-системе")
    connect(page, 14.9, 9.6, 15.575, 9.6, label="Нет [Отказ]")
    ev_fail2 = draw_event(page, 18.1, 9.6, 0.25, "Конец: Отказ клиента\nот услуги", "EndFail")
    connect(page, 17.425, 9.6, 17.85, 9.6)

    # Branch YES -> Formation of Work Order (Наряд-заказ) in EIS
    t_k3 = draw_task(page, 14.6, 3.1, TW, TH, "1.9 Формирование наряда:\nприсвоение ID заказа,\nпривязка к VIN и ТЗ")
    connect(page, 14.6, 9.3, 14.6, 3.525, label="Да [Согласовано]", label_offset_x=0.65, label_offset_y=0.0)

    # DATA OBJECT: НАРИД-ЗАКАЗ (DOCUMENT)
    doc_order = draw_document(page, 16.8, 4.3, 1.8, 1.0, "Наряд-заказ\n№ КИС-2026", "Ключевой документ")
    connect(page, 15.525, 3.1, 15.9, 4.3, dashed=True) # Data association from generation to doc

    # Master: signing order and taking keys
    t_m3 = draw_task(page, 16.8, 6.75, TW, TH, "1.10 Подписание наряда\nклиентом и мастером,\nприемка ключей и ТС")
    connect(page, 16.8, 4.8, 16.8, 6.325, dashed=True) # Data association into signing
    connect_elbow(page, [(15.525, 3.1), (15.7, 3.1), (15.7, 6.75), (15.875, 6.75)]) # Sequence flow from EIS to Master

    ev_ok = draw_event(page, 18.7, 6.75, 0.25, "Финал: Наряд-заказ\nготов и передан\nв цех тюнинга", "EndSuccess")
    connect(page, 17.725, 6.75, 18.45, 6.75)

    p_vsdx = os.path.abspath(r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\Подпроцесс_1_Оформление_заявки_Данил_BPMN.vsdx")
    p_png = os.path.abspath(r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\Подпроцесс_1_Оформление_заявки_Данил_BPMN.png")
    if os.path.exists(p_vsdx): os.remove(p_vsdx)
    if os.path.exists(p_png): os.remove(p_png)
    doc.SaveAs(p_vsdx)
    page.Export(p_png)
    doc.Saved = True
    doc.Close()
    print("Danil's subprocess OK!")

# ==============================================================================
# 2. ПОДПРОЦЕСС КАРИМА: «ИСПОЛНЕНИЕ ЗАКАЗА -> ВЫПОЛНЕНИЕ + КОММЕНТ -> АКТ»
# ==============================================================================
def create_karim_diagram():
    print("Building Karim's subprocess: Исполнение наряд-заказа...")
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)
    W, H = 21.5, 13.0
    TW, TH = 1.85, 0.85

    lanes = [
        ("Мастер-приемщик\n(Диспетчеризация и выдача)", 9.5, 11.7, "RGB(255, 251, 240)"),
        ("Инженер-диагност\n(Входной/выходной контроль)", 6.7, 9.5, "RGB(242, 253, 245)"),
        ("Специалист по чип-тюнингу\n(Калибровщик ЭБУ)", 3.8, 6.7, "RGB(240, 247, 255)"),
        ("КИС «Автосервис: Тюнинг»\n(Модуль калибровки и БД)", 1.2, 3.8, "RGB(252, 246, 255)")
    ]

    setup_base_pool(page, W, H, "ПОДПРОЦЕСС 2: ИСПОЛНЕНИЕ НАРИД-ЗАКАЗА И ВЫДАЧА АКТА (СТУДЕНТ: КАРИМ)", lanes)

    # 1. INPUT DOCUMENT: НАРИД-ЗАКАЗ (Placed cleanly above Step 2.1)
    ev_start = draw_event(page, 3.6, 8.1, 0.25, "Старт: Наряд-заказ\nпоступил в цех", "Start")
    doc_in = draw_document(page, 5.2, 10.6, 1.8, 1.0, "Наряд-заказ", "Входной документ")
    
    # Sequence flow from start to 2.1
    t_d1 = draw_task(page, 5.2, 8.1, TW, TH, "2.1 Входная диагностика ЭСУД:\nпроверка параметров,\nчтение кодов DTC")
    connect(page, 3.85, 8.1, 4.275, 8.1)
    connect(page, 5.2, 10.1, 5.2, 8.525, dashed=True) # Input data association

    gw_d = draw_gateway(page, 7.2, 8.1, 0.3, "ДВС исправен?", q_pos="below")
    connect(page, 6.125, 8.1, 6.9, 8.1)

    # Fail branch (Defect)
    t_fail = draw_task(page, 7.2, 10.6, TW, TH, "2.2 Оформление акта\nдефектовки в КИС и отказ\nот прошивки")
    connect(page, 7.2, 8.4, 7.2, 10.175, label="Нет [Ошибки]", label_offset_x=0.45, label_offset_y=0.0)
    ev_fail = draw_event(page, 8.9, 10.6, 0.25, "Конец: Снятие с поста\n(неисправность)", "EndFail")
    connect(page, 8.125, 10.6, 8.65, 10.6)

    # OK branch -> Тюнер: подключение и вычитывание стока (Y = 5.25)
    t_t1 = draw_task(page, 9.0, 5.25, TW, TH, "2.3 Подключение питания\nАКБ (14V), считывание\nсток-дампа ROM")
    connect_elbow(page, [(7.5, 8.1), (9.0, 8.1), (9.0, 5.675)], label="Да [Допуск]", label_seg=0, label_offset_y=0.15)

    # EIS: автобэкап стока (Y = 2.5)
    t_k1 = draw_task(page, 9.0, 2.5, TW, TH, "2.4 Автосохранение стока\nв архиве КИС\n(хэш SHA-256)")
    connect(page, 9.0, 4.825, 9.0, 2.925)

    # Тюнер: калибровка и запись
    t_t2 = draw_task(page, 11.2, 5.25, TW, TH, "2.5 Модификация карт\n(Stage 1 / Евро-2) в WinOLS,\nпересчет CRC и запись")
    connect_elbow(page, [(9.925, 2.5), (11.2, 2.5), (11.2, 4.825)])

    gw_flash = draw_gateway(page, 12.9, 5.25, 0.3, "Запись OK?", q_pos="below")
    connect(page, 12.125, 5.25, 12.6, 5.25)

    # Fail branch: Recovery
    t_rec = draw_task(page, 12.9, 6.35, TW, 0.75, "2.6 Recovery Mode:\nАварийный откат на сток\n(режим Bench / Boot)")
    connect(page, 12.9, 5.55, 12.9, 5.975, label="Сбой", label_offset_x=0.3, label_offset_y=0.0)
    connect_elbow(page, [(12.9, 6.725), (12.9, 6.95), (11.2, 6.95), (11.2, 5.675)], label="Повтор", label_seg=1, label_offset_y=0.15)

    # OK branch: Reset adaptations & TECHNICAL COMMENT (Выполнение + коммент)
    t_t3 = draw_task(page, 14.8, 5.25, TW, TH, "2.7 Сброс адаптаций ЭБУ,\nочистка сервисных DTC\n+ ТЕХ. КОММЕНТАРИЙ")
    connect(page, 13.2, 5.25, 13.875, 5.25, label="Успех")

    # Output log test drive by Diagnostician (Y = 8.1)
    t_d2 = draw_task(page, 14.8, 8.1, TW, TH, "2.8 Выходной тест-драйв:\nснятие логов параметров\n(наддув, детонация, AFR)")
    connect(page, 14.8, 5.675, 14.8, 7.675)

    gw_log = draw_gateway(page, 16.7, 8.1, 0.3, "Логи в норме?", q_pos="above")
    connect(page, 15.725, 8.1, 16.4, 8.1)

    connect_elbow(page, [(16.7, 8.4), (16.7, 9.0), (11.2, 9.0), (11.2, 5.675)], label="Отклонение [Доработка]", label_seg=1, label_offset_y=0.15)

    # EIS: Generation of Work Completion Act (АКТ ВЫПОЛНЕННЫХ РАБОТ)
    t_k2 = draw_task(page, 16.7, 2.5, TW, TH, "2.9 Архивирование прошивки\nв КИС, расчет гарантии\nи генерация Акта")
    connect(page, 16.7, 7.8, 16.7, 2.925, label="Норма [Да]", label_offset_x=0.45, label_offset_y=0.0)

    # DATA OBJECT: АКТ ВЫПОЛНЕННЫХ РАБОТ
    doc_act = draw_document(page, 19.0, 4.3, 1.9, 1.05, "Акт выполненных\nработ + Отчет", "Выходной документ")
    connect(page, 17.625, 2.5, 19.0, 3.775, dashed=True) # Data association from EIS

    # Master: Final check and delivery
    t_m_final = draw_task(page, 19.0, 10.6, TW, TH, "2.10 Подписание акта,\nвыдача авто и гарантийного\nсертификата клиенту")
    connect(page, 19.0, 4.825, 19.0, 10.175, dashed=True)
    connect_elbow(page, [(17.625, 2.5), (17.8, 2.5), (17.8, 10.6), (18.075, 10.6)]) # Sequence flow from EIS to Master

    ev_ok = draw_event(page, 20.7, 10.6, 0.25, "Финал: Акт подписан,\nавтомобиль сдан\nклиенту", "EndSuccess")
    connect(page, 19.925, 10.6, 20.45, 10.6)

    p_vsdx = os.path.abspath(r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\Подпроцесс_2_Исполнение_заказа_Карим_BPMN.vsdx")
    p_png = os.path.abspath(r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\Подпроцесс_2_Исполнение_заказа_Карим_BPMN.png")
    if os.path.exists(p_vsdx): os.remove(p_vsdx)
    if os.path.exists(p_png): os.remove(p_png)
    doc.SaveAs(p_vsdx)
    page.Export(p_png)
    doc.Saved = True
    doc.Close()
    print("Karim's subprocess OK!")

# ==============================================================================
# 3. ОБЩАЯ СКВОЗНАЯ ДИНАМИЧЕСКАЯ МОДЕЛЬ (ПЕРЕРАБОТАННАЯ В ЕДИНОМ СТИЛЕ)
# ==============================================================================
def create_overview_diagram():
    print("Building General Overview BPMN model...")
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)
    W, H = 22.5, 13.5
    TW, TH = 1.85, 0.85

    lanes = [
        ("Мастер-приемщик\n(Сервисный консультант)", 9.8, 12.2, "RGB(255, 251, 240)"),
        ("Инженер-диагност\n(Входной/выходной контроль)", 7.0, 9.8, "RGB(242, 253, 245)"),
        ("Специалист по чип-тюнингу\n(Калибровщик ЭБУ)", 4.0, 7.0, "RGB(240, 247, 255)"),
        ("КИС «Автосервис: Тюнинг»\n(Модуль калибровки и БД)", 1.2, 4.0, "RGB(252, 246, 255)")
    ]

    setup_base_pool(page, W, H, "ОБЩАЯ СКВОЗНАЯ ДИНАМИЧЕСКАЯ МОДЕЛЬ ПРОЦЕССА ЧИП-ТЮНИНГА (BPMN 2.0)", lanes)

    # 1. Start event
    ev_start = draw_event(page, 3.6, 11.0, 0.25, "Старт: Обращение\nклиента", "Start")
    t1 = draw_task(page, 5.0, 11.0, TW, TH, "1. Оформление заявки:\nосмотр авто и согласование\nтехнического задания")
    connect(page, 3.85, 11.0, 4.075, 11.0)

    # Output document from Step 1: Наряд-заказ (Placed cleanly between lanes)
    doc_order = draw_document(page, 6.9, 9.8, 1.7, 0.95, "Наряд-заказ", "Итог этапа 1")
    connect(page, 5.925, 11.0, 6.9, 10.3, dashed=True)

    # Step 2: Diagnostician
    t2 = draw_task(page, 6.9, 8.4, TW, TH, "2. Входная диагностика ЭСУД:\nчтение кодов DTC\nи проверка систем ДВС")
    connect(page, 5.0, 10.575, 5.0, 8.4)
    connect(page, 5.0, 8.4, 5.975, 8.4)
    connect(page, 6.9, 9.3, 6.9, 8.825, dashed=True)

    gw1 = draw_gateway(page, 8.8, 8.4, 0.3, "ДВС исправен?", q_pos="below")
    connect(page, 7.825, 8.4, 8.5, 8.4)

    # Fail
    t_fail = draw_task(page, 8.8, 11.0, TW, TH, "3. Акт дефектовки в КИС\nи мотивированный отказ\nв чип-тюнинге")
    connect(page, 8.8, 8.7, 8.8, 10.575, label="Нет [Ошибки]", label_offset_x=0.45, label_offset_y=0.0)
    ev_fail = draw_event(page, 10.4, 11.0, 0.25, "Конец: Отказ в услуге\n(неисправность)", "EndFail")
    connect(page, 9.725, 11.0, 10.15, 11.0)

    # Tuner step
    t3 = draw_task(page, 10.5, 5.5, TW, TH, "4. Подключение питания АКБ,\nвычитывание сток-дампа\nи сохранение в КИС")
    connect_elbow(page, [(9.1, 8.4), (10.5, 8.4), (10.5, 5.925)], label="Да [Допуск]", label_seg=0, label_offset_y=0.15)

    t_k1 = draw_task(page, 10.5, 2.6, TW, TH, "5. Архивация стока в КИС\n(хэш SHA-256) и подбор\nкалибровок по ТЗ")
    connect(page, 10.5, 5.075, 10.5, 3.025)

    # Tuner flash
    t4 = draw_task(page, 12.7, 5.5, TW, TH, "6. Калибровка в WinOLS,\nпересчет CRC и запись\nво flash-память ЭБУ")
    connect_elbow(page, [(11.425, 2.6), (12.7, 2.6), (12.7, 5.075)])

    gw2 = draw_gateway(page, 14.4, 5.5, 0.3, "Запись OK?", q_pos="below")
    connect(page, 13.625, 5.5, 14.1, 5.5)

    t_rec = draw_task(page, 14.4, 6.45, TW, 0.75, "7. Recovery Mode:\nАварийный откат на сток\n(режим Bench)")
    connect(page, 14.4, 5.8, 14.4, 6.075, label="Сбой", label_offset_x=0.3, label_offset_y=0.0)
    connect_elbow(page, [(14.4, 6.825), (14.4, 7.05), (12.7, 7.05), (12.7, 5.925)], label="Повтор", label_seg=1, label_offset_y=0.15)

    t5 = draw_task(page, 16.3, 5.5, TW, TH, "8. Сброс адаптаций ЭБУ,\nочистка кодов DTC\n+ комментарий мастера")
    connect(page, 14.7, 5.5, 15.375, 5.5, label="Успех")

    # Test drive
    t6 = draw_task(page, 16.3, 8.4, TW, TH, "9. Выходной тест-драйв:\nснятие динамических логов\n(давление, детонация, AFR)")
    connect(page, 16.3, 5.925, 16.3, 7.975)

    gw3 = draw_gateway(page, 18.1, 8.4, 0.3, "Логи в норме?", q_pos="above")
    connect(page, 17.225, 8.4, 17.8, 8.4)

    connect_elbow(page, [(18.1, 8.7), (18.1, 9.3), (12.7, 9.3), (12.7, 5.925)], label="Отклонение [Доработка]", label_seg=1, label_offset_y=0.15)

    # Act generation in EIS
    t_k2 = draw_task(page, 18.1, 2.6, TW, TH, "10. Архивирование тюнинг-файла,\nрасчет стоимости и генерация\nАкта выполненных работ")
    connect(page, 18.1, 8.1, 18.1, 3.025, label="Норма [Да]", label_offset_x=0.45, label_offset_y=0.0)

    # Output document: АКТ ВЫПОЛНЕННЫХ РАБОТ
    doc_act = draw_document(page, 20.0, 4.3, 1.8, 1.0, "Акт выполненных\nработ", "Итог этапа 2")
    connect(page, 19.025, 2.6, 20.0, 3.8, dashed=True)

    # Master: Final closing
    t7 = draw_task(page, 20.0, 11.0, TW, TH, "11. Закрытие наряда в КИС,\nвыдача авто, гарантийного\nсертификата и расчет")
    connect(page, 20.0, 4.8, 20.0, 10.575, dashed=True)
    connect_elbow(page, [(19.025, 2.6), (19.2, 2.6), (19.2, 11.0), (19.075, 11.0)])

    ev_ok = draw_event(page, 21.7, 11.0, 0.25, "Финал: Авто сдан,\nуслуга успешно\nзавершена", "EndSuccess")
    connect(page, 20.925, 11.0, 21.45, 11.0)

    p_vsdx = os.path.abspath(r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\Общая_динамическая_модель_чип_тюнинг_BPMN.vsdx")
    p_png = os.path.abspath(r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\Общая_динамическая_модель_чип_тюнинг_BPMN.png")
    if os.path.exists(p_vsdx): os.remove(p_vsdx)
    if os.path.exists(p_png): os.remove(p_png)
    doc.SaveAs(p_vsdx)
    page.Export(p_png)
    doc.Saved = True
    doc.Close()
    print("Overview model OK!")

try:
    create_danil_diagram()
    create_karim_diagram()
    create_overview_diagram()
    print("ALL 3 DIAGRAMS GENERATED SUCCESSFULLY!")
except Exception as e:
    print("Error:", e)
    import traceback
    traceback.print_exc()
finally:
    visio.Quit()

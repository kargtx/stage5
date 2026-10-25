import win32com.client
import os

print("=== Generating Detailed Business Scenario (DBS) Process Diagram in Microsoft Visio ===")

visio = win32com.client.Dispatch("Visio.Application")
visio.Visible = False
visio.AlertResponse = 7

def set_font(shape, size_pt=9.0, bold=False, color="RGB(15, 23, 42)"):
    shape.Cells("Char.Size").FormulaU = f"{size_pt} pt"
    shape.Cells("Char.Color").FormulaU = color
    shape.Cells("Char.Style").FormulaU = "1" if bold else "0"

def draw_swimlane(page, y_top, y_bottom, x_left, x_right, title, color_header="RGB(30, 58, 138)", color_bg="RGB(248, 250, 252)"):
    lane = page.DrawRectangle(x_left, y_bottom, x_right, y_top)
    lane.Cells("FillForegnd").FormulaU = color_bg
    lane.Cells("FillPattern").FormulaU = "1"
    lane.Cells("LineColor").FormulaU = "RGB(203, 213, 225)"
    lane.Cells("LineWeight").FormulaU = "1 pt"

    hdr_w = 2.4
    hdr = page.DrawRectangle(x_left, y_bottom, x_left + hdr_w, y_top)
    hdr.Cells("FillForegnd").FormulaU = color_header
    hdr.Cells("LineColor").FormulaU = "RGB(203, 213, 225)"
    hdr.Cells("LineWeight").FormulaU = "1 pt"
    hdr.Text = title
    set_font(hdr, 10.5, bold=True, color="RGB(255, 255, 255)")
    hdr.Cells("Para.HorzAlign").FormulaU = "1"
    hdr.Cells("VerticalAlign").FormulaU = "1"
    return lane

def draw_step_task(page, cx, cy, w, h, step_num, title, task_type="auto"):
    x1, y1 = cx - w/2, cy - h/2
    x2, y2 = cx + w/2, cy + h/2
    box = page.DrawRectangle(x1, y1, x2, y2)
    box.Cells("Rounding").FormulaU = "0.08 in"

    if task_type == "auto":
        box.Cells("FillForegnd").FormulaU = "RGB(219, 234, 254)" # Soft blue
        box.Cells("LineColor").FormulaU = "RGB(29, 78, 216)"    # Royal navy
        box.Cells("LineWeight").FormulaU = "1.5 pt"
        badge = " [КИС]"
    elif task_type == "system":
        box.Cells("FillForegnd").FormulaU = "RGB(238, 242, 255)" # Soft indigo
        box.Cells("LineColor").FormulaU = "RGB(79, 70, 229)"    # Indigo
        box.Cells("LineWeight").FormulaU = "1.5 pt"
        badge = " [СИСТЕМА]"
    else: # manual
        box.Cells("FillForegnd").FormulaU = "RGB(241, 245, 249)" # Light slate
        box.Cells("LineColor").FormulaU = "RGB(71, 85, 105)"    # Slate
        box.Cells("LineWeight").FormulaU = "1.2 pt"
        badge = " [ВНЕ ИС]"

    box.Text = f"{step_num}{badge}\n{title}"
    set_font(box, 8.5, bold=False, color="RGB(15, 23, 42)")
    box.Cells("Para.HorzAlign").FormulaU = "1"
    box.Cells("VerticalAlign").FormulaU = "1"
    return {"cx": cx, "cy": cy, "w": w, "h": h, "top": (cx, y2), "bottom": (cx, y1), "left": (x1, cy), "right": (x2, cy)}

def draw_gateway(page, cx, cy, size=0.7):
    pts = [
        (cx, cy + size/2),
        (cx + size/2, cy),
        (cx, cy - size/2),
        (cx - size/2, cy),
        (cx, cy + size/2)
    ]
    for i in range(len(pts) - 1):
        l = page.DrawLine(pts[i][0], pts[i][1], pts[i+1][0], pts[i+1][1])
        l.Cells("LineColor").FormulaU = "RGB(217, 119, 6)"
        l.Cells("LineWeight").FormulaU = "2 pt"

    fill_box = page.DrawRectangle(cx - size/3, cy - size/3, cx + size/3, cy + size/3)
    fill_box.Cells("FillForegnd").FormulaU = "RGB(254, 243, 199)"
    fill_box.Cells("LinePattern").FormulaU = "0"
    fill_box.Text = "X"
    set_font(fill_box, 13.0, bold=True, color="RGB(180, 83, 9)")
    fill_box.Cells("Para.HorzAlign").FormulaU = "1"
    fill_box.Cells("VerticalAlign").FormulaU = "1"

    return {"cx": cx, "cy": cy, "top": (cx, cy + size/2), "bottom": (cx, cy - size/2),
            "left": (cx - size/2, cy), "right": (cx + size/2, cy)}

def draw_doc_shape(page, cx, cy, w, h, doc_title):
    x1, y1 = cx - w/2, cy - h/2
    x2, y2 = cx + w/2, cy + h/2
    box = page.DrawRectangle(x1, y1, x2, y2)
    box.Cells("FillForegnd").FormulaU = "RGB(255, 255, 255)"
    box.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
    box.Cells("LineWeight").FormulaU = "1.5 pt"
    box.Cells("Rounding").FormulaU = "0.02 in"
    box.Text = f"📄 {doc_title}"
    set_font(box, 8.5, bold=True, color="RGB(30, 58, 138)")
    box.Cells("Para.HorzAlign").FormulaU = "1"
    box.Cells("VerticalAlign").FormulaU = "1"
    return {"cx": cx, "cy": cy, "top": (cx, y2), "bottom": (cx, y1), "left": (x1, cy), "right": (x2, cy)}

def draw_start_event(page, cx, cy, r=0.28):
    c = page.DrawOval(cx - r, cy - r, cx + r, cy + r)
    c.Cells("FillForegnd").FormulaU = "RGB(220, 252, 231)"
    c.Cells("LineColor").FormulaU = "RGB(22, 163, 74)"
    c.Cells("LineWeight").FormulaU = "2.5 pt"
    c.Text = "Старт"
    set_font(c, 7.5, bold=True, color="RGB(22, 101, 52)")
    c.Cells("Para.HorzAlign").FormulaU = "1"
    c.Cells("VerticalAlign").FormulaU = "1"
    return {"cx": cx, "cy": cy, "top": (cx, cy + r), "bottom": (cx, cy - r), "left": (cx - r, cy), "right": (cx + r, cy)}

def draw_end_event(page, cx, cy, r=0.28, is_error=False):
    c = page.DrawOval(cx - r, cy - r, cx + r, cy + r)
    if is_error:
        c.Cells("FillForegnd").FormulaU = "RGB(254, 226, 226)"
        c.Cells("LineColor").FormulaU = "RGB(220, 38, 38)"
        c.Cells("LineWeight").FormulaU = "3.5 pt"
        c.Text = "Отказ"
        set_font(c, 7.5, bold=True, color="RGB(153, 27, 27)")
    else:
        c.Cells("FillForegnd").FormulaU = "RGB(224, 231, 255)"
        c.Cells("LineColor").FormulaU = "RGB(49, 46, 129)"
        c.Cells("LineWeight").FormulaU = "3.5 pt"
        c.Text = "Финиш"
        set_font(c, 7.5, bold=True, color="RGB(30, 27, 75)")
    c.Cells("Para.HorzAlign").FormulaU = "1"
    c.Cells("VerticalAlign").FormulaU = "1"
    return {"cx": cx, "cy": cy, "top": (cx, cy + r), "bottom": (cx, cy - r), "left": (cx - r, cy), "right": (cx + r, cy)}

def connect_flow(page, pt1, pt2, label="", dashed=False):
    s = page.DrawLine(pt1[0], pt1[1], pt2[0], pt2[1])
    s.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
    s.Cells("LineWeight").FormulaU = "1.5 pt"
    if dashed:
        s.Cells("LinePattern").FormulaU = "2"
    s.Cells("EndArrow").FormulaU = "13"
    s.Cells("EndArrowSize").FormulaU = "2"
    if label:
        cx, cy = (pt1[0] + pt2[0]) / 2, (pt1[1] + pt2[1]) / 2 + 0.14
        lbl = page.DrawRectangle(cx - 0.7, cy - 0.12, cx + 0.7, cy + 0.12)
        lbl.Cells("FillForegnd").FormulaU = "RGB(255, 255, 255)"
        lbl.Cells("LinePattern").FormulaU = "0"
        lbl.Text = label
        set_font(lbl, 8.0, bold=True, color="RGB(30, 58, 138)")
        lbl.Cells("Para.HorzAlign").FormulaU = "1"
        lbl.Cells("VerticalAlign").FormulaU = "1"

def connect_flow_elbow(page, waypoints, label="", dashed=False):
    for i in range(len(waypoints) - 1):
        pA, pB = waypoints[i], waypoints[i+1]
        s = page.DrawLine(pA[0], pA[1], pB[0], pB[1])
        s.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
        s.Cells("LineWeight").FormulaU = "1.5 pt"
        if dashed:
            s.Cells("LinePattern").FormulaU = "2"
        if i == len(waypoints) - 2:
            s.Cells("EndArrow").FormulaU = "13"
            s.Cells("EndArrowSize").FormulaU = "2"
    if label and len(waypoints) >= 2:
        mid_idx = len(waypoints) // 2
        cx = (waypoints[mid_idx-1][0] + waypoints[mid_idx][0]) / 2
        cy = (waypoints[mid_idx-1][1] + waypoints[mid_idx][1]) / 2 + 0.14
        lbl = page.DrawRectangle(cx - 0.7, cy - 0.12, cx + 0.7, cy + 0.12)
        lbl.Cells("FillForegnd").FormulaU = "RGB(255, 255, 255)"
        lbl.Cells("LinePattern").FormulaU = "0"
        lbl.Text = label
        set_font(lbl, 8.0, bold=True, color="RGB(30, 58, 138)")
        lbl.Cells("Para.HorzAlign").FormulaU = "1"
        lbl.Cells("VerticalAlign").FormulaU = "1"

def build_visio_dbs():
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)

    PAGE_W = 28.5
    PAGE_H = 16.0
    page.PageSheet.Cells("PageWidth").FormulaU = f"{PAGE_W} in"
    page.PageSheet.Cells("PageHeight").FormulaU = f"{PAGE_H} in"

    # Header Banner
    banner = page.DrawRectangle(0.8, 14.8, 27.7, 15.6)
    banner.Cells("FillForegnd").FormulaU = "RGB(29, 78, 216)"
    banner.Cells("LinePattern").FormulaU = "0"
    banner.Text = "ДЕТАЛЬНЫЙ БИЗНЕС-СЦЕНАРИЙ (ДБС): ОКАЗАНИЕ УСЛУГ ЧИП-ТЮНИНГА ДВС В КИС «АВТОСЕРВИС»\nСквозной процесс: Заявка -> Наряд-заказ (Данил) -> Диагностика -> Исполнение + Коммент (Карим) -> Акт выполненных работ"
    set_font(banner, 11.5, bold=True, color="RGB(255, 255, 255)")
    banner.Cells("Para.HorzAlign").FormulaU = "1"
    banner.Cells("VerticalAlign").FormulaU = "1"

    # 5 Swimlanes
    L_X1 = 0.8
    L_X2 = 27.7
    
    Y_L1_TOP, Y_L1_BOT = 14.5, 12.1 # Клиент
    Y_L2_TOP, Y_L2_BOT = 12.1, 9.4  # Мастер-приемщик (Данил)
    Y_L3_TOP, Y_L3_BOT = 9.4, 6.7   # Специалист-диагност
    Y_L4_TOP, Y_L4_BOT = 6.7, 4.0   # Инженер-калибровщик (Карим)
    Y_L5_TOP, Y_L5_BOT = 4.0, 1.4   # КИС «Автосервис»

    draw_swimlane(page, Y_L1_TOP, Y_L1_BOT, L_X1, L_X2, "1. Клиент\n(Автовладелец)", "RGB(30, 58, 138)", "RGB(255, 255, 255)")
    draw_swimlane(page, Y_L2_TOP, Y_L2_BOT, L_X1, L_X2, "2. Мастер-приемщик\n(Контур Данила)", "RGB(29, 78, 216)", "RGB(248, 250, 252)")
    draw_swimlane(page, Y_L3_TOP, Y_L3_BOT, L_X1, L_X2, "3. Специалист-диагност\n(Вход/Выход ДВС)", "RGB(2, 132, 199)", "RGB(255, 255, 255)")
    draw_swimlane(page, Y_L4_TOP, Y_L4_BOT, L_X1, L_X2, "4. Инженер-калибровщик\n(Контур Карима)", "RGB(79, 70, 229)", "RGB(248, 250, 252)")
    draw_swimlane(page, Y_L5_TOP, Y_L5_BOT, L_X1, L_X2, "5. КИС «Автосервис»\n(Системные сервисы)", "RGB(15, 23, 42)", "RGB(241, 245, 249)")

    # Legend
    leg = page.DrawRectangle(19.0, 0.35, 27.7, 1.25)
    leg.Cells("FillForegnd").FormulaU = "RGB(255, 255, 255)"
    leg.Cells("LineColor").FormulaU = "RGB(203, 213, 225)"
    leg.Text = "ОБОЗНАЧЕНИЯ ДБС:\n[КИС] — Автоматизированный сеанс системы | [ВНЕ ИС] — Внесистемное физическое действие | [СИСТЕМА] — Фоновый сервис КИС\n📄 — Документ КИС | ◇ — Логический шлюз развилки (XOR) | О — Событие (Старт / Финиш)"
    set_font(leg, 8.5, bold=False, color="RGB(30, 41, 59)")
    leg.Cells("Para.HorzAlign").FormulaU = "1"
    leg.Cells("VerticalAlign").FormulaU = "1"

    TW, TH = 1.7, 0.9

    # LANE 1: КЛИЕНТ
    ev_start = draw_start_event(page, cx=3.6, cy=13.3)
    t1_1 = draw_step_task(page, cx=5.3, cy=13.3, w=TW, h=TH, step_num="Шаг 1.1",
                          title="Обращение в сервис,\nцели тюнинга (Stage 1/2)", task_type="manual")
    t1_7 = draw_step_task(page, cx=12.2, cy=13.3, w=TW, h=TH, step_num="Шаг 1.7",
                          title="Рассмотрение сметы\nи условий гарантии", task_type="manual")
    gw_agree = draw_gateway(page, cx=13.9, cy=13.3)
    ev_rej_client = draw_end_event(page, cx=13.9, cy=14.1, is_error=True)

    t1_17 = draw_step_task(page, cx=25.6, cy=13.3, w=TW, h=TH, step_num="Шаг 1.17",
                           title="Приемка автомобиля,\nподписание Акта и оплата", task_type="manual")
    ev_end_success = draw_end_event(page, cx=27.2, cy=13.3, is_error=False)

    # LANE 2: МАСТЕР-ПРИЕМЩИК (ДАНИЛ)
    t1_2 = draw_step_task(page, cx=5.3, cy=10.75, w=TW, h=TH, step_num="Шаг 1.2",
                          title="Осмотр автомобиля,\nсчитывание VIN и СТС", task_type="manual")
    t1_3 = draw_step_task(page, cx=7.6, cy=10.75, w=TW, h=TH, step_num="Шаг 1.3",
                          title="Ввод данных клиента\nи ТС в КИС (Сеанс 1)", task_type="auto")
    t1_5 = draw_step_task(page, cx=10.0, cy=10.75, w=TW, h=TH, step_num="Шаг 1.5",
                          title="Подбор прошивки,\nконсультация клиента", task_type="manual")
    t1_8 = draw_step_task(page, cx=13.9, cy=10.75, w=TW, h=TH, step_num="Шаг 1.8",
                          title="Утверждение и печать\nНаряд-заказа (Сеанс 3)", task_type="auto")
    doc_wo = draw_doc_shape(page, cx=13.9, cy=9.8, w=1.8, h=0.42, doc_title="Наряд-заказ")

    t1_16 = draw_step_task(page, cx=25.6, cy=10.75, w=TW, h=TH, step_num="Шаг 1.16",
                           title="Оформление Акта\nи закрытие наряда (Сеанс 6)", task_type="auto")
    doc_act = draw_doc_shape(page, cx=24.1, cy=9.8, w=1.8, h=0.42, doc_title="Акт сдачи-приемки")

    # LANE 3: СПЕЦИАЛИСТ-ДИАГНОСТ
    t1_9 = draw_step_task(page, cx=16.2, cy=8.05, w=TW, h=TH, step_num="Шаг 1.9",
                          title="Подключение OBD-II,\nсканирование кодов DTC", task_type="manual")
    t1_10 = draw_step_task(page, cx=18.3, cy=8.05, w=TW, h=TH, step_num="Шаг 1.10",
                           title="Заполнение диагн.\nкарты ДВС (Сеанс 4)", task_type="auto")
    doc_diag = draw_doc_shape(page, cx=18.3, cy=7.1, w=1.7, h=0.42, doc_title="Диагн. карта")
    gw_dtc = draw_gateway(page, cx=19.8, cy=8.05)
    ev_rej_dtc = draw_end_event(page, cx=19.8, cy=8.9, is_error=True)

    # LANE 4: ИНЖЕНЕР-КАЛИБРОВЩИК (КАРИМ)
    t1_11 = draw_step_task(page, cx=19.8, cy=5.35, w=TW, h=TH, step_num="Шаг 1.11",
                           title="Считывание сток-ROM,\nпроверка защиты ЭБУ", task_type="manual")
    t1_12 = draw_step_task(page, cx=21.8, cy=5.35, w=TW, h=TH, step_num="Шаг 1.12",
                           title="Калибровка карт ДВС,\nпересчет CRC32", task_type="auto")
    t1_13 = draw_step_task(page, cx=23.8, cy=5.35, w=TW, h=TH, step_num="Шаг 1.13",
                           title="Запись ПО в ЭБУ +\nКОММЕНТАРИЙ (Сеанс 5)", task_type="auto")
    t1_14 = draw_step_task(page, cx=25.7, cy=5.35, w=TW, h=TH, step_num="Шаг 1.14",
                           title="Выходной тест-драйв,\nзамер параметров", task_type="manual")

    # LANE 5: КИС «АВТОСЕРВИС»
    t1_4 = draw_step_task(page, cx=7.6, cy=2.7, w=TW, h=TH, step_num="Шаг 1.4",
                          title="Проверка совместимости\nЭБУ по базе данных", task_type="system")
    t1_6 = draw_step_task(page, cx=10.0, cy=2.7, w=TW, h=TH, step_num="Шаг 1.6",
                          title="Авто-калькуляция сметы\nпо прайс-листу и скидкам", task_type="system")
    t1_12b = draw_step_task(page, cx=21.8, cy=2.7, w=TW, h=TH, step_num="Шаг 1.12b",
                            title="Архивирование сток-ROM\nс хэшем SHA-256 в БД", task_type="system")
    t1_18 = draw_step_task(page, cx=23.8, cy=2.7, w=TW, h=TH, step_num="Шаг 1.18",
                           title="Проводка оплаты, списание\nработ и фискализация", task_type="system")

    # ================= CONNECTORS / FLOW =================
    # Client Start -> 1.1
    connect_flow(page, ev_start["right"], t1_1["left"])
    # 1.1 -> 1.2
    connect_flow(page, t1_1["bottom"], t1_2["top"])
    # 1.2 -> 1.3
    connect_flow(page, t1_2["right"], t1_3["left"])
    # 1.3 -> 1.4
    connect_flow(page, t1_3["bottom"], t1_4["top"])
    # 1.4 -> 1.5
    connect_flow_elbow(page, [t1_4["right"], (8.9, 2.7), (8.9, 10.75), t1_5["left"]])
    # 1.5 -> 1.6
    connect_flow_elbow(page, [t1_5["bottom"], t1_6["top"]])
    # 1.6 -> 1.7
    connect_flow_elbow(page, [t1_6["right"], (11.1, 2.7), (11.1, 13.3), t1_7["left"]])
    # 1.7 -> Gateway Agree?
    connect_flow(page, t1_7["right"], gw_agree["left"])
    # Gateway -> Reject (up)
    connect_flow(page, gw_agree["top"], ev_rej_client["bottom"], label="Отказ")
    # Gateway -> 1.8 (down to manager Work Order)
    connect_flow(page, gw_agree["bottom"], t1_8["top"], label="Согласовано")
    # 1.8 -> Document Work Order
    connect_flow(page, t1_8["bottom"], doc_wo["top"], dashed=True)

    # 1.8 -> 1.9 (Diagnostics)
    connect_flow_elbow(page, [t1_8["right"], (15.2, 10.75), (15.2, 8.05), t1_9["left"]])
    # 1.9 -> 1.10
    connect_flow(page, t1_9["right"], t1_10["left"])
    # 1.10 -> doc Diag Card
    connect_flow(page, t1_10["bottom"], doc_diag["top"], dashed=True)
    # 1.10 -> Gateway DTC?
    connect_flow(page, t1_10["right"], gw_dtc["left"])
    # Gateway DTC -> Reject (up)
    connect_flow(page, gw_dtc["top"], ev_rej_dtc["bottom"], label="DTC > 0")
    # Gateway DTC -> 1.11 (down to Tuner)
    connect_flow(page, gw_dtc["bottom"], t1_11["top"], label="Норма")

    # 1.11 -> 1.12
    connect_flow(page, t1_11["right"], t1_12["left"])
    # 1.12 -> 1.12b (System Backup)
    connect_flow(page, t1_12["bottom"], t1_12b["top"], dashed=True)
    # 1.12 -> 1.13 (Flash + Comment)
    connect_flow(page, t1_12["right"], t1_13["left"])
    # 1.13 -> 1.14 (Test drive)
    connect_flow(page, t1_13["right"], t1_14["left"])
    # 1.14 -> 1.16 (Up to Manager Act)
    connect_flow(page, t1_14["top"], t1_16["bottom"])
    # 1.16 -> doc Act
    connect_flow(page, (24.7, 10.3), doc_act["top"], dashed=True)
    # 1.16 -> 1.17 (Up to Client Acceptance)
    connect_flow(page, t1_16["top"], t1_17["bottom"])
    # 1.17 -> 1.18 (Down to System Fiscalization)
    connect_flow_elbow(page, [(26.4, 12.85), (26.8, 12.85), (26.8, 2.7), (24.7, 2.7)], dashed=True)
    # 1.17 -> End Success
    connect_flow(page, t1_17["right"], ev_end_success["left"])

    # Output paths
    out_dir = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС"
    p_vsd = os.path.abspath(os.path.join(out_dir, "ДБС_Схема_бизнес_процесса.vsd"))
    p_vsdx = os.path.abspath(os.path.join(out_dir, "ДБС_Схема_бизнес_процесса.vsdx"))
    p_png = os.path.abspath(os.path.join(out_dir, "ДБС_Схема_бизнес_процесса.png"))

    for p in [p_vsd, p_vsdx, p_png]:
        if os.path.exists(p):
            try: os.remove(p)
            except: pass

    doc.SaveAs(p_vsd)
    print("SUCCESS: Saved VSD to:", p_vsd)

    try:
        with open(p_vsd, "rb") as f_in, open(p_vsdx, "wb") as f_out:
            f_out.write(f_in.read())
        print("SUCCESS: Copied VSDX to:", p_vsdx)
    except Exception as e:
        print("Copy error:", e)

    page.Export(p_png)
    print("SUCCESS: Exported PNG to:", p_png)

    doc.Saved = True
    doc.Close()

try:
    build_visio_dbs()
except Exception as e:
    print("Error:", e)
    import traceback
    traceback.print_exc()
finally:
    visio.Quit()

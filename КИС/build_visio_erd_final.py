import win32com.client
import os
import math

print("=== Generating Final 3NF ER-Model (Corrected Crow's Foot Standard) in Microsoft Visio ===")

visio = win32com.client.Dispatch("Visio.Application")
visio.Visible = False
visio.AlertResponse = 7

def set_font(shape, size_pt=10.5, bold=False, color="RGB(15, 23, 42)"):
    shape.Cells("Char.Size").FormulaU = f"{size_pt} pt"
    shape.Cells("Char.Color").FormulaU = color
    shape.Cells("Char.Style").FormulaU = "1" if bold else "0"

def draw_entity_table(page, x, top_y, w, title, subtitle, pk_list, fk_list, attr_list):
    """
    Standard Crow's foot entity table with row dividers:
    - Header: Deep Royal Blue (RGB 29, 78, 216) with White text
    - Keys (PK / FK): Soft cyan (RGB 219, 234, 254), Navy text, NO trailing '(FK)'
    - Attributes: White / Off-white, Slate text
    """
    hdr_h = 0.72
    row_h_key = 0.35
    row_h_attr = 0.32

    x1, x2 = x - w/2, x + w/2
    curr_y = top_y

    # 1. Header
    hdr = page.DrawRectangle(x1, curr_y - hdr_h, x2, curr_y)
    hdr.Cells("FillForegnd").FormulaU = "RGB(29, 78, 216)"
    hdr.Cells("LineColor").FormulaU = "RGB(30, 58, 138)"
    hdr.Cells("LineWeight").FormulaU = "1.5 pt"
    hdr.Cells("TopMargin").FormulaU = "0.04 in"
    hdr.Cells("BottomMargin").FormulaU = "0.04 in"
    hdr.Text = f"{title}\n{subtitle}" if subtitle else title
    set_font(hdr, 11.5, bold=True, color="RGB(255, 255, 255)")
    hdr.Cells("Para.HorzAlign").FormulaU = "1" # Center
    curr_y -= hdr_h

    # 2. PK rows
    for pk in pk_list:
        r = page.DrawRectangle(x1, curr_y - row_h_key, x2, curr_y)
        r.Cells("FillForegnd").FormulaU = "RGB(219, 234, 254)"
        r.Cells("LineColor").FormulaU = "RGB(147, 197, 253)"
        r.Cells("LineWeight").FormulaU = "1 pt"
        r.Cells("LeftMargin").FormulaU = "0.12 in"
        r.Cells("RightMargin").FormulaU = "0.06 in"
        r.Text = f"[PK]  {pk}"
        set_font(r, 11.0, bold=True, color="RGB(15, 23, 42)")
        r.Cells("Para.HorzAlign").FormulaU = "0" # Left
        r.Cells("VerticalAlign").FormulaU = "1"  # Middle
        curr_y -= row_h_key

    # 3. FK rows (Clean [FK] at start, NO '(FK)' at end!)
    for fk in fk_list:
        r = page.DrawRectangle(x1, curr_y - row_h_key, x2, curr_y)
        r.Cells("FillForegnd").FormulaU = "RGB(219, 234, 254)"
        r.Cells("LineColor").FormulaU = "RGB(147, 197, 253)"
        r.Cells("LineWeight").FormulaU = "1 pt"
        r.Cells("LeftMargin").FormulaU = "0.12 in"
        r.Cells("RightMargin").FormulaU = "0.06 in"
        r.Text = f"[FK]  {fk}"
        set_font(r, 10.5, bold=True, color="RGB(30, 58, 138)")
        r.Cells("Para.HorzAlign").FormulaU = "0"
        r.Cells("VerticalAlign").FormulaU = "1"
        curr_y -= row_h_key

    # 4. Ordinary Attributes
    for idx, attr in enumerate(attr_list):
        r = page.DrawRectangle(x1, curr_y - row_h_attr, x2, curr_y)
        fill = "RGB(255, 255, 255)" if idx % 2 == 0 else "RGB(248, 250, 252)"
        r.Cells("FillForegnd").FormulaU = fill
        r.Cells("LineColor").FormulaU = "RGB(226, 232, 240)"
        r.Cells("LineWeight").FormulaU = "1 pt"
        r.Cells("LeftMargin").FormulaU = "0.12 in"
        r.Cells("RightMargin").FormulaU = "0.06 in"
        r.Text = f"       {attr}"
        set_font(r, 10.5, bold=False, color="RGB(30, 41, 59)")
        r.Cells("Para.HorzAlign").FormulaU = "0"
        r.Cells("VerticalAlign").FormulaU = "1"
        curr_y -= row_h_attr

    total_h = top_y - curr_y
    center_y = (top_y + curr_y) / 2

    # Outer border
    outer = page.DrawRectangle(x1, curr_y, x2, top_y)
    outer.Cells("FillPattern").FormulaU = "0" # Transparent fill
    outer.Cells("LineColor").FormulaU = "RGB(30, 58, 138)"
    outer.Cells("LineWeight").FormulaU = "1.5 pt"

    return {
        "x": x, "y": center_y, "w": w, "h": total_h,
        "x1": x1, "x2": x2, "y1": curr_y, "y2": top_y,
        "top": (x, top_y),
        "bottom": (x, curr_y),
        "left": (x1, center_y),
        "right": (x2, center_y)
    }

def draw_pill_badge(page, cx, cy, text, w_lbl=2.4, h_lbl=0.34):
    """
    Crisp opaque pill badge with white background and soft blue border
    so lines never cut through the text!
    """
    lbl = page.DrawRectangle(cx - w_lbl/2, cy - h_lbl/2, cx + w_lbl/2, cy + h_lbl/2)
    lbl.Cells("FillForegnd").FormulaU = "RGB(255, 255, 255)"
    lbl.Cells("FillPattern").FormulaU = "1" # Solid Opaque
    lbl.Cells("LineColor").FormulaU = "RGB(147, 197, 253)"
    lbl.Cells("LineWeight").FormulaU = "1 pt"
    lbl.Cells("Rounding").FormulaU = "0.05 in"
    lbl.Cells("TopMargin").FormulaU = "0.02 in"
    lbl.Cells("BottomMargin").FormulaU = "0.02 in"
    lbl.Text = text
    set_font(lbl, 10.0, bold=True, color="RGB(30, 58, 138)")
    lbl.Cells("Para.HorzAlign").FormulaU = "1"
    lbl.Cells("VerticalAlign").FormulaU = "1"
    return lbl

def draw_crows_foot_end(page, p_near, p_end, card_type="N"):
    """
    Draws Crow's Foot end symbols:
    - card_type == '1'   : Mandatory One '||' (two tick marks)
    - card_type == 'N'   : Many '|<' (tick mark + crow's foot fork)
    - card_type == '0..1': Optional One '0|' (circle + tick mark)
    """
    vx = p_end[0] - p_near[0]
    vy = p_end[1] - p_near[1]
    L = math.hypot(vx, vy)
    if L < 0.001:
        return
    ux, uy = vx / L, vy / L
    px, py = -uy, ux # Perpendicular unit vector

    line_color = "RGB(37, 99, 235)"
    tick_len = 0.12

    if card_type == "1":
        # Two perpendicular tick marks: d1 = 0.10, d2 = 0.20 from p_end
        for d in [0.10, 0.20]:
            cx, cy = p_end[0] - ux * d, p_end[1] - uy * d
            s = page.DrawLine(cx - px * tick_len, cy - py * tick_len, cx + px * tick_len, cy + py * tick_len)
            s.Cells("LineColor").FormulaU = line_color
            s.Cells("LineWeight").FormulaU = "1.5 pt"

    elif card_type == "N":
        # Mandatory Many: Tick mark at d = 0.25, Crow's foot fork at d = 0.18 to p_end
        # 1. Tick mark
        d_tick = 0.25
        cx, cy = p_end[0] - ux * d_tick, p_end[1] - uy * d_tick
        s = page.DrawLine(cx - px * tick_len, cy - py * tick_len, cx + px * tick_len, cy + py * tick_len)
        s.Cells("LineColor").FormulaU = line_color
        s.Cells("LineWeight").FormulaU = "1.5 pt"
        # 2. Fork prongs
        d_fork = 0.18
        fx, fy = p_end[0] - ux * d_fork, p_end[1] - uy * d_fork
        s1 = page.DrawLine(fx, fy, p_end[0] + px * tick_len, p_end[1] + py * tick_len)
        s1.Cells("LineColor").FormulaU = line_color
        s1.Cells("LineWeight").FormulaU = "1.5 pt"
        s2 = page.DrawLine(fx, fy, p_end[0] - px * tick_len, p_end[1] - py * tick_len)
        s2.Cells("LineColor").FormulaU = line_color
        s2.Cells("LineWeight").FormulaU = "1.5 pt"

    elif card_type == "0..1":
        # Optional One: Circle at d = 0.22 (radius = 0.06), Tick mark at d = 0.10
        # 1. Circle
        d_circ = 0.22
        r_circ = 0.06
        cx, cy = p_end[0] - ux * d_circ, p_end[1] - uy * d_circ
        circ = page.DrawOval(cx - r_circ, cy - r_circ, cx + r_circ, cy + r_circ)
        circ.Cells("FillForegnd").FormulaU = "RGB(255, 255, 255)"
        circ.Cells("FillPattern").FormulaU = "1" # Solid white
        circ.Cells("LineColor").FormulaU = line_color
        circ.Cells("LineWeight").FormulaU = "1.5 pt"
        # 2. Tick mark
        d_tick = 0.10
        tx, ty = p_end[0] - ux * d_tick, p_end[1] - uy * d_tick
        s = page.DrawLine(tx - px * tick_len, ty - py * tick_len, tx + px * tick_len, ty + py * tick_len)
        s.Cells("LineColor").FormulaU = line_color
        s.Cells("LineWeight").FormulaU = "1.5 pt"

def draw_straight_connection(page, pt1, pt2, label, card_start="1", card_end="N",
                             label_cx=None, label_cy=None, label_w=2.2):
    """
    Connects pt1 to pt2 with a straight line, badge in between, and Crow's foot symbols.
    """
    if label_cx is None:
        label_cx = (pt1[0] + pt2[0]) / 2
    if label_cy is None:
        label_cy = (pt1[1] + pt2[1]) / 2

    badge_half_w = label_w / 2 + 0.05
    badge_half_h = 0.22

    # Check horizontal or vertical
    if abs(pt1[1] - pt2[1]) < 0.05:
        # Horizontal
        y = pt1[1]
        x_min, x_max = min(pt1[0], pt2[0]), max(pt1[0], pt2[0])
        # Segment 1: from x_min to badge left
        s1 = page.DrawLine(x_min, y, label_cx - badge_half_w, y)
        s1.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
        s1.Cells("LineWeight").FormulaU = "1.5 pt"
        # Segment 2: from badge right to x_max
        s2 = page.DrawLine(label_cx + badge_half_w, y, x_max, y)
        s2.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
        s2.Cells("LineWeight").FormulaU = "1.5 pt"
    elif abs(pt1[0] - pt2[0]) < 0.05:
        # Vertical
        x = pt1[0]
        y_min, y_max = min(pt1[1], pt2[1]), max(pt1[1], pt2[1])
        s1 = page.DrawLine(x, y_min, x, label_cy - badge_half_h)
        s1.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
        s1.Cells("LineWeight").FormulaU = "1.5 pt"
        s2 = page.DrawLine(x, label_cy + badge_half_h, x, y_max)
        s2.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
        s2.Cells("LineWeight").FormulaU = "1.5 pt"
    else:
        # Fallback simple line
        s = page.DrawLine(pt1[0], pt1[1], pt2[0], pt2[1])
        s.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
        s.Cells("LineWeight").FormulaU = "1.5 pt"

    draw_pill_badge(page, label_cx, label_cy, label, w_lbl=label_w)

    # Start symbol
    draw_crows_foot_end(page, (label_cx, label_cy), pt1, card_type=card_start)
    # End symbol
    draw_crows_foot_end(page, (label_cx, label_cy), pt2, card_type=card_end)

def draw_elbow_connection(page, waypoints, label, card_start="1", card_end="N",
                          label_cx=None, label_cy=None, label_w=2.2):
    """
    Draws orthogonal multi-segment line with inline badge and Crow's foot symbols.
    """
    for i in range(len(waypoints) - 1):
        pA, pB = waypoints[i], waypoints[i+1]
        line = page.DrawLine(pA[0], pA[1], pB[0], pB[1])
        line.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
        line.Cells("LineWeight").FormulaU = "1.5 pt"

    if label_cx is not None and label_cy is not None:
        draw_pill_badge(page, label_cx, label_cy, label, w_lbl=label_w)

    # Start symbol (near waypoints[1] pointing to waypoints[0])
    draw_crows_foot_end(page, waypoints[1], waypoints[0], card_type=card_start)
    # End symbol (near waypoints[-2] pointing to waypoints[-1])
    draw_crows_foot_end(page, waypoints[-2], waypoints[-1], card_type=card_end)

def build_erd_final():
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)

    PAGE_W = 30.0
    PAGE_H = 16.0
    page.PageSheet.Cells("PageWidth").FormulaU = f"{PAGE_W} in"
    page.PageSheet.Cells("PageHeight").FormulaU = f"{PAGE_H} in"

    # ================= BANNER =================
    banner = page.DrawRectangle(0.8, 14.8, 29.2, 15.6)
    banner.Cells("FillForegnd").FormulaU = "RGB(29, 78, 216)"
    banner.Cells("LinePattern").FormulaU = "0"
    banner.Text = "ОБЩАЯ РЕЛЯЦИОННАЯ МОДЕЛЬ ДАННЫХ (ERD) КИС «АВТОСЕРВИС: ЧИП-ТЮНИНГ»\nАрхитектура 3NF (нотация Crow's Foot): Клиенты, Автомобили, Заявки, Наряд-заказ, Диагностика, Исполнение и Акт"
    set_font(banner, 12.0, bold=True, color="RGB(255, 255, 255)")
    banner.Cells("Para.HorzAlign").FormulaU = "1"
    banner.Cells("VerticalAlign").FormulaU = "1"

    # Coordinates
    EW = 4.2
    X_COL1 = 4.2
    X_COL2 = 11.3
    X_COL3 = 18.4
    X_COL4 = 25.5

    Y_ROW1_TOP = 13.5
    Y_ROW2_TOP = 8.5

    # ================= ROW 1 (TOP ENTITIES) =================
    # 1. CLIENT (Col 1)
    e_client = draw_entity_table(
        page, x=X_COL1, top_y=Y_ROW1_TOP, w=EW,
        title="Client (Клиенты)",
        subtitle="Учет автовладельцев и заказчиков",
        pk_list=["id_client : INT"],
        fk_list=[],
        attr_list=[
            "full_name : VARCHAR(100)",
            "phone_number : VARCHAR(20)",
            "client_type : VARCHAR(20)"
        ]
    )

    # 2. VEHICLE (Col 2) - Атомарные поля по 1NF/3NF (brand, model, engine_model, ecu_type)
    e_vehicle = draw_entity_table(
        page, x=X_COL2, top_y=Y_ROW1_TOP, w=EW,
        title="Vehicle (Автомобили)",
        subtitle="Обслуживаемые транспортные средства",
        pk_list=["vin_code : VARCHAR(17)"],
        fk_list=[
            "id_client : INT"
        ],
        attr_list=[
            "brand : VARCHAR(50)",
            "model : VARCHAR(50)",
            "engine_model : VARCHAR(50)",
            "ecu_type : VARCHAR(50)",
            "production_year : INT"
        ]
    )

    # 3. DIAGNOSTIC_CARD (Col 3) - Атомарный код ошибки dtc_fault_code
    e_diag = draw_entity_table(
        page, x=X_COL3, top_y=Y_ROW1_TOP, w=EW,
        title="Diagnostic_Card (Диагностика)",
        subtitle="Входной и выходной контроль ДВС",
        pk_list=["id_diag : INT"],
        fk_list=[
            "id_order : INT",
            "id_diagnostician : INT"
        ],
        attr_list=[
            "diag_stage : VARCHAR(20)",
            "dtc_fault_code : VARCHAR(10)",
            "diag_conclusion : VARCHAR(50)"
        ]
    )

    # 4. EMPLOYEE (Col 4)
    e_emp = draw_entity_table(
        page, x=X_COL4, top_y=Y_ROW1_TOP, w=EW,
        title="Employee (Сотрудники)",
        subtitle="Штатные роли автосервиса",
        pk_list=["id_employee : INT"],
        fk_list=[],
        attr_list=[
            "full_name : VARCHAR(100)",
            "role_title : VARCHAR(50)",
            "phone_number : VARCHAR(20)"
        ]
    )

    # ================= ROW 2 (BOTTOM ENTITIES: ПРОИЗВОДСТВЕННЫЙ ПОТОК) =================
    # 5. SERVICE_REQUEST (Col 1) - Связь строго через vin_code к Vehicle (без избыточного id_client!)
    e_request = draw_entity_table(
        page, x=X_COL1, top_y=Y_ROW2_TOP, w=EW,
        title="Service_Request (Заявки)",
        subtitle="Первичные обращения на чип-тюнинг",
        pk_list=["id_request : INT"],
        fk_list=[
            "vin_code : VARCHAR(17)"
        ],
        attr_list=[
            "request_date : DATETIME",
            "tuning_goals : VARCHAR(200)",
            "request_status : VARCHAR(30)"
        ]
    )

    # 6. WORK_ORDER (Col 2 - ДОКУМЕНТ ДАНИЛА) - Убран vin_code (3NF, однозначно через id_request)!
    e_order = draw_entity_table(
        page, x=X_COL2, top_y=Y_ROW2_TOP, w=EW,
        title="Work_Order (Наряд-заказ)",
        subtitle="Первичный производственный документ",
        pk_list=["id_order : INT"],
        fk_list=[
            "id_request : INT [UQ]",
            "id_master : INT"
        ],
        attr_list=[
            "order_number : VARCHAR(30)",
            "order_date : DATETIME",
            "order_status : VARCHAR(30)",
            "total_agreed_cost : DECIMAL(10,2)"
        ]
    )

    # 7. ORDER_EXECUTION (Col 3 - ПРОЦЕСС И КОММЕНТАРИЙ КАРИМА)
    e_exec = draw_entity_table(
        page, x=X_COL3, top_y=Y_ROW2_TOP, w=EW,
        title="Order_Execution (Исполнение)",
        subtitle="Выполнение калибровки и коммент мастера",
        pk_list=["id_execution : INT"],
        fk_list=[
            "id_order : INT [UQ]",
            "id_tuner : INT"
        ],
        attr_list=[
            "firmware_version : VARCHAR(50)",
            "power_increase_hp : INT",
            "technical_comment : TEXT",
            "execution_status : VARCHAR(30)"
        ]
    )

    # 8. WORK_COMPLETION_ACT (Col 4 - ДОКУМЕНТ КАРИМА) - Связь строго с Work_Order по id_order
    e_act = draw_entity_table(
        page, x=X_COL4, top_y=Y_ROW2_TOP, w=EW,
        title="Work_Completion_Act (Акт)",
        subtitle="Итоговый финансово-отчетный документ",
        pk_list=["id_act : INT"],
        fk_list=[
            "id_order : INT [UQ]",
            "id_master : INT"
        ],
        attr_list=[
            "act_number : VARCHAR(30)",
            "act_date : DATETIME",
            "total_paid_amount : DECIMAL(10,2)",
            "warranty_months : INT",
            "is_signed : BOOLEAN"
        ]
    )

    # ================= RELATIONSHIPS / CONNECTORS =================

    # 1. Client (1) -> Vehicle (N)
    # Horizontal line in Row 1 at Y = 12.2
    draw_straight_connection(page, (e_client["right"][0], 12.2), (e_vehicle["left"][0], 12.2),
                             label="1 : N [Владеет]", card_start="1", card_end="N",
                             label_cx=7.75, label_cy=12.2, label_w=2.1)

    # 2. Vehicle (1) -> Service_Request (N)
    # Route through Gutter 1-2
    draw_elbow_connection(page, [(e_vehicle["left"][0], 11.2),
                                 (7.75, 11.2),
                                 (7.75, 7.3),
                                 (e_request["right"][0], 7.3)],
                          label="1 : N [Обслуживается]", card_start="1", card_end="N",
                          label_cx=7.75, label_cy=9.25, label_w=2.4)

    # 3. Service_Request (1) -> Work_Order (0..1)
    # Horizontal line in Row 2 at Y = 6.8 (Необязательная связь 1:0..1)
    draw_straight_connection(page, (e_request["right"][0], 6.8), (e_order["left"][0], 6.8),
                             label="1 : 0..1 [Порождает]", card_start="1", card_end="0..1",
                             label_cx=7.75, label_cy=6.8, label_w=2.4)

    # 4. Work_Order (1) -> Diagnostic_Card (N)
    # Route through Gutter 2-3
    draw_elbow_connection(page, [(e_order["right"][0], 7.7),
                                 (15.2, 7.7),
                                 (15.2, 11.8),
                                 (e_diag["left"][0], 11.8)],
                          label="1 : N [Диагностика]", card_start="1", card_end="N",
                          label_cx=15.2, label_cy=9.75, label_w=2.3)

    # 5. Work_Order (1) -> Order_Execution (0..1)
    # Horizontal line in Row 2 at Y = 6.8 (Необязательная связь 1:0..1)
    draw_straight_connection(page, (e_order["right"][0], 6.8), (e_exec["left"][0], 6.8),
                             label="1 : 0..1 [Исполнение]", card_start="1", card_end="0..1",
                             label_cx=14.85, label_cy=6.8, label_w=2.4)

    # 6. Work_Order (1) -> Work_Completion_Act (0..1)
    # Highway underneath Row 2 at Y = 4.2 (Сплошная прямая линия связи по id_order!)
    draw_elbow_connection(page, [(e_order["right"][0], 5.8),
                                 (14.1, 5.8),
                                 (14.1, 4.2),
                                 (22.6, 4.2),
                                 (22.6, 5.8),
                                 (e_act["left"][0], 5.8)],
                          label="1 : 0..1 [Выдача Акта]", card_start="1", card_end="0..1",
                          label_cx=18.4, label_cy=4.2, label_w=2.4)

    # 7. Employee (1) -> Diagnostic_Card (N)
    # Horizontal line in Row 1 at Y = 12.2
    draw_straight_connection(page, (e_emp["left"][0], 12.2), (e_diag["right"][0], 12.2),
                             label="1 : N [Диагност]", card_start="1", card_end="N",
                             label_cx=21.95, label_cy=12.2, label_w=2.0)

    # 8. Employee (1) -> Work_Order (N)
    # Overhead highway at Y = 14.3, down Gutter 2-3 at X = 14.1 into Work_Order right edge at Y = 8.1
    draw_elbow_connection(page, [(e_emp["top"][0], 13.5),
                                 (e_emp["top"][0], 14.3),
                                 (14.1, 14.3),
                                 (14.1, 8.1),
                                 (e_order["right"][0], 8.1)],
                          label="1 : N [Мастер]", card_start="1", card_end="N",
                          label_cx=19.8, label_cy=14.3, label_w=2.0)

    # 9. Employee (1) -> Order_Execution (N)
    # Route through Gutter 3-4
    draw_elbow_connection(page, [(e_emp["left"][0], 11.2),
                                 (21.95, 11.2),
                                 (21.95, 7.7),
                                 (e_exec["right"][0], 7.7)],
                          label="1 : N [Тюнер]", card_start="1", card_end="N",
                          label_cx=21.95, label_cy=9.45, label_w=1.8)

    # 10. Employee (1) -> Work_Completion_Act (N)
    # Vertical line in Col 4 at X = 26.5
    draw_straight_connection(page, (26.5, e_emp["bottom"][1]), (26.5, e_act["top"][1]),
                             label="1 : N [Подписание]", card_start="1", card_end="N",
                             label_cx=26.5, label_cy=10.0, label_w=2.2)

    # ================= LEGEND (FOOTER) =================
    footer = page.DrawRectangle(0.8, 0.4, 29.2, 1.2)
    footer.Cells("FillPattern").FormulaU = "0"
    footer.Cells("LineColor").FormulaU = "RGB(147, 197, 253)"
    footer.Cells("LineWeight").FormulaU = "1 pt"
    footer.Text = (
        "СТАНДАРТ ERD (нотация Crow's Foot / «Воронья лапка», 3NF)\n"
        "|| — Родительская сущность (Mandatory 1)   |   0| — Необязательный дочерний документ (0..1 Optional)   |   |< — Дочерняя сущность (1..N Many)\n"
        "Оформление: [Синий] Заголовок сущности | [Голубой] Ключевые атрибуты ([PK] / [FK]) | Разделительные линии между строками | Однозначные внешние ключи [UQ]"
    )
    set_font(footer, 10.0, bold=False, color="RGB(30, 58, 138)")
    footer.Cells("Para.HorzAlign").FormulaU = "1"
    footer.Cells("VerticalAlign").FormulaU = "1"

    # Save to both .vsd (for native Visio 2010/2016 opening) and .vsdx
    out_dir = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС"
    p_vsd = os.path.abspath(os.path.join(out_dir, "ER_модель_КИС_ЧипТюнинг_Итоговая.vsd"))
    p_vsdx = os.path.abspath(os.path.join(out_dir, "ER_модель_КИС_ЧипТюнинг_Итоговая.vsdx"))
    p_png = os.path.abspath(os.path.join(out_dir, "ER_модель_КИС_ЧипТюнинг_Итоговая.png"))

    for p in [p_vsd, p_vsdx, p_png]:
        if os.path.exists(p):
            try: os.remove(p)
            except: pass

    # Save .vsd
    doc.SaveAs(p_vsd)
    print("SUCCESS: Saved VSD to:", p_vsd)

    # Save copy as .vsdx
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
    build_erd_final()
except Exception as e:
    print("Error:", e)
    import traceback
    traceback.print_exc()
finally:
    visio.Quit()

import win32com.client
import os

print("=== Generating Streamlined Unified ER-Model (8 Tables, 3NF) in Microsoft Visio ===")

visio = win32com.client.Dispatch("Visio.Application")
visio.Visible = False
visio.AlertResponse = 7

def set_font(shape, size_pt=10.5, bold=False, color="RGB(15, 23, 42)"):
    shape.Cells("Char.Size").FormulaU = f"{size_pt} pt"
    shape.Cells("Char.Color").FormulaU = color
    shape.Cells("Char.Style").FormulaU = "1" if bold else "0"

def draw_streamlined_entity(page, x, top_y, w, title, subtitle, pk_list, fk_list, attr_list):
    """
    Streamlined 3-tier ER Entity Card with row dividers and large readable fonts:
    1. Header: Royal Navy Blue (RGB 29, 78, 216) with White text (11.5pt / 9pt)
    2. Keys (PK/FK): Individual rows with soft cyan fill (RGB 219, 234, 254), 11pt bold Navy text, distinct border lines
    3. Attributes: Individual rows with alternating crisp white / soft tint, 10.5pt Slate text, distinct border lines
    4. Crisp outer border around the whole table
    """
    hdr_h = 0.72
    row_h_key = 0.35
    row_h_attr = 0.32

    x1, x2 = x - w/2, x + w/2
    curr_y = top_y

    # 1. Header (Entity Title & Subtitle)
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

    # 2. Key section with individual row dividers
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

    # 3. Attribute section with individual row dividers
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

    # 4. Outer border around the whole table
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

def draw_pill_label(page, cx, cy, text, w_lbl=2.4, h_lbl=0.34):
    """
    Draws a crisp inline pill-box badge with white background and soft blue border
    so that lines NEVER cut through the text!
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
    lbl.Cells("Para.HorzAlign").FormulaU = "1" # Center
    lbl.Cells("VerticalAlign").FormulaU = "1"  # Middle
    return lbl

def connect_erd_crows_foot(page, p_start, p_end, label="", card_start="1", card_end="N", dashed=False, 
                           label_cx=None, label_cy=None, label_w=2.4):
    line = page.DrawLine(p_start[0], p_start[1], p_end[0], p_end[1])
    line.Cells("LineWeight").FormulaU = "1.5 pt"
    line.Cells("LineColor").FormulaU = "RGB(30, 58, 138)"
    line.Cells("BeginArrow").FormulaU = "25" if card_start == "1" else "28"
    line.Cells("EndArrow").FormulaU = "28" if card_end == "N" else "25"
    
    if dashed:
        line.Cells("LinePattern").FormulaU = "2"
        
    if label:
        cx = label_cx if label_cx is not None else (p_start[0] + p_end[0]) / 2
        cy = label_cy if label_cy is not None else (p_start[1] + p_end[1]) / 2
        draw_pill_label(page, cx, cy, label, w_lbl=label_w)
    return line

def connect_erd_crows_foot_elbow(page, points, label="", card_start="1", card_end="N", dashed=False, 
                                 label_cx=None, label_cy=None, label_w=2.4):
    for i in range(len(points) - 1):
        p1 = points[i]
        p2 = points[i+1]
        line = page.DrawLine(p1[0], p1[1], p2[0], p2[1])
        line.Cells("LineWeight").FormulaU = "1.5 pt"
        line.Cells("LineColor").FormulaU = "RGB(30, 58, 138)"
        if dashed:
            line.Cells("LinePattern").FormulaU = "2"
        if i == 0:
            line.Cells("BeginArrow").FormulaU = "25" if card_start == "1" else "28"
        if i == len(points) - 2:
            line.Cells("EndArrow").FormulaU = "28" if card_end == "N" else "25"
            
    if label and label_cx is not None and label_cy is not None:
        draw_pill_label(page, label_cx, label_cy, label, w_lbl=label_w)

def build_erd_streamlined():
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)
    
    W, H = 30.0, 15.5
    page.PageSheet.Cells("PageWidth").FormulaU = f"{W} in"
    page.PageSheet.Cells("PageHeight").FormulaU = f"{H} in"
    
    # 1. Top Banner
    tb = page.DrawRectangle(0.8, H - 1.15, W - 0.8, H - 0.45)
    tb.Text = "ОБЩАЯ РЕЛЯЦИОННАЯ МОДЕЛЬ ДАННЫХ (ERD) КИС «АВТОСЕРВИС: ЧИП-ТЮНИНГ»\n" \
              "Базовая архитектура БД (3NF): Клиенты, Заявки, Наряд-заказ, Диагностика, Исполнение и Акт"
    tb.Cells("FillForegnd").FormulaU = "RGB(29, 78, 216)"
    tb.Cells("LineColor").FormulaU = "RGB(30, 58, 138)"
    tb.Cells("LineWeight").FormulaU = "1.5 pt"
    set_font(tb, 14.0, bold=True, color="RGB(255, 255, 255)")
    
    # 2. Bottom Legend
    lg = page.DrawRectangle(0.8, 0.35, W - 0.8, 1.00)
    lg.Cells("FillForegnd").FormulaU = "RGB(248, 250, 252)"
    lg.Cells("LineColor").FormulaU = "RGB(191, 219, 254)"
    lg.Cells("LineWeight").FormulaU = "1.2 pt"
    lg.Text = "СТАНДАРТ ERD (IDEF1X / Crow's Foot, 3NF):   || — Родительская сущность (Mandatory 1)   |   |< — Дочерняя сущность (1..N Many)   |   1 : 1 — Однозначная связь\n" \
              "Оформление сущностей: [Синий] Заголовок таблицы   |   [Голубой] Ключевые атрибуты ([PK] / [FK])   |   Разделительные линии между строками   |   Четкие подписи связей"
    set_font(lg, 10.0, bold=False, color="RGB(30, 58, 138)")

    EW = 3.80 # Width for all 8 entity tables
    
    # 4 Columns - balanced and spacious across W = 30.0 in
    X_COL1 = 4.0
    X_COL2 = 11.3
    X_COL3 = 18.7
    X_COL4 = 26.0

    # 2 Clean Rows:
    # Row 1 Top = 13.2 in (ends around Y = 10.5 in)
    # Row 2 Top = 8.5 in  (ends around Y = 5.0 in)
    Y_ROW1_TOP = 13.2
    Y_ROW2_TOP = 8.5

    # ================= ROW 1 (TOP ENTITIES: СУБЪЕКТЫ, ОБЪЕКТЫ И ПЕРСОНАЛ) =================
    # 1. CLIENT (Col 1)
    e_client = draw_streamlined_entity(
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

    # 2. VEHICLE (Col 2)
    e_vehicle = draw_streamlined_entity(
        page, x=X_COL2, top_y=Y_ROW1_TOP, w=EW,
        title="Vehicle (Автомобили)",
        subtitle="Обслуживаемые транспортные средства",
        pk_list=["vin_code : VARCHAR(17)"],
        fk_list=[
            "id_client : INT (FK)"
        ],
        attr_list=[
            "brand_model : VARCHAR(100)",
            "production_year : INT",
            "engine_and_ecu : VARCHAR(100)"
        ]
    )

    # 3. DIAGNOSTIC_CARD (Col 3)
    e_diag = draw_streamlined_entity(
        page, x=X_COL3, top_y=Y_ROW1_TOP, w=EW,
        title="Diagnostic_Card (Диагностика)",
        subtitle="Входной и выходной контроль ДВС",
        pk_list=["id_diag : INT"],
        fk_list=[
            "id_order : INT (FK)",
            "id_diagnostician : INT (FK)"
        ],
        attr_list=[
            "diag_stage : VARCHAR(20)",
            "dtc_errors : VARCHAR(150)",
            "diag_conclusion : VARCHAR(50)"
        ]
    )

    # 4. EMPLOYEE (Col 4)
    e_emp = draw_streamlined_entity(
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
    # 5. SERVICE_REQUEST (Col 1)
    e_request = draw_streamlined_entity(
        page, x=X_COL1, top_y=Y_ROW2_TOP, w=EW,
        title="Service_Request (Заявки)",
        subtitle="Первичные обращения на чип-тюнинг",
        pk_list=["id_request : INT"],
        fk_list=[
            "id_client : INT (FK)",
            "vin_code : VARCHAR(17) (FK)"
        ],
        attr_list=[
            "request_date : DATETIME",
            "tuning_goals : VARCHAR(200)",
            "request_status : VARCHAR(30)"
        ]
    )

    # 6. WORK_ORDER (Col 2 - КЛЮЧЕВОЙ ДОКУМЕНТ ДАНИЛА)
    e_order = draw_streamlined_entity(
        page, x=X_COL2, top_y=Y_ROW2_TOP, w=EW,
        title="Work_Order (Наряд-заказ)",
        subtitle="Первичный производственный документ",
        pk_list=["id_order : INT"],
        fk_list=[
            "id_request : INT (FK)",
            "vin_code : VARCHAR(17) (FK)",
            "id_master : INT (FK)"
        ],
        attr_list=[
            "order_number : VARCHAR(30)",
            "order_date : DATETIME",
            "order_status : VARCHAR(30)",
            "total_agreed_cost : DECIMAL(10,2)"
        ]
    )

    # 7. ORDER_EXECUTION (Col 3 - ВЫПОЛНЕНИЕ + КОММЕНТАРИЙ КАРИМА)
    e_exec = draw_streamlined_entity(
        page, x=X_COL3, top_y=Y_ROW2_TOP, w=EW,
        title="Order_Execution (Исполнение)",
        subtitle="Выполнение калибровки и коммент мастера",
        pk_list=["id_execution : INT"],
        fk_list=[
            "id_order : INT (FK)",
            "id_tuner : INT (FK)"
        ],
        attr_list=[
            "firmware_version : VARCHAR(50)",
            "power_increase_hp : INT",
            "technical_comment : TEXT",
            "execution_status : VARCHAR(30)"
        ]
    )

    # 8. WORK_COMPLETION_ACT (Col 4 - КЛЮЧЕВОЙ ДОКУМЕНТ КАРИМА)
    e_act = draw_streamlined_entity(
        page, x=X_COL4, top_y=Y_ROW2_TOP, w=EW,
        title="Work_Completion_Act (Акт)",
        subtitle="Итоговый финансово-отчетный документ",
        pk_list=["id_act : INT"],
        fk_list=[
            "id_order : INT (FK)",
            "id_master : INT (FK)"
        ],
        attr_list=[
            "act_number : VARCHAR(30)",
            "act_date : DATETIME",
            "total_paid_amount : DECIMAL(10,2)",
            "warranty_months : INT",
            "is_signed : BOOLEAN"
        ]
    )

    # ================= RELATIONSHIPS / CONNECTORS (ZERO OVERLAPS WITH INLINE BADGES) =================
    
    # 1. Client (1) -> Vehicle (N)
    # Horizontal line in Row 1 at Y = 12.0
    connect_erd_crows_foot(page, (e_client["right"][0], 12.0), (e_vehicle["left"][0], 12.0), 
                           label="1 : N [Владеет]", card_start="1", card_end="N",
                           label_cx=7.65, label_cy=12.0, label_w=2.1)

    # 2. Client (1) -> Service_Request (N)
    # Vertical line in Col 1 at X = 4.0
    connect_erd_crows_foot(page, e_client["bottom"], e_request["top"], 
                           label="1 : N [Подает]", card_start="1", card_end="N",
                           label_cx=4.0, label_cy=9.8, label_w=1.9)

    # 3. Vehicle (1) -> Service_Request (N)
    # Elbow in Gutter 1-2
    connect_erd_crows_foot_elbow(page, [(e_vehicle["left"][0], 11.2),
                                        (7.65, 11.2),
                                        (7.65, 7.5),
                                        (e_request["right"][0], 7.5)],
                                 label="1 : N [Указан в заявке]", card_start="1", card_end="N",
                                 label_cx=7.65, label_cy=9.3, label_w=2.4)

    # 4. Service_Request (1) -> Work_Order (1)
    # Horizontal line in Row 2 at Y = 6.8 (dashed)
    connect_erd_crows_foot(page, (e_request["right"][0], 6.8), (e_order["left"][0], 6.8), 
                           label="1 : 1 [Порождает]", card_start="1", card_end="1", dashed=True,
                           label_cx=7.65, label_cy=6.8, label_w=2.1)

    # 5. Vehicle (1) -> Work_Order (N)
    # Vertical line in Col 2 at X = 11.3
    connect_erd_crows_foot(page, e_vehicle["bottom"], e_order["top"], 
                           label="1 : N [Обслуживается]", card_start="1", card_end="N",
                           label_cx=11.3, label_cy=9.6, label_w=2.4)

    # 6. Work_Order (1) -> Diagnostic_Card (N)
    # Elbow between Col 2 and Col 3
    connect_erd_crows_foot_elbow(page, [(e_order["right"][0], 7.8),
                                        (15.0, 7.8),
                                        (15.0, 11.8),
                                        (e_diag["left"][0], 11.8)],
                                 label="1 : N [Диагностика]", card_start="1", card_end="N",
                                 label_cx=15.0, label_cy=9.8, label_w=2.3)

    # 7. Work_Order (1) -> Order_Execution (1)
    # Horizontal line in Row 2 at Y = 6.8
    connect_erd_crows_foot(page, (e_order["right"][0], 6.8), (e_exec["left"][0], 6.8), 
                           label="1 : 1 [Исполнение]", card_start="1", card_end="1",
                           label_cx=15.0, label_cy=6.8, label_w=2.1)

    # 8. Order_Execution (1) -> Work_Completion_Act (1)
    # Horizontal line in Row 2 at Y = 6.8
    connect_erd_crows_foot(page, (e_exec["right"][0], 6.8), (e_act["left"][0], 6.8), 
                           label="1 : 1 [Закрытие]", card_start="1", card_end="1",
                           label_cx=22.35, label_cy=6.8, label_w=2.0)

    # 9. Work_Order (1) -> Work_Completion_Act (1)
    # Highway underneath Row 2 at Y = 4.0 (dashed)
    connect_erd_crows_foot_elbow(page, [(e_order["right"][0], 5.6),
                                        (14.2, 5.6),
                                        (14.2, 4.0),
                                        (23.2, 4.0),
                                        (23.2, 5.6),
                                        (e_act["left"][0], 5.6)],
                                 label="1 : 1 [Выдача Акта]", card_start="1", card_end="1", dashed=True,
                                 label_cx=18.7, label_cy=4.0, label_w=2.1)

    # 10. Employee (1) -> Work_Order (N)
    # Overhead highway at Y = 13.9, down through Gutter 2-3 at X = 14.3 into Work_Order right edge at Y = 8.2
    connect_erd_crows_foot_elbow(page, [(e_emp["top"][0], 13.2),
                                        (e_emp["top"][0], 13.9),
                                        (14.3, 13.9),
                                        (14.3, 8.2),
                                        (e_order["right"][0], 8.2)],
                                 label="1 : N [Мастер]", card_start="1", card_end="N",
                                 label_cx=20.0, label_cy=13.9, label_w=2.0)

    # 11. Employee (1) -> Diagnostic_Card (N)
    # Horizontal line in Row 1 at Y = 12.0
    connect_erd_crows_foot(page, (e_emp["left"][0], 12.0), (e_diag["right"][0], 12.0), 
                           label="1 : N [Диагност]", card_start="1", card_end="N",
                           label_cx=22.35, label_cy=12.0, label_w=2.0)

    # 12. Employee (1) -> Order_Execution (N)
    # Elbow in Gutter 3-4 at X = 21.6
    connect_erd_crows_foot_elbow(page, [(e_emp["left"][0], 11.2),
                                        (21.6, 11.2),
                                        (21.6, 7.8),
                                        (e_exec["right"][0], 7.8)],
                                 label="1 : N [Тюнер]", card_start="1", card_end="N",
                                 label_cx=21.6, label_cy=9.5, label_w=1.8)

    # 13. Employee (1) -> Work_Completion_Act (N)
    # Vertical line in Col 4 at X = 26.0
    connect_erd_crows_foot(page, e_emp["bottom"], e_act["top"], 
                           label="1 : N [Подписание]", card_start="1", card_end="N",
                           label_cx=26.0, label_cy=9.8, label_w=2.2)

    # Output paths
    p_vsdx = os.path.abspath(r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ER_модель_КИС_ЧипТюнинг_Общая.vsdx")
    p_png = os.path.abspath(r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ER_модель_КИС_ЧипТюнинг_Общая.png")
    
    if os.path.exists(p_vsdx):
        try: os.remove(p_vsdx)
        except: pass
    if os.path.exists(p_png):
        try: os.remove(p_png)
        except: pass
        
    doc.SaveAs(p_vsdx)
    page.Export(p_png)
    doc.Saved = True
    doc.Close()
    print("SUCCESS: Streamlined Unified ERD generated at:")
    print(" - VSDX:", p_vsdx)
    print(" - PNG:", p_png)

try:
    build_erd_streamlined()
except Exception as e:
    print("Error:", e)
    import traceback
    traceback.print_exc()
finally:
    visio.Quit()

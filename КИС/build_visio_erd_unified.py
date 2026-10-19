import win32com.client
import os

print("=== Generating Perfect Standard Unified ER-Model (ERD) in Microsoft Visio ===")

visio = win32com.client.Dispatch("Visio.Application")
visio.Visible = False
visio.AlertResponse = 7

def set_font(shape, size_pt=10.5, bold=False, color="RGB(15, 23, 42)"):
    shape.Cells("Char.Size").FormulaU = f"{size_pt} pt"
    shape.Cells("Char.Color").FormulaU = color
    shape.Cells("Char.Style").FormulaU = "1" if bold else "0"

def draw_unified_entity(page, x, top_y, w, title, subtitle, pk_list, fk_list, attr_list):
    """
    Standard 3-tier ER Entity Card with individual rows and horizontal divider lines:
    1. Header: Royal Navy Blue (RGB 29, 78, 216) with White text (11.5pt / 9pt)
    2. Keys (PK/FK): Individual rows with light cyan fill (RGB 219, 234, 254), 11pt bold Navy text, distinct border lines
    3. Attributes: Individual rows with alternating crisp white / soft tint, 10.5pt Slate text, distinct border lines
    4. Crisp outer border around the whole table
    """
    hdr_h = 0.72
    row_h_key = 0.34
    row_h_attr = 0.31

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
        set_font(r, 10.0, bold=False, color="RGB(30, 41, 59)")
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

def draw_pill_label(page, cx, cy, text, w_lbl=2.4, h_lbl=0.32):
    """
    Draws a crisp pill-box badge with white background and soft blue border
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
        cy = label_cy if label_cy is not None else (p_start[1] + p_end[1]) / 2 + 0.28
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

def build_erd_unified():
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)
    
    W, H = 33.0, 21.5
    page.PageSheet.Cells("PageWidth").FormulaU = f"{W} in"
    page.PageSheet.Cells("PageHeight").FormulaU = f"{H} in"
    
    # 1. Top Banner
    tb = page.DrawRectangle(0.8, H - 1.25, W - 0.8, H - 0.45)
    tb.Text = "ЕДИНАЯ РЕЛЯЦИОННАЯ МОДЕЛЬ ДАННЫХ (ERD) КИС «АВТОСЕРВИС: ЧИП-ТЮНИНГ»\n" \
              "Сквозная архитектура БД (3NF): Клиенты, Наряд-заказы, Калибровка ЭБУ, Исполнение и Акты выполненных работ"
    tb.Cells("FillForegnd").FormulaU = "RGB(29, 78, 216)"
    tb.Cells("LineColor").FormulaU = "RGB(30, 58, 138)"
    tb.Cells("LineWeight").FormulaU = "1.5 pt"
    set_font(tb, 14.0, bold=True, color="RGB(255, 255, 255)")
    
    # 2. Bottom Legend
    lg = page.DrawRectangle(0.8, 0.25, W - 0.8, 0.90)
    lg.Cells("FillForegnd").FormulaU = "RGB(248, 250, 252)"
    lg.Cells("LineColor").FormulaU = "RGB(191, 219, 254)"
    lg.Cells("LineWeight").FormulaU = "1.2 pt"
    lg.Text = "СТАНДАРТ ERD (IDEF1X / Crow's Foot, 3NF):   || — Родительская сущность (Mandatory 1)   |   |< — Дочерняя сущность (1..N Many)   |   1 : 1 — Однозначная связь\n" \
              "Оформление сущностей: [Синий] Заголовок таблицы   |   [Голубой] Ключевые атрибуты ([PK] / [FK])   |   Разделительные линии между строками   |   Четкие подписи связей"
    set_font(lg, 10.0, bold=False, color="RGB(30, 58, 138)")

    EW = 3.65 # Width for all entity tables
    
    # Column Centers
    X_COL1 = 4.4
    X_COL2 = 12.2
    X_COL3 = 20.0
    X_COL4 = 27.8

    # Row Top Coordinates (Top-aligned for clean architectural rhythm)
    Y_ROW1_TOP = 19.4
    Y_ROW2_TOP = 14.6
    Y_ROW3_TOP = 9.2

    # ---------------- ROW 1 (TOP ENTITIES) ----------------
    # 1. CLIENT (Col 1)
    e_client = draw_unified_entity(
        page, x=X_COL1, top_y=Y_ROW1_TOP, w=EW,
        title="Client (Клиенты)",
        subtitle="Учет автовладельцев и контрагентов",
        pk_list=["id_client : INT"],
        fk_list=[],
        attr_list=[
            "client_type : VARCHAR(20)",
            "full_name : VARCHAR(150)",
            "phone_number : VARCHAR(20)",
            "email : VARCHAR(100)",
            "passport_inn : VARCHAR(50)",
            "discount_rate_pct : DECIMAL(4,2)",
            "registered_date : DATETIME"
        ]
    )

    # 2. VEHICLE (Col 2)
    e_vehicle = draw_unified_entity(
        page, x=X_COL2, top_y=Y_ROW1_TOP, w=EW,
        title="Vehicle (Автомобили)",
        subtitle="Реестр обслуживаемых автомобилей",
        pk_list=["vin_code : VARCHAR(17)"],
        fk_list=[
            "id_client : INT (FK)",
            "id_ecu : INT (FK)"
        ],
        attr_list=[
            "license_plate : VARCHAR(15)",
            "brand : VARCHAR(50)",
            "model : VARCHAR(50)",
            "production_year : SMALLINT",
            "engine_code : VARCHAR(30)",
            "factory_power_hp : SMALLINT",
            "current_mileage : INT"
        ]
    )

    # 3. ECU_SPECIFICATION (Col 3)
    e_ecu = draw_unified_entity(
        page, x=X_COL3, top_y=Y_ROW1_TOP, w=EW,
        title="ECU_Specification (ЭБУ)",
        subtitle="Справочник микроконтроллеров ДВС",
        pk_list=["id_ecu : INT"],
        fk_list=[],
        attr_list=[
            "ecu_vendor : VARCHAR(50)",
            "ecu_model : VARCHAR(50)",
            "mcu_processor : VARCHAR(50)",
            "flash_capacity_bytes : INT",
            "protocol_type : VARCHAR(50)",
            "connection_mode : VARCHAR(30)"
        ]
    )

    # 4. EMPLOYEE (Col 4)
    e_emp = draw_unified_entity(
        page, x=X_COL4, top_y=Y_ROW1_TOP, w=EW,
        title="Employee (Сотрудники)",
        subtitle="Штатные роли и персонал КИС",
        pk_list=["id_employee : INT"],
        fk_list=[],
        attr_list=[
            "personnel_number : VARCHAR(20)",
            "full_name : VARCHAR(150)",
            "role_title : VARCHAR(50)",
            "phone_number : VARCHAR(20)",
            "access_level : VARCHAR(30)",
            "is_active : BOOLEAN"
        ]
    )

    # ---------------- ROW 2 (MIDDLE-TOP ENTITIES) ----------------
    # 5. SERVICE_REQUEST (Col 1)
    e_request = draw_unified_entity(
        page, x=X_COL1, top_y=Y_ROW2_TOP, w=EW,
        title="Service_Request (Заявки)",
        subtitle="Первичные обращения клиентов",
        pk_list=["id_request : INT"],
        fk_list=[
            "id_client : INT (FK)",
            "vin_code : VARCHAR(17) (FK)",
            "id_employee : INT (FK)"
        ],
        attr_list=[
            "request_number : VARCHAR(30)",
            "request_date : DATETIME",
            "tuning_goals : TEXT",
            "request_status : VARCHAR(30)",
            "estimated_budget : DECIMAL(10,2)"
        ]
    )

    # 6. WORK_ORDER (Col 2 - Central Document)
    e_order = draw_unified_entity(
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
            "planned_deadline : DATETIME",
            "order_status : VARCHAR(30)",
            "total_agreed_cost : DECIMAL(10,2)",
            "intake_fuel_level : VARCHAR(20)",
            "is_client_agreed : BOOLEAN",
            "client_signature_date : DATETIME"
        ]
    )

    # 7. STOCK_FIRMWARE (Col 3)
    e_stock = draw_unified_entity(
        page, x=X_COL3, top_y=Y_ROW2_TOP, w=EW,
        title="Stock_Firmware (Сток-дамп)",
        subtitle="Резервный заводской дамп ROM (Бэкап)",
        pk_list=["id_stock : INT"],
        fk_list=[
            "vin_code : VARCHAR(17) (FK)",
            "id_order : INT (FK)"
        ],
        attr_list=[
            "hw_number : VARCHAR(50)",
            "sw_number : VARCHAR(50)",
            "calibration_id : VARCHAR(50)",
            "file_size_bytes : INT",
            "sha256_hash : VARCHAR(64)",
            "storage_path : VARCHAR(255)",
            "backup_timestamp : DATETIME"
        ]
    )

    # 8. EXECUTION_LOG (Col 4 - Выполнение + комментарий)
    e_log = draw_unified_entity(
        page, x=X_COL4, top_y=Y_ROW2_TOP, w=EW,
        title="Execution_Log (Журнал работ)",
        subtitle="Исполнение и комментарий мастера",
        pk_list=["id_log : INT"],
        fk_list=[
            "id_order : INT (FK)",
            "id_tuned : INT (FK)",
            "id_tuner : INT (FK)"
        ],
        attr_list=[
            "flasher_hardware : VARCHAR(50)",
            "battery_voltage : DECIMAL(4,2)",
            "write_duration_sec : INT",
            "write_result_status : VARCHAR(20)",
            "is_adaptations_reset : BOOLEAN",
            "technical_comment : TEXT",
            "completion_time : DATETIME"
        ]
    )

    # ---------------- ROW 3 (MIDDLE-BOTTOM ENTITIES) ----------------
    # 9. PRICE_LIST_SERVICE (Col 1)
    e_price = draw_unified_entity(
        page, x=X_COL1, top_y=Y_ROW3_TOP, w=EW,
        title="Price_List_Service (Услуги)",
        subtitle="Прейскурант и нормо-часы",
        pk_list=["id_service : INT"],
        fk_list=[],
        attr_list=[
            "service_code : VARCHAR(30)",
            "service_name : VARCHAR(150)",
            "category : VARCHAR(50)",
            "standard_hours : DECIMAL(4,2)",
            "base_price : DECIMAL(10,2)",
            "is_active : BOOLEAN"
        ]
    )

    # 10. WORK_ORDER_ITEM (Col 2)
    e_order_item = draw_unified_entity(
        page, x=X_COL2, top_y=Y_ROW3_TOP, w=EW,
        title="Work_Order_Item (Строки наряда)",
        subtitle="Калькуляция работ и нормо-часов",
        pk_list=["id_order_item : INT"],
        fk_list=[
            "id_order : INT (FK)",
            "id_service : INT (FK)"
        ],
        attr_list=[
            "quantity : SMALLINT",
            "unit_price : DECIMAL(10,2)",
            "discount_pct : DECIMAL(4,2)",
            "item_total : DECIMAL(10,2)"
        ]
    )

    # 11. TUNING_FIRMWARE (Col 3)
    e_tuned = draw_unified_entity(
        page, x=X_COL3, top_y=Y_ROW3_TOP, w=EW,
        title="Tuning_Firmware (Калибровки)",
        subtitle="Модифицированное калибровочное ПО",
        pk_list=["id_tuned : INT"],
        fk_list=[
            "id_stock : INT (FK)"
        ],
        attr_list=[
            "tuning_name : VARCHAR(100)",
            "power_increase_hp : SMALLINT",
            "torque_increase_nm : SMALLINT",
            "disabled_eco_masks : VARCHAR(150)",
            "checksum_crc32 : VARCHAR(10)",
            "tuning_software : VARCHAR(50)",
            "is_tested_verified : BOOLEAN"
        ]
    )

    # 12. WORK_COMPLETION_ACT (Col 4 - Key Output Document)
    e_act = draw_unified_entity(
        page, x=X_COL4, top_y=Y_ROW3_TOP, w=EW,
        title="Work_Completion_Act (Акт)",
        subtitle="Итоговый финансово-отчетный документ",
        pk_list=["id_act : INT"],
        fk_list=[
            "id_order : INT (FK)",
            "id_tuned : INT (FK)",
            "id_master : INT (FK)"
        ],
        attr_list=[
            "act_number : VARCHAR(30)",
            "act_date : DATETIME",
            "warranty_period_months : SMALLINT",
            "total_paid_amount : DECIMAL(10,2)",
            "payment_type : VARCHAR(30)",
            "client_signed_flag : BOOLEAN",
            "master_signed_flag : BOOLEAN",
            "satisfaction_rating : SMALLINT"
        ]
    )

    # ---------------- ROW 4 (LOWER COL 3: DIAGNOSTIC) ----------------
    # 13. DIAGNOSTIC_CARD (Col 3, below Tuning_Firmware)
    # Tuning_Firmware bottom is Y ≈ 5.63 -> Diagnostic_Card top at 5.00 -> bottom at 1.09
    e_diag = draw_unified_entity(
        page, x=X_COL3, top_y=5.00, w=EW,
        title="Diagnostic_Card (Диагностика)",
        subtitle="Карты входного и выходного контроля",
        pk_list=["id_diag : INT"],
        fk_list=[
            "id_order : INT (FK)",
            "id_diagnostician : INT (FK)"
        ],
        attr_list=[
            "diag_stage : VARCHAR(20)",
            "dtc_codes_list : TEXT",
            "fuel_pressure_bar : DECIMAL(6,2)",
            "boost_pressure_bar : DECIMAL(4,2)",
            "knock_retard_deg : DECIMAL(4,2)",
            "diag_conclusion : VARCHAR(30)",
            "diag_timestamp : DATETIME"
        ]
    )

    # ================= RELATIONSHIPS / CONNECTORS (ZERO OVERLAPS WITH INLINE BADGES) =================
    # 1. Client (1) -> Vehicle (N)
    connect_erd_crows_foot(page, (e_client["right"][0], 17.5), (e_vehicle["left"][0], 17.5), 
                           label="1 : N [Владеет]", card_start="1", card_end="N",
                           label_cx=8.3, label_cy=17.5, label_w=2.1)

    # 2. Client (1) -> Service_Request (N)
    connect_erd_crows_foot(page, e_client["bottom"], e_request["top"], 
                           label="1 : N [Подает]", card_start="1", card_end="N",
                           label_cx=4.4, label_cy=15.38, label_w=1.9)

    # 3. Vehicle (1) -> Service_Request (N)
    connect_erd_crows_foot_elbow(page, [(e_vehicle["left"][0], 16.0),
                                        (8.3, 16.0),
                                        (8.3, 13.5),
                                        (e_request["right"][0], 13.5)],
                                 label="1 : N [Указан в заявке]", card_start="1", card_end="N",
                                 label_cx=9.3, label_cy=16.0, label_w=2.4)

    # 4. Service_Request (1) -> Work_Order (1)
    connect_erd_crows_foot(page, (e_request["right"][0], 12.8), (e_order["left"][0], 12.8), 
                           label="1 : 1 [Порождает]", card_start="1", card_end="1", dashed=True,
                           label_cx=8.3, label_cy=12.8, label_w=2.1)

    # 5. Vehicle (1) -> Work_Order (N)
    connect_erd_crows_foot(page, e_vehicle["bottom"], e_order["top"], 
                           label="1 : N [Обслуживается]", card_start="1", card_end="N",
                           label_cx=12.2, label_cy=15.05, label_w=2.4)

    # 6. Work_Order (1) -> Work_Order_Item (N)
    connect_erd_crows_foot(page, e_order["bottom"], e_order_item["top"], 
                           label="1 : N [Содержит строки]", card_start="1", card_end="N",
                           label_cx=12.2, label_cy=9.62, label_w=2.4)

    # 7. Price_List_Service (1) -> Work_Order_Item (N)
    connect_erd_crows_foot(page, (e_price["right"][0], 7.7), (e_order_item["left"][0], 7.7), 
                           label="1 : N [Тарифицирует]", card_start="1", card_end="N",
                           label_cx=8.3, label_cy=7.7, label_w=2.2)

    # 8. ECU_Specification (1) -> Vehicle (N)
    connect_erd_crows_foot(page, (e_ecu["left"][0], 17.5), (e_vehicle["right"][0], 17.5), 
                           label="1 : N [Комплектует]", card_start="1", card_end="N",
                           label_cx=16.1, label_cy=17.5, label_w=2.2)

    # 9. Work_Order (1) -> Stock_Firmware (1)
    connect_erd_crows_foot(page, (e_order["right"][0], 13.5), (e_stock["left"][0], 13.5), 
                           label="1 : 1 [Считывание дампа]", card_start="1", card_end="1",
                           label_cx=16.1, label_cy=13.5, label_w=2.5)

    # 10. Stock_Firmware (1) -> Tuning_Firmware (N)
    connect_erd_crows_foot(page, e_stock["bottom"], e_tuned["top"], 
                           label="1 : N [Модификация ROM]", card_start="1", card_end="N",
                           label_cx=20.0, label_cy=9.94, label_w=2.4)

    # 11. Work_Order (1) -> Diagnostic_Card (N)
    connect_erd_crows_foot_elbow(page, [(e_order["right"][0], 11.2),
                                        (16.8, 11.2),
                                        (16.8, 3.8),
                                        (e_diag["left"][0], 3.8)],
                                 label="1 : N [Вход/Выход ДВС]", card_start="1", card_end="N",
                                 label_cx=15.4, label_cy=11.2, label_w=2.3)

    # 12. Work_Order (1) -> Execution_Log (N)
    connect_erd_crows_foot_elbow(page, [(e_order["right"][0], 10.5),
                                        (15.2, 10.5),
                                        (15.2, 9.60),
                                        (24.5, 9.60),
                                        (24.5, 12.0),
                                        (e_log["left"][0], 12.0)],
                                 label="1 : N [Выполнение + Коммент]", card_start="1", card_end="N",
                                 label_cx=16.5, label_cy=9.60, label_w=2.6)

    # 13. Tuning_Firmware (1) -> Execution_Log (N)
    connect_erd_crows_foot_elbow(page, [(e_tuned["right"][0], 8.0),
                                        (23.8, 8.0),
                                        (23.8, 11.0),
                                        (e_log["left"][0], 11.0)],
                                 label="1 : N [Запись ПО в ЭБУ]", card_start="1", card_end="N",
                                 label_cx=22.8, label_cy=8.0, label_w=2.3)

    # 14. Work_Order (1) -> Work_Completion_Act (1)
    # Highway at Y = 5.20 between Tuning and Diag
    connect_erd_crows_foot_elbow(page, [(e_order["right"][0], 10.2),
                                        (14.6, 10.2),
                                        (14.6, 5.20),
                                        (24.8, 5.20),
                                        (24.8, 7.5),
                                        (e_act["left"][0], 7.5)],
                                 label="1 : 1 [Выдача Акта]", card_start="1", card_end="1", dashed=True,
                                 label_cx=18.0, label_cy=5.20, label_w=2.2)

    # 15. Tuning_Firmware (1) -> Work_Completion_Act (N)
    connect_erd_crows_foot_elbow(page, [(e_tuned["right"][0], 6.5),
                                        (23.2, 6.5),
                                        (23.2, 8.2),
                                        (e_act["left"][0], 8.2)],
                                 label="1 : N [Сертификация]", card_start="1", card_end="N",
                                 label_cx=22.5, label_cy=6.5, label_w=2.2)

    # 16. Employee (1) -> Work_Order (N)
    # Overhead highway at Y = 19.80
    connect_erd_crows_foot_elbow(page, [(e_emp["top"][0], 19.4),
                                        (e_emp["top"][0], 19.80),
                                        (12.8, 19.80),
                                        (12.8, 14.6)],
                                 label="1 : N [Мастер-приемщик]", card_start="1", card_end="N",
                                 label_cx=20.3, label_cy=19.80, label_w=2.4)

    # 17. Employee (1) -> Execution_Log (N)
    connect_erd_crows_foot(page, e_emp["bottom"], e_log["top"], 
                           label="1 : N [Специалист-тюнер]", card_start="1", card_end="N",
                           label_cx=27.8, label_cy=15.54, label_w=2.4)

    # 18. Employee (1) -> Diagnostic_Card (N)
    connect_erd_crows_foot_elbow(page, [(e_emp["left"][0], 17.0),
                                        (25.4, 17.0),
                                        (25.4, 3.0),
                                        (e_diag["right"][0], 3.0)],
                                 label="1 : N [Диагност]", card_start="1", card_end="N",
                                 label_cx=24.2, label_cy=17.0, label_w=1.9)

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
    print("SUCCESS: Standard Unified ERD generated at:")
    print(" - VSDX:", p_vsdx)
    print(" - PNG:", p_png)

try:
    build_erd_unified()
except Exception as e:
    print("Error:", e)
    import traceback
    traceback.print_exc()
finally:
    visio.Quit()

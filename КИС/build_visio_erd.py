import win32com.client
import os

print("=== Generating Perfect Common ER-Model (ERD) in Microsoft Visio ===")

visio = win32com.client.Dispatch("Visio.Application")
visio.Visible = False
visio.AlertResponse = 7

def set_font(shape, size_pt=8.5, bold=False, color="RGB(15, 23, 42)"):
    shape.Cells("Char.Size").FormulaU = f"{size_pt} pt"
    shape.Cells("Char.Color").FormulaU = color
    if bold:
        shape.Cells("Char.Style").FormulaU = "1"
    else:
        shape.Cells("Char.Style").FormulaU = "0"

def draw_entity_table(page, x, y, w, title, subtitle, pk_list, fk_list, attr_list, header_color="RGB(30, 58, 138)"):
    row_h = 0.22
    header_h = 0.58
    body_rows = len(pk_list) + len(fk_list) + len(attr_list)
    body_h = max(1.8, body_rows * row_h + 0.25)
    total_h = header_h + body_h
    
    x1, y1 = x - w/2, y - total_h/2
    x2, y2 = x + w/2, y + total_h/2
    
    # Body rectangle
    body = page.DrawRectangle(x1, y1, x2, y2 - header_h)
    body.Cells("FillForegnd").FormulaU = "RGB(255, 255, 255)"
    body.Cells("LineColor").FormulaU = header_color
    body.Cells("LineWeight").FormulaU = "1.2 pt"
    body.Cells("LeftMargin").FormulaU = "0.08 in"
    body.Cells("RightMargin").FormulaU = "0.06 in"
    body.Cells("TopMargin").FormulaU = "0.06 in"
    body.Cells("BottomMargin").FormulaU = "0.06 in"
    
    # Construct body text
    lines = []
    for pk in pk_list:
        lines.append(f"[PK] {pk}")
    for fk in fk_list:
        lines.append(f"[FK] {fk}")
    for attr in attr_list:
        lines.append(f"     {attr}")
    
    body.Text = "\n".join(lines)
    set_font(body, 7.5, bold=False, color="RGB(30, 41, 59)")
    body.Cells("Para.HorzAlign").FormulaU = "0" # Left align
    body.Cells("VerticalAlign").FormulaU = "0"   # Top align
    
    # Header rectangle
    hdr = page.DrawRectangle(x1, y2 - header_h, x2, y2)
    hdr.Cells("FillForegnd").FormulaU = header_color
    hdr.Cells("LineColor").FormulaU = header_color
    hdr.Cells("LineWeight").FormulaU = "1.2 pt"
    hdr.Cells("LeftMargin").FormulaU = "0.04 in"
    hdr.Cells("RightMargin").FormulaU = "0.04 in"
    hdr.Cells("TopMargin").FormulaU = "0.04 in"
    hdr.Cells("BottomMargin").FormulaU = "0.04 in"
    hdr_text = title
    if subtitle:
        hdr_text += f"\n({subtitle})"
    hdr.Text = hdr_text
    set_font(hdr, 8.5, bold=True, color="RGB(255, 255, 255)")
    hdr.Cells("Para.HorzAlign").FormulaU = "1" # Center
    
    return {
        "x": x, "y": y, "w": w, "h": total_h,
        "x1": x1, "x2": x2, "y1": y1, "y2": y2,
        "top": (x, y2),
        "bottom": (x, y1),
        "left": (x1, y - header_h/2),
        "right": (x2, y - header_h/2)
    }

def connect_erd(page, p1, p2, label="", dashed=False, label_offset_y=0.15, label_offset_x=0.0):
    line = page.DrawLine(p1[0], p1[1], p2[0], p2[1])
    line.Cells("LineWeight").FormulaU = "1.2 pt"
    line.Cells("LineColor").FormulaU = "RGB(51, 65, 85)"
    line.Cells("EndArrow").FormulaU = "13"
    line.Cells("BeginArrow").FormulaU = "0"
    if dashed:
        line.Cells("LinePattern").FormulaU = "2"
        line.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
        
    if label:
        cx = (p1[0] + p2[0])/2 + label_offset_x
        cy = (p1[1] + p2[1])/2 + label_offset_y
        lbl = page.DrawRectangle(cx - 0.8, cy - 0.12, cx + 0.8, cy + 0.12)
        lbl.Text = label
        lbl.Cells("FillPattern").FormulaU = "0" # STRICT REQUIREMENT: TRANSPARENT BACKGROUND
        lbl.Cells("LinePattern").FormulaU = "0"
        set_font(lbl, 7.5, bold=True, color="RGB(15, 23, 42)")
    return line

def connect_erd_elbow(page, points, label="", dashed=False, label_seg=0, label_offset_y=0.15, label_offset_x=0.0):
    for i in range(len(points) - 1):
        p_start = points[i]
        p_end = points[i+1]
        is_last = (i == len(points) - 2)
        line = page.DrawLine(p_start[0], p_start[1], p_end[0], p_end[1])
        line.Cells("LineWeight").FormulaU = "1.2 pt"
        line.Cells("LineColor").FormulaU = "RGB(51, 65, 85)"
        if dashed:
            line.Cells("LinePattern").FormulaU = "2"
            line.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
        if is_last:
            line.Cells("EndArrow").FormulaU = "13"
            
    if label and len(points) > label_seg + 1:
        p_start = points[label_seg]
        p_end = points[label_seg + 1]
        cx = (p_start[0] + p_end[0])/2 + label_offset_x
        cy = (p_start[1] + p_end[1])/2 + label_offset_y
        lbl = page.DrawRectangle(cx - 0.8, cy - 0.12, cx + 0.8, cy + 0.12)
        lbl.Text = label
        lbl.Cells("FillPattern").FormulaU = "0" # TRANSPARENT BACKGROUND
        lbl.Cells("LinePattern").FormulaU = "0"
        set_font(lbl, 7.5, bold=True, color="RGB(15, 23, 42)")

def build_erd():
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)
    
    W, H = 29.0, 19.0
    page.PageSheet.Cells("PageWidth").FormulaU = f"{W} in"
    page.PageSheet.Cells("PageHeight").FormulaU = f"{H} in"
    
    # Top banner
    tb = page.DrawRectangle(0.5, H - 1.1, W - 0.5, H - 0.5)
    tb.Text = "ОБЩАЯ РЕЛЯЦИОННАЯ МОДЕЛЬ ДАННЫХ (ERD) КИС «АВТОСЕРВИС: ЧИП-ТЮНИНГ»\n" \
              "КОНТУР ДАНИЛА (ОФОРМЛЕНИЕ ЗАЯВКИ И НАРЯД-ЗАКАЗ) + КОНТУР КАРИМА (ИСПОЛНЕНИЕ, КАЛИБРОВКА И АКТ)"
    tb.Cells("FillForegnd").FormulaU = "RGB(15, 23, 42)"
    tb.Cells("LineColor").FormulaU = "RGB(15, 23, 42)"
    set_font(tb, 11, bold=True, color="RGB(255, 255, 255)")
    
    # Subdomain background panels
    # Panel 1: Danil's Client & Order Domain
    p1 = page.DrawRectangle(0.8, 1.2, 14.1, H - 1.3)
    p1.Cells("FillForegnd").FormulaU = "RGB(248, 250, 252)"
    p1.Cells("LineColor").FormulaU = "RGB(191, 219, 254)"
    p1.Cells("LineWeight").FormulaU = "1.5 pt"
    
    p1_lbl = page.DrawRectangle(1.0, H - 1.7, 9.5, H - 1.35)
    p1_lbl.Text = "ЗОНА 1: ПОДПРОЦЕСС ДАНИЛА (КЛИЕНТЫ, АВТОМОБИЛИ, ЗАЯВКИ И НАРЯД-ЗАКАЗ)"
    p1_lbl.Cells("FillPattern").FormulaU = "0"
    p1_lbl.Cells("LinePattern").FormulaU = "0"
    set_font(p1_lbl, 8.5, bold=True, color="RGB(30, 58, 138)")
    
    # Panel 2: Karim's Production & Act Domain
    p2 = page.DrawRectangle(14.5, 1.2, W - 0.8, H - 1.3)
    p2.Cells("FillForegnd").FormulaU = "RGB(246, 253, 251)"
    p2.Cells("LineColor").FormulaU = "RGB(153, 246, 228)"
    p2.Cells("LineWeight").FormulaU = "1.5 pt"
    
    p2_lbl = page.DrawRectangle(14.8, H - 1.7, 23.8, H - 1.35)
    p2_lbl.Text = "ЗОНА 2: ПОДПРОЦЕСС КАРИМА (ДИАГНОСТИКА, ПРОШИВКИ, ВЫПОЛНЕНИЕ+КОММЕНТ И АКТ ВЫПОЛНЕННЫХ РАБОТ)"
    p2_lbl.Cells("FillPattern").FormulaU = "0"
    p2_lbl.Cells("LinePattern").FormulaU = "0"
    set_font(p2_lbl, 8.5, bold=True, color="RGB(13, 148, 136)")

    # Bottom Legend
    lg = page.DrawRectangle(0.8, 0.35, W - 0.8, 1.05)
    lg.Cells("FillForegnd").FormulaU = "RGB(255, 255, 255)"
    lg.Cells("LineColor").FormulaU = "RGB(203, 213, 225)"
    lg.Cells("LineWeight").FormulaU = "1 pt"
    lg.Text = "СТАНДАРТ ERD (Третья нормальная форма - 3NF):  [PK] Первичный ключ  |  [FK] Внешний ключ связи\n" \
              "Цвета заголовков: [Синий] Контур Данила (Заказ)  |  [Зеленый] Контур Карима (Калибровка и Акт)  |  [Темно-серый] Справочники  |  Связи: 1 : N, 1 : 1"
    set_font(lg, 8.5, bold=False, color="RGB(51, 65, 85)")

    COL_DANIL = "RGB(30, 58, 138)"     # Deep Blue
    COL_ORDER = "RGB(29, 78, 216)"     # Royal Blue (Work Order)
    COL_KARIM = "RGB(13, 148, 136)"    # Teal/Emerald
    COL_ACT   = "RGB(15, 118, 110)"    # Dark Teal (Work Completion Act)
    COL_COMM  = "RGB(51, 65, 85)"      # Slate (Dictionaries/Employees)

    EW = 2.85 # Entity width

    # ---------------- COLUMN 1: КЛИЕНТЫ И УСЛУГИ (ДАНИЛ) ----------------
    # 1. CLIENT
    e_client = draw_entity_table(
        page, x=4.1, y=14.8, w=EW,
        title="Client (Клиент)",
        subtitle="Автовладелец / Юрлицо",
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
        ],
        header_color=COL_DANIL
    )

    # 2. SERVICE_REQUEST
    e_request = draw_entity_table(
        page, x=4.1, y=9.8, w=EW,
        title="Service_Request (Заявка)",
        subtitle="Первичное обращение",
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
        ],
        header_color=COL_DANIL
    )

    # 3. PRICE_LIST_SERVICE
    e_price = draw_entity_table(
        page, x=4.1, y=4.5, w=EW,
        title="Price_List_Service (Услуга)",
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
        ],
        header_color=COL_COMM
    )

    # ---------------- COLUMN 2: АВТОМОБИЛИ И НАРЯД-ЗАКАЗ ----------------
    # 4. VEHICLE
    e_vehicle = draw_entity_table(
        page, x=10.8, y=14.8, w=EW,
        title="Vehicle (Автомобиль)",
        subtitle="Транспортное средство",
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
        ],
        header_color=COL_DANIL
    )

    # 5. WORK_ORDER - KEY DOCUMENT OF DANIL
    e_order = draw_entity_table(
        page, x=10.8, y=9.8, w=EW + 0.15,
        title="Work_Order (Наряд-заказ)",
        subtitle="КЛЮЧЕВОЙ ДОКУМЕНТ ДАНИЛА",
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
        ],
        header_color=COL_ORDER
    )

    # 6. WORK_ORDER_ITEM
    e_order_item = draw_entity_table(
        page, x=10.8, y=4.5, w=EW,
        title="Work_Order_Item (Строка наряда)",
        subtitle="Спецификация работ",
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
        ],
        header_color=COL_DANIL
    )

    # ---------------- COLUMN 3: БЛОКИ ЭБУ, ПРОШИВКИ И ДИАГНОСТИКА (КАРИМ) ----------------
    # 7. ECU_SPECIFICATION
    e_ecu = draw_entity_table(
        page, x=17.8, y=14.8, w=EW,
        title="ECU_Specification (ЭБУ)",
        subtitle="Справочник контроллеров ДВС",
        pk_list=["id_ecu : INT"],
        fk_list=[],
        attr_list=[
            "ecu_vendor : VARCHAR(50)",
            "ecu_model : VARCHAR(50)",
            "mcu_processor : VARCHAR(50)",
            "flash_capacity_bytes : INT",
            "protocol_type : VARCHAR(50)",
            "connection_mode : VARCHAR(30)"
        ],
        header_color=COL_COMM
    )

    # 8. STOCK_FIRMWARE
    e_stock = draw_entity_table(
        page, x=17.8, y=10.8, w=EW,
        title="Stock_Firmware (Сток-дамп)",
        subtitle="Заводской дамп ROM (Бэкап)",
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
        ],
        header_color=COL_KARIM
    )

    # 9. TUNING_FIRMWARE
    e_tuned = draw_entity_table(
        page, x=17.8, y=6.6, w=EW,
        title="Tuning_Firmware (Калибровка)",
        subtitle="Модифицированное ПО",
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
        ],
        header_color=COL_KARIM
    )

    # 10. DIAGNOSTIC_CARD
    e_diag = draw_entity_table(
        page, x=17.8, y=2.6, w=EW,
        title="Diagnostic_Card (Диагностика)",
        subtitle="Входной / выходной контроль",
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
        ],
        header_color=COL_KARIM
    )

    # ---------------- COLUMN 4: ИСПОЛНЕНИЕ, АКТ И ПЕРСОНАЛ (КАРИМ) ----------------
    # 11. EMPLOYEE
    e_emp = draw_entity_table(
        page, x=24.5, y=14.8, w=EW,
        title="Employee (Сотрудник)",
        subtitle="Штатные роли КИС",
        pk_list=["id_employee : INT"],
        fk_list=[],
        attr_list=[
            "personnel_number : VARCHAR(20)",
            "full_name : VARCHAR(150)",
            "role_title : VARCHAR(50)",
            "phone_number : VARCHAR(20)",
            "access_level : VARCHAR(30)",
            "is_active : BOOLEAN"
        ],
        header_color=COL_COMM
    )

    # 12. EXECUTION_LOG - REQUIREMENT: ВЫПОЛНЕНИЕ + КОММЕНТАРИЙ
    e_log = draw_entity_table(
        page, x=24.5, y=9.8, w=EW + 0.15,
        title="Execution_Log (Журнал работ)",
        subtitle="ВЫПОЛНЕНИЕ + КОММЕНТ (КАРИМ)",
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
        ],
        header_color=COL_KARIM
    )

    # 13. WORK_COMPLETION_ACT - KEY DOCUMENT OF KARIM
    e_act = draw_entity_table(
        page, x=24.5, y=4.5, w=EW + 0.15,
        title="Work_Completion_Act (Акт)",
        subtitle="КЛЮЧЕВОЙ ДОКУМЕНТ КАРИМА",
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
        ],
        header_color=COL_ACT
    )

    # ================= RELATIONSHIPS / CONNECTORS =================
    # 1. Client (1) -> Vehicle (M)
    connect_erd(page, e_client["right"], e_vehicle["left"], label="1 : N [Владеет]")

    # 2. Client (1) -> Service_Request (M)
    connect_erd(page, e_client["bottom"], e_request["top"], label="1 : N [Подает]")

    # 3. Vehicle (1) -> Service_Request (M)
    connect_erd_elbow(page, [(e_vehicle["left"][0], e_vehicle["bottom"][1] + 0.3), 
                             (7.3, e_vehicle["bottom"][1] + 0.3), 
                             (7.3, e_request["right"][1]),
                             (e_request["right"][0], e_request["right"][1])], 
                      label="1 : N", label_seg=0, label_offset_y=0.15)

    # 4. Service_Request (1) -> Work_Order (1)
    connect_erd(page, e_request["right"], e_order["left"], label="1 : 1 [Порождает]", dashed=True)

    # 5. Vehicle (1) -> Work_Order (M)
    connect_erd(page, e_vehicle["bottom"], e_order["top"], label="1 : N [Обслуживается]")

    # 6. Work_Order (1) -> Work_Order_Item (M)
    connect_erd(page, e_order["bottom"], e_order_item["top"], label="1 : N [Содержит]")

    # 7. Price_List_Service (1) -> Work_Order_Item (M)
    connect_erd(page, e_price["right"], e_order_item["left"], label="1 : N [Тарифицирует]")

    # 8. Vehicle (M) -> ECU_Specification (1)
    connect_erd(page, e_vehicle["right"], e_ecu["left"], label="N : 1 [Оснащен]")

    # 9. Work_Order (1) -> Stock_Firmware (1)
    connect_erd_elbow(page, [(e_order["right"][0], e_order["top"][1] - 0.2),
                             (14.3, e_order["top"][1] - 0.2),
                             (14.3, e_stock["left"][1] + 0.2),
                             (e_stock["left"][0], e_stock["left"][1] + 0.2)],
                      label="1 : 1 [Считывание]", label_seg=0, label_offset_y=0.15)

    # 10. Stock_Firmware (1) -> Tuning_Firmware (M)
    connect_erd(page, e_stock["bottom"], e_tuned["top"], label="1 : N [Модификация]")

    # 11. Work_Order (1) -> Diagnostic_Card (M)
    connect_erd_elbow(page, [(e_order["right"][0], e_order["bottom"][1] + 0.3),
                             (14.3, e_order["bottom"][1] + 0.3),
                             (14.3, e_diag["left"][1]),
                             (e_diag["left"][0], e_diag["left"][1])],
                      label="1 : N [Вход/Выход]", label_seg=0, label_offset_y=0.15)

    # 12. Work_Order (1) -> Execution_Log (1..N) (ВЫПОЛНЕНИЕ + КОММЕНТАРИЙ)
    connect_erd_elbow(page, [(e_order["right"][0], e_order["right"][1]),
                             (16.0, e_order["right"][1]),
                             (16.0, 8.6),
                             (22.4, 8.6),
                             (22.4, e_log["left"][1] + 0.2),
                             (e_log["left"][0], e_log["left"][1] + 0.2)],
                      label="1 : N [Выполнение+Коммент]", label_seg=2, label_offset_y=0.15)

    # 13. Tuning_Firmware (1) -> Execution_Log (M)
    connect_erd(page, e_tuned["right"], (e_log["left"][0], e_log["bottom"][1] + 0.4), label="1 : N [Запись]")

    # 14. Work_Order (1) -> Work_Completion_Act (1) (ГЛАВНЫЙ ДОКУМЕНТ КАРИМА)
    # Clean routing underneath Diagnostic_Card at Y = 1.05
    connect_erd_elbow(page, [(e_order["right"][0], e_order["bottom"][1] + 0.1),
                             (14.0, e_order["bottom"][1] + 0.1),
                             (14.0, 1.05),
                             (24.5, 1.05),
                             (e_act["bottom"][0], e_act["bottom"][1])],
                      label="1 : 1 [Выдача Акта]", label_seg=2, label_offset_y=0.15, dashed=True)

    # 15. Tuning_Firmware (1) -> Work_Completion_Act (M)
    connect_erd_elbow(page, [(e_tuned["right"][0], e_tuned["bottom"][1] + 0.2),
                             (21.8, e_tuned["bottom"][1] + 0.2),
                             (21.8, e_act["top"][1] - 0.2),
                             (e_act["left"][0], e_act["top"][1] - 0.2)],
                      label="1 : N [Сертификат]", label_seg=0, label_offset_y=0.15)

    # 16. Employee (1) -> Work_Order (M)
    connect_erd_elbow(page, [(e_emp["left"][0], e_emp["left"][1] - 0.2),
                             (21.0, e_emp["left"][1] - 0.2),
                             (21.0, 17.0),
                             (12.5, 17.0),
                             (12.5, e_order["top"][1]),
                             (e_order["top"][0] + 0.4, e_order["top"][1])],
                      label="1 : N [Оформляет]", label_seg=2, label_offset_y=0.15)

    # 17. Employee (1) -> Execution_Log (M)
    connect_erd(page, e_emp["bottom"], e_log["top"], label="1 : N [Тюнер]")

    # 18. Employee (1) -> Diagnostic_Card (M)
    connect_erd_elbow(page, [(e_emp["left"][0], e_emp["bottom"][1] + 0.2),
                             (20.8, e_emp["bottom"][1] + 0.2),
                             (20.8, 2.6),
                             (e_diag["right"][0], 2.6)],
                      label="1 : N [Диагност]", label_seg=1, label_offset_x=0.5)

    p_vsdx = os.path.abspath(r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ER_модель_КИС_ЧипТюнинг.vsdx")
    p_png = os.path.abspath(r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ER_модель_КИС_ЧипТюнинг.png")
    if os.path.exists(p_vsdx): os.remove(p_vsdx)
    if os.path.exists(p_png): os.remove(p_png)
    doc.SaveAs(p_vsdx)
    page.Export(p_png)
    doc.Saved = True
    doc.Close()
    print("ERD generated successfully!")

try:
    build_erd()
except Exception as e:
    print("Error:", e)
    import traceback
    traceback.print_exc()
finally:
    visio.Quit()

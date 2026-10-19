import win32com.client
import os

print("=== Starting Improved Visio BPMN 2.0 Dynamic Model Generator ===")

visio = win32com.client.Dispatch("Visio.Application")
visio.Visible = False
visio.AlertResponse = 7 # Do not block on alerts

try:
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)

    # Page size: 22.0 x 13.5 inches landscape for clean, beautiful layout
    page.PageSheet.Cells("PageWidth").FormulaU = "22.0 in"
    page.PageSheet.Cells("PageHeight").FormulaU = "13.5 in"

    # Helper styling function
    def set_font(shape, size_pt=9, bold=False, color="RGB(40,40,40)"):
        shape.Cells("Char.Size").FormulaU = f"{size_pt} pt"
        shape.Cells("Char.Color").FormulaU = color
        if bold:
            shape.Cells("Char.Style").FormulaU = "1"

    # 1. Top Title Banner
    title_banner = page.DrawRectangle(0.5, 12.4, 21.5, 13.0)
    title_banner.Text = "ДИСПЕТЧЕРИЗАЦИЯ И ДИНАМИЧЕСКАЯ МОДЕЛЬ ВЫПОЛНЕНИЯ ЧИП-ТЮНИНГА АВТОМОБИЛЯ (BPMN 2.0)"
    title_banner.Cells("FillForegnd").FormulaU = "RGB(15, 23, 42)" # Deep Navy
    title_banner.Cells("LineColor").FormulaU = "RGB(15, 23, 42)"
    set_font(title_banner, 13, bold=True, color="RGB(255,255,255)")

    # 2. Main Pool Vertical Header on the left
    pool_header = page.DrawRectangle(0.4, 1.2, 1.5, 12.2)
    pool_header.Text = "ПУЛ: ЧИП-ТЮНИНГ И КАЛИБРОВКА ЭБУ (BPMN 2.0)"
    pool_header.Cells("FillForegnd").FormulaU = "RGB(30, 41, 59)" # Dark slate
    pool_header.Cells("LineColor").FormulaU = "RGB(15, 23, 42)"
    pool_header.Cells("LineWeight").FormulaU = "1.5 pt"
    pool_header.Cells("LeftMargin").FormulaU = "0.02 in"
    pool_header.Cells("RightMargin").FormulaU = "0.02 in"
    pool_header.Cells("TopMargin").FormulaU = "0.02 in"
    pool_header.Cells("BottomMargin").FormulaU = "0.02 in"
    set_font(pool_header, 10.5, bold=True, color="RGB(255,255,255)")
    pool_header.Cells("TxtAngle").FormulaU = "90 deg"

    # 3. Lanes definition: (role_title, y_bottom, y_top, bg_color)
    lanes = [
        ("Мастер-приемщик\n(Сервисный консультант)", 9.5, 12.2, "RGB(255, 251, 240)"),
        ("Инженер-диагност\n(Входной/выходной контроль)", 6.7, 9.5, "RGB(242, 253, 245)"),
        ("Специалист по чип-тюнингу\n(Калибровщик ЭБУ)", 3.8, 6.7, "RGB(240, 247, 255)"),
        ("КИС «Автосервис: Тюнинг»\n(Модуль калибровки и БД)", 1.2, 3.8, "RGB(252, 246, 255)")
    ]

    for role_name, yb, yt, color in lanes:
        lane = page.DrawRectangle(1.5, yb, 21.5, yt)
        lane.Cells("FillForegnd").FormulaU = color
        lane.Cells("LineColor").FormulaU = "RGB(203, 213, 225)"
        lane.Cells("LineWeight").FormulaU = "1 pt"

        rh = page.DrawRectangle(1.5, yb, 3.1, yt)
        rh.Text = role_name
        rh.Cells("FillForegnd").FormulaU = "RGB(241, 245, 249)"
        rh.Cells("LineColor").FormulaU = "RGB(148, 163, 184)"
        rh.Cells("LineWeight").FormulaU = "1.2 pt"
        set_font(rh, 8.5, bold=True, color="RGB(51, 65, 85)")

    # 4. Drawing helper functions
    def draw_task(x, y, w, h, text, task_type="User"):
        shape = page.DrawRectangle(x - w/2, y - h/2, x + w/2, y + h/2)
        shape.Cells("Rounding").FormulaU = "0.12 in"
        shape.Cells("LineWeight").FormulaU = "1.2 pt"
        
        if task_type == "User":
            shape.Cells("FillForegnd").FormulaU = "RGB(255, 255, 255)"
            shape.Cells("LineColor").FormulaU = "RGB(37, 99, 235)"
            prefix = "Оператор КИС\n"
        elif task_type == "Service":
            shape.Cells("FillForegnd").FormulaU = "RGB(245, 243, 255)"
            shape.Cells("LineColor").FormulaU = "RGB(124, 58, 237)"
            prefix = "Сервис КИС\n"
        elif task_type == "Manual":
            shape.Cells("FillForegnd").FormulaU = "RGB(255, 255, 255)"
            shape.Cells("LineColor").FormulaU = "RGB(217, 119, 6)"
            prefix = "Ручной пост\n"
        elif task_type == "Exception":
            shape.Cells("FillForegnd").FormulaU = "RGB(254, 242, 242)"
            shape.Cells("LineColor").FormulaU = "RGB(220, 38, 38)"
            prefix = "Сбой / Отказ\n"

        shape.Cells("LeftMargin").FormulaU = "0.05 in"
        shape.Cells("RightMargin").FormulaU = "0.05 in"
        shape.Cells("TopMargin").FormulaU = "0.05 in"
        shape.Cells("BottomMargin").FormulaU = "0.05 in"
        shape.Text = prefix + text
        set_font(shape, 8, bold=False, color="RGB(30, 41, 59)")
        return shape

    def draw_event(x, y, r, text, ev_type="Start"):
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
        lbl.Cells("FillPattern").FormulaU = "0"
        lbl.Cells("LinePattern").FormulaU = "0"
        set_font(lbl, 7.5, bold=True, color="RGB(30, 41, 59)")
        return shape

    def draw_gateway(x, y, size, question="", q_pos="above"):
        xy = [x - size, y,  x, y + size,  x + size, y,  x, y - size,  x - size, y]
        shape = page.DrawPolyline(xy, 0)
        shape.Cells("FillForegnd").FormulaU = "RGB(254, 249, 195)"
        shape.Cells("LineColor").FormulaU = "RGB(202, 138, 4)"
        shape.Cells("LineWeight").FormulaU = "1.5 pt"
        shape.Text = "✕"
        set_font(shape, 12, bold=True, color="RGB(161, 98, 7)")

        if question:
            if q_pos == "above":
                qlbl = page.DrawRectangle(x - 1.0, y + size + 0.08, x + 1.0, y + size + 0.48)
            else:
                qlbl = page.DrawRectangle(x - 1.0, y - size - 0.48, x + 1.0, y - size - 0.08)
            qlbl.Text = question
            qlbl.Cells("FillPattern").FormulaU = "0"
            qlbl.Cells("LinePattern").FormulaU = "0"
            set_font(qlbl, 8, bold=True, color="RGB(161, 98, 7)")
        return shape

    def draw_datastore(x, y, w, h, text):
        shape = page.DrawRectangle(x - w/2, y - h/2, x + w/2, y + h/2)
        shape.Cells("FillForegnd").FormulaU = "RGB(241, 245, 249)"
        shape.Cells("LineColor").FormulaU = "RGB(71, 85, 105)"
        shape.Cells("LineWeight").FormulaU = "1.2 pt"
        shape.Text = "База данных:\n" + text
        set_font(shape, 8, bold=True, color="RGB(30, 41, 59)")
        return shape

    def connect(x1, y1, x2, y2, label="", dashed=False, label_offset_y=0.15):
        line = page.DrawLine(x1, y1, x2, y2)
        line.Cells("EndArrow").FormulaU = "13"
        line.Cells("LineWeight").FormulaU = "1.2 pt"
        line.Cells("LineColor").FormulaU = "RGB(51, 65, 85)"
        if dashed:
            line.Cells("LinePattern").FormulaU = "2"
            line.Cells("LineColor").FormulaU = "RGB(100, 116, 139)"

        if label:
            cx = (x1 + x2) / 2
            cy = (y1 + y2) / 2 + label_offset_y
            lbl = page.DrawRectangle(cx - 0.7, cy - 0.12, cx + 0.7, cy + 0.12)
            lbl.Text = label
            lbl.Cells("FillForegnd").FormulaU = "RGB(255, 255, 255)"
            lbl.Cells("LinePattern").FormulaU = "0"
            set_font(lbl, 7.5, bold=True, color="RGB(71, 85, 105)")
        return line

    def connect_elbow(points, label="", dashed=False, label_seg=0, label_offset_y=0.15):
        for i in range(len(points) - 1):
            p1 = points[i]
            p2 = points[i+1]
            is_last = (i == len(points) - 2)
            line = page.DrawLine(p1[0], p1[1], p2[0], p2[1])
            line.Cells("LineWeight").FormulaU = "1.2 pt"
            line.Cells("LineColor").FormulaU = "RGB(51, 65, 85)"
            if dashed:
                line.Cells("LinePattern").FormulaU = "2"
                line.Cells("LineColor").FormulaU = "RGB(100, 116, 139)"
            if is_last:
                line.Cells("EndArrow").FormulaU = "13"
        if label:
            p1, p2 = points[label_seg], points[label_seg+1]
            cx, cy = (p1[0] + p2[0])/2, (p1[1] + p2[1])/2 + label_offset_y
            lbl = page.DrawRectangle(cx - 0.7, cy - 0.12, cx + 0.7, cy + 0.12)
            lbl.Text = label
            lbl.Cells("FillForegnd").FormulaU = "RGB(255, 255, 255)"
            lbl.Cells("LinePattern").FormulaU = "0"
            set_font(lbl, 7.5, bold=True, color="RGB(71, 85, 105)")

    print("Placing BPMN Elements...")

    # Task box dimensions: width = 1.95, height = 0.90
    TW, TH = 1.95, 0.90

    # --- LANE 1: МАСТЕР-ПРИЕМЩИК (Y_center = 10.85) ---
    start_ev = draw_event(3.7, 10.85, 0.25, "Старт: Заказ-наряд\nпередан в цех", "Start")
    task_1_1 = draw_task(5.0, 10.85, TW, TH, "1.1 Регистрация авто\nв КИС и передача\nна пост тюнинга", "User")
    connect(3.95, 10.85, 4.025, 10.85)

    task_1_2 = draw_task(8.0, 10.85, TW, TH, "1.2 Оформление акта\nдефектовки в КИС\nи отказ клиенту", "Exception")
    end_fail = draw_event(9.7, 10.85, 0.25, "Конец: Отказ в чип-тюнинге\n(ДВС неисправен)", "EndFail")
    connect(8.975, 10.85, 9.45, 10.85)

    task_1_3 = draw_task(19.3, 10.85, TW, TH, "1.3 Закрытие наряда\nв КИС, выдача авто\nи гарантийного талона", "User")
    end_ok = draw_event(20.9, 10.85, 0.25, "Успешный финал:\nАвтомобиль сдан\nклиенту (Stage 1)", "EndSuccess")
    connect(20.275, 10.85, 20.65, 10.85)

    # --- LANE 2: ИНЖЕНЕР-ДИАГНОСТ (Y_center = 8.1) ---
    task_2_1 = draw_task(5.0, 8.1, TW, TH, "2.1 Входная компьютерная\nдиагностика ЭСУД\n(чтение DTC и датчиков)", "Manual")
    connect(5.0, 10.4, 5.0, 8.55) # From 1.1 down to 2.1

    gw_1 = draw_gateway(6.6, 8.1, 0.32, "Критические\nошибки DTC?", q_pos="above")
    connect(5.975, 8.1, 6.28, 8.1)

    # Branch YES -> Up to 1.2
    connect_elbow([(6.6, 8.42), (6.6, 10.85), (7.025, 10.85)], label="Есть ошибки [Да]")

    # Branch NO -> Right to 2.2
    task_2_2 = draw_task(8.2, 8.1, TW, TH, "2.2 Оформление\nэлектронного допуска\nк прошивке в КИС", "User")
    connect(6.92, 8.1, 7.225, 8.1, label="Норма [Нет]")

    # Exit check
    task_2_3 = draw_task(17.2, 8.1, TW, TH, "2.3 Выходной тест-драйв\nи динамическое снятие\nлогов (наддув/детонация)", "Manual")

    gw_4 = draw_gateway(18.7, 8.1, 0.32, "Логи в норме?", q_pos="above")
    connect(18.175, 8.1, 18.38, 8.1)

    # --- LANE 3: СПЕЦИАЛИСТ ПО ЧИП-ТЮНИНГУ (Y_center = 5.25) ---
    task_3_1 = draw_task(8.2, 5.25, TW, TH, "3.1 Подключение\nстабилизатора АКБ (14V)\nи программатора", "Manual")
    connect(8.2, 7.65, 8.2, 5.7) # From 2.2 down to 3.1

    task_3_2 = draw_task(10.2, 5.25, TW, TH, "3.2 Идентификация\nЭБУ и вычитывание\nсток-дампа (ROM)", "User")
    connect(9.175, 5.25, 9.225, 5.25)

    gw_2 = draw_gateway(11.9, 5.25, 0.32, "Выбор варианта\nкалибровки?", q_pos="above")

    task_3_3 = draw_task(13.8, 6.1, TW, 0.8, "3.3 Индивидуальная\nкалибровка карт\nв WinOLS / Titanium", "User")
    # Connect gw_2 to 3.3: horizontal segment gets the label
    connect_elbow([(11.9, 5.57), (11.9, 6.1), (12.825, 6.1)], label="Кастом [ТЗ]", label_seg=1)

    task_3_4 = draw_task(15.5, 5.25, TW, TH, "3.4 Пересчет CRC\nи запись прошивки\nво flash-память ЭБУ", "User")
    connect_elbow([(14.775, 6.1), (15.5, 6.1), (15.5, 5.7)]) # From 3.3 to 3.4

    gw_3 = draw_gateway(17.0, 5.25, 0.32, "Запись OK?", q_pos="below")
    connect(16.475, 5.25, 16.68, 5.25)

    task_3_5 = draw_task(17.0, 6.35, TW, 0.75, "3.5 Recovery Mode:\nАварийный откат\nна сток (Bench mode)", "Exception")
    connect(17.0, 5.57, 17.0, 5.975, label="Сбой [Ошибка]")
    connect_elbow([(17.0, 6.725), (17.0, 6.95), (15.5, 6.95), (15.5, 5.7)], label="Повтор", label_seg=1)

    task_3_6 = draw_task(18.7, 5.25, TW, TH, "3.6 Сброс адаптаций\nЭБУ и очистка\nсервисных кодов DTC", "User")
    connect(17.32, 5.25, 17.725, 5.25, label="Успех")

    connect_elbow([(18.7, 5.7), (18.7, 7.3), (17.2, 7.3), (17.2, 7.65)]) # From 3.6 to 2.3

    # GW4 branches:
    connect_elbow([(18.7, 8.42), (18.7, 9.1), (13.8, 9.1), (13.8, 6.5)], label="Отклонение [Доработка]")

    # --- LANE 4: КИС «АВТОСЕРВИС» (Y_center = 2.5) ---
    task_4_1 = draw_task(10.2, 2.5, TW, TH, "4.1 Автосохранение\nстока в базе КИС\n(хэш SHA-256)", "Service")
    connect(10.2, 4.8, 10.2, 2.95) # From 3.2 down to 4.1

    db_stock = draw_datastore(11.9, 2.5, 1.5, 0.75, "Хранилище дампов\n(Stock ROM DB)")
    connect(11.175, 2.5, 11.15, 2.5, dashed=True)
    connect(11.9, 2.88, 11.9, 4.93) # From DB to gw_2

    task_4_2 = draw_task(13.8, 2.5, TW, TH, "4.2 Выгрузка готовой\nпрошивки Stage 1\nиз репозитория КИС", "Service")
    connect_elbow([(11.9, 4.93), (11.9, 3.8), (12.825, 3.8), (12.825, 2.5)], label="Из базы")
    connect_elbow([(14.775, 2.5), (15.5, 2.5), (15.5, 4.8)]) # From 4.2 up to 3.4

    task_4_3 = draw_task(19.3, 2.5, TW, TH, "4.3 Архивирование тюнинг-\nфайла, расчет стоимости,\nпечать сертификата", "Service")
    connect_elbow([(19.02, 8.1), (19.3, 8.1), (19.3, 2.95)], label="Норма [Да]")

    connect(19.3, 2.95, 19.3, 10.4) # From 4.3 up to 1.3

    # Bottom Legend
    legend = page.DrawRectangle(3.1, 0.35, 21.5, 1.05)
    legend.Cells("FillForegnd").FormulaU = "RGB(248, 250, 252)"
    legend.Cells("LineColor").FormulaU = "RGB(203, 213, 225)"
    legend.Cells("LineWeight").FormulaU = "1 pt"
    legend.Text = "УСЛОВНЫЕ ОБОЗНАЧЕНИЯ BPMN 2.0:\n" \
                  "(O) Зеленый круг - Начальное событие | (O) Двойной красный круг - Конечное событие | <X> Ромб - Исключающий шлюз XOR\n" \
                  "Оператор КИС - Задача оператора в КИС | Сервис КИС - Автоматический сервис системы | Ручной пост - Физическая операция | Сбой / Отказ - Обработка ошибки"
    set_font(legend, 8.5, bold=False, color="RGB(51, 65, 85)")

    # Save to VSDX
    save_path = os.path.abspath(r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\Динамическая_модель_чип_тюнинг_BPMN.vsdx")
    if os.path.exists(save_path):
        try:
            os.remove(save_path)
        except Exception:
            pass

    print(f"Saving to {save_path}...")
    doc.SaveAs(save_path)

    # Export to PNG
    png_path = os.path.abspath(r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\Динамическая_модель_чип_тюнинг_BPMN.png")
    if os.path.exists(png_path):
        try:
            os.remove(png_path)
        except Exception:
            pass
    print(f"Exporting to {png_path}...")
    page.Export(png_path)

    doc.Saved = True
    doc.Close()
    print("SUCCESS! Beautiful BPMN dynamic model and PNG generated.")

except Exception as e:
    print("Error during Visio generation:", e)
    import traceback
    traceback.print_exc()

finally:
    try:
        visio.Quit()
    except:
        pass

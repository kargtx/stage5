import win32com.client
import os

print("Starting Visio COM application for v2...")
visio = win32com.client.Dispatch("Visio.Application")
visio.Visible = False

try:
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)

    # Set page size to Ledger (17 x 11 inches) landscape to fit the expanded diagram
    page.PageSheet.Cells("PageWidth").FormulaU = "17 in"
    page.PageSheet.Cells("PageHeight").FormulaU = "11 in"

    print("Drawing enterprise boundary...")
    boundary = page.DrawRectangle(3.0, 1.0, 16.0, 10.0)
    boundary.Text = "КОНТУР ПРЕДПРИЯТИЯ: АВТОСЕРВИС (РАСШИРЕННЫЙ СЦЕНАРИЙ)"
    boundary.Cells("FillPattern").FormulaU = "0"
    boundary.Cells("LinePattern").FormulaU = "2" # Dashed
    boundary.Cells("LineWeight").FormulaU = "2 pt"
    boundary.Cells("VerticalAlign").FormulaU = "0" # Top align
    boundary.Cells("Char.Size").FormulaU = "16 pt"

    # Helper functions
    def draw_role(x, y, w, h, text, color):
        shape = page.DrawRectangle(x - w/2, y - h/2, x + w/2, y + h/2)
        shape.Text = text
        shape.Cells("FillForegnd").FormulaU = f"RGB({color})"
        shape.Cells("LineWeight").FormulaU = "1.5 pt"
        shape.Cells("Rounding").FormulaU = "0.3 in" # Rounded corners for a modern look
        shape.Cells("Char.Size").FormulaU = "12 pt"
        return shape

    def draw_doc(x, y, text):
        shape = page.DrawRectangle(x - 0.9, y - 0.4, x + 0.9, y + 0.4)
        shape.Text = text
        shape.Cells("FillForegnd").FormulaU = "RGB(255,255,255)"
        shape.Cells("LineColor").FormulaU = "RGB(0,0,150)"
        shape.Cells("LineWeight").FormulaU = "1 pt"
        shape.Cells("Char.Size").FormulaU = "10 pt"
        return shape

    def draw_line(x1, y1, x2, y2, text="", text_y_offset=0.2):
        line = page.DrawLine(x1, y1, x2, y2)
        line.Cells("EndArrow").FormulaU = "13"
        line.Cells("LineWeight").FormulaU = "1.2 pt"
        if text:
            # Add a text label above the line
            cx = (x1 + x2) / 2
            cy = ((y1 + y2) / 2) + text_y_offset
            lbl = page.DrawRectangle(cx - 1, cy - 0.2, cx + 1, cy + 0.2)
            lbl.Text = text
            lbl.Cells("LinePattern").FormulaU = "0"
            lbl.Cells("FillPattern").FormulaU = "0"
            lbl.Cells("Char.Size").FormulaU = "10 pt"
            lbl.Cells("Char.Color").FormulaU = "RGB(100,100,100)"
        return line

    print("Drawing roles...")
    # Outside
    client = draw_role(1.5, 5.5, 1.8, 2.0, "👤 Клиент", "200, 230, 255")

    # Inside
    manager = draw_role(5.0, 5.5, 2.0, 1.5, "👨‍💼 Мастер-приемщик", "255, 230, 153")
    diagnost = draw_role(9.5, 8.5, 2.2, 1.5, "👨‍⚕️ Диагност", "200, 255, 200")
    tuner = draw_role(9.5, 5.5, 2.2, 1.5, "👨‍💻 Программист-тюнер", "200, 240, 255")
    cashier = draw_role(9.5, 2.5, 2.2, 1.5, "👩‍💼 Кассир / Бухгалтерия", "240, 200, 255")

    print("Drawing flows and documents...")
    
    # 1. Client to Manager (Request)
    draw_line(2.4, 6.0, 4.0, 6.0)
    draw_doc(3.2, 6.0, "📝 Заявка +\nТехпаспорт (СТС)")

    # 2. Manager to Diagnostician
    draw_line(5.0, 6.25, 5.0, 8.5)
    draw_line(5.0, 8.5, 8.4, 8.5)
    draw_doc(6.7, 8.5, "📄 Наряд на\nдиагностику")

    # 3. Diagnostician to Manager
    draw_line(8.4, 8.0, 6.0, 8.0)
    draw_line(6.0, 8.0, 6.0, 6.25)
    draw_doc(7.2, 8.0, "✅ Лист диагностики\n(Допуск к прошивке)")

    # 4. Manager to Client (Contract)
    draw_line(4.0, 5.6, 2.4, 5.6)
    draw_doc(3.2, 5.6, "🤝 Договор на\nоказание услуг")

    # 5. Manager to Tuner
    draw_line(6.0, 5.8, 8.4, 5.8)
    draw_doc(7.2, 5.8, "⚙️ Задание на\nпрошивку (ТЗ)")

    # 6. Tuner internal action (Backup, Modify, Write)
    spec_act = page.DrawRectangle(12.0, 4.8, 14.5, 6.2)
    spec_act.Text = "Действия:\n1. Бэкап стока\n2. Редактирование ПО\n3. Запись Stage 1\n4. Тестовый запуск"
    spec_act.Cells("LinePattern").FormulaU = "2" # dashed
    spec_act.Cells("FillPattern").FormulaU = "0"
    spec_act.Cells("Char.Size").FormulaU = "11 pt"
    draw_line(10.6, 5.5, 12.0, 5.5, "Выполнение")

    # 7. Tuner to Manager
    draw_line(8.4, 5.2, 6.0, 5.2)
    draw_doc(7.2, 5.2, "📊 Отчет (Лог)\nо прошивке")

    # 8. Manager to Cashier
    draw_line(5.0, 4.75, 5.0, 2.5)
    draw_line(5.0, 2.5, 8.4, 2.5)
    draw_doc(6.7, 2.5, "🧾 Счет на\nоплату")

    # 9. Client to Cashier
    draw_line(1.5, 4.5, 1.5, 1.5)
    draw_line(1.5, 1.5, 9.5, 1.5)
    draw_line(9.5, 1.5, 9.5, 1.75)
    draw_doc(5.5, 1.5, "💵 Оплата (Чек)")

    # 10. Manager to Client (Final)
    draw_line(4.0, 5.0, 2.4, 5.0)
    draw_doc(3.2, 5.0, "📑 Акт выполненных\nработ + Гарантия")

    save_path = os.path.abspath("Мнемосхема_ЧипТюнинг_Расширенная.vsdx")
    if os.path.exists(save_path):
        try:
            os.remove(save_path)
        except:
            pass
    print(f"Saving to {save_path}...")
    doc.SaveAs(save_path)
    doc.Close()
    print("Success!")
except Exception as e:
    print(f"Error: {e}")
finally:
    visio.Quit()

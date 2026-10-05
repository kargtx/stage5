import win32com.client
import os

print("Starting Visio COM application...")
visio = win32com.client.Dispatch("Visio.Application")
visio.Visible = False

try:
    doc = visio.Documents.Add("")
    page = doc.Pages.Item(1)

    print("Drawing enterprise boundary...")
    # x1, y1, x2, y2 (in inches, default page is 8.5 x 11)
    boundary = page.DrawRectangle(2.5, 3.5, 8.0, 8.5)
    boundary.Text = "Контур предприятия: Автосервис"
    boundary.Cells("FillPattern").FormulaU = "0" # Transparent
    boundary.Cells("LinePattern").FormulaU = "2" # Dashed line
    boundary.Cells("LineWeight").FormulaU = "1.5 pt"
    boundary.Cells("VerticalAlign").FormulaU = "0" # Top align text

    print("Drawing roles...")
    # Client (outside)
    client = page.DrawRectangle(0.5, 6.0, 1.8, 7.0)
    client.Text = "Клиент\n(Внешняя среда)"

    # Manager (inside)
    manager = page.DrawRectangle(3.0, 6.0, 4.5, 7.0)
    manager.Text = "Менеджер"

    # Specialist (inside)
    specialist = page.DrawRectangle(6.0, 6.0, 7.5, 7.0)
    specialist.Text = "Специалист\nпо чип-тюнингу"

    print("Drawing flows and documents...")
    # Client -> Manager
    line1 = page.DrawLine(1.8, 6.7, 3.0, 6.7)
    line1.Cells("EndArrow").FormulaU = "13" 
    doc_req = page.DrawRectangle(1.9, 6.8, 2.9, 7.2)
    doc_req.Text = "Заявка"
    doc_req.Cells("LinePattern").FormulaU = "0"
    doc_req.Cells("FillPattern").FormulaU = "0"

    # Manager -> Specialist
    line2 = page.DrawLine(4.5, 6.7, 6.0, 6.7)
    line2.Cells("EndArrow").FormulaU = "13"
    doc_order = page.DrawRectangle(4.7, 6.8, 5.8, 7.2)
    doc_order.Text = "Наряд-заказ"
    doc_order.Cells("LinePattern").FormulaU = "0"
    doc_order.Cells("FillPattern").FormulaU = "0"

    # Specialist -> Manager (Return)
    line3 = page.DrawLine(6.0, 6.3, 4.5, 6.3)
    line3.Cells("EndArrow").FormulaU = "13"
    doc_report = page.DrawRectangle(4.7, 5.9, 5.8, 6.2)
    doc_report.Text = "Отчет +\nкомментарий"
    doc_report.Cells("LinePattern").FormulaU = "0"
    doc_report.Cells("FillPattern").FormulaU = "0"

    # Manager -> Client (Act)
    line4 = page.DrawLine(3.0, 6.3, 1.8, 6.3)
    line4.Cells("EndArrow").FormulaU = "13"
    doc_act = page.DrawRectangle(1.9, 5.8, 2.9, 6.2)
    doc_act.Text = "Акт выполненных\nработ"
    doc_act.Cells("LinePattern").FormulaU = "0"
    doc_act.Cells("FillPattern").FormulaU = "0"

    # Specialist internal action
    spec_act = page.DrawRectangle(6.0, 4.5, 7.5, 5.5)
    spec_act.Text = "Выполнение работ\n+ Комментарий"
    spec_act.Cells("LinePattern").FormulaU = "2" # dashed
    spec_act.Cells("FillPattern").FormulaU = "0"
    line5 = page.DrawLine(6.75, 6.0, 6.75, 5.5)
    line5.Cells("EndArrow").FormulaU = "13"

    save_path = os.path.abspath("Мнемосхема_ЧипТюнинг.vsdx")
    if os.path.exists(save_path):
        os.remove(save_path)
    print(f"Saving to {save_path}...")
    doc.SaveAs(save_path)
    doc.Close()
    print("Success!")
except Exception as e:
    print(f"Error: {e}")
finally:
    visio.Quit()

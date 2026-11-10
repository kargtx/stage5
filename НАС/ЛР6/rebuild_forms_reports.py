import os
import time
import win32com.client

DB_PATH = r"C:\Users\danch\OneDrive\Рабочий стол\stage5\НАС\ЛР6\Продажи.accdb"

def recreate_all_forms_reports():
    access = win32com.client.Dispatch("Access.Application")
    access.Visible = False
    access.OpenCurrentDatabase(DB_PATH, True) # Exclusive mode

    # Ensure Navigation Pane shows Object Types
    try:
        access.DoCmd.NavigateTo("acNavigationCategoryObjectType")
    except Exception as e:
        print("NavigateTo error:", e)

    # -------------------------------------------------------------
    # 1. Форма_Товары (Одностолбцовая форма с кнопками)
    # -------------------------------------------------------------
    print("Rebuilding Форма_Товары...")
    try:
        # Delete old if exists
        try:
            access.DoCmd.DeleteObject(2, "Форма_Товары") # 2 = acForm
        except:
            pass

        frm = access.CreateForm()
        f_name = frm.Name
        frm.RecordSource = "Товары"
        frm.Caption = "Каталог товаров"
        frm.DefaultView = 0 # Single Form

        # Detail fields
        fields = [
            ("Код товара:", "Код товара", False),
            ("Наименование:", "Наименование", True),
            ("Марка:", "Марка", True),
            ("Цена ($):", "Цена", True),
            ("Номер склада:", "Номер склада", True),
            ("Количество:", "Количество", True),
            ("Описание:", "Описание", True)
        ]

        top = 400
        for lbl_text, col_name, is_editable in fields:
            lbl = access.CreateControl(f_name, 100, 0, "", "", 400, top, 2400, 350)
            lbl.Caption = lbl_text
            lbl.FontBold = True
            lbl.FontSize = 10

            txt = access.CreateControl(f_name, 109, 0, "", col_name, 2900, top, 4500, 350 if col_name != "Описание" else 700)
            txt.ControlSource = col_name
            txt.FontSize = 10
            if col_name == "Цена":
                txt.Format = "Currency"

            top += 500 if col_name != "Описание" else 850

        # Navigation buttons
        top += 200
        btn_defs = [
            ("cmdFirst", "|<< Первая", 400),
            ("cmdPrev", "< Назад", 1900),
            ("cmdNext", "Вперед >", 3400),
            ("cmdLast", ">>| Последняя", 4900),
            ("cmdNew", "+ Добавить", 6400)
        ]
        for b_name, b_cap, b_left in btn_defs:
            btn = access.CreateControl(f_name, 104, 0, "", "", b_left, top, 1400, 420)
            btn.Name = b_name
            btn.Caption = b_cap
            btn.FontBold = True
            btn.FontSize = 9

        access.DoCmd.Close(2, f_name, 1) # acSaveYes = 1
        access.DoCmd.Rename("Форма_Товары", 2, f_name)
        print("Форма_Товары rebuilt successfully.")
    except Exception as e:
        print("Error rebuilding Форма_Товары:", e)

    # -------------------------------------------------------------
    # 2. Форма_Фирмы (Master-Detail: Фирмы + Контактные лица)
    # -------------------------------------------------------------
    print("Rebuilding Форма_Фирмы...")
    try:
        try:
            access.DoCmd.DeleteObject(2, "Форма_Фирмы")
        except:
            pass

        frm = access.CreateForm()
        f_name = frm.Name
        frm.RecordSource = "Фирмы"
        frm.Caption = "Фирмы и контактные лица"

        # Master fields
        m_fields = [
            ("Код фирмы:", "Код фирмы", 400, 300, 2200, 1500),
            ("Название фирмы:", "Название фирмы", 4200, 300, 2400, 3500),
            ("Адрес:", "Адрес", 400, 800, 2200, 3500),
            ("Телефон:", "Телефон", 4200, 800, 2400, 3500)
        ]
        for lbl_text, col_name, lx, tx, lw, tw in m_fields:
            lbl = access.CreateControl(f_name, 100, 0, "", "", lx, tx, lw, 350)
            lbl.Caption = lbl_text
            lbl.FontBold = True
            txt = access.CreateControl(f_name, 109, 0, "", col_name, lx + lw + 100, tx, tw, 350)
            txt.ControlSource = col_name

        # Label for subform
        lbl_sub = access.CreateControl(f_name, 100, 0, "", "", 400, 1400, 5000, 350)
        lbl_sub.Caption = "Сотрудники и контактные лица фирмы:"
        lbl_sub.FontBold = True
        lbl_sub.ForeColor = 10040119 # Dark Blue

        # Subform
        sub = access.CreateControl(f_name, 112, 0, "", "", 400, 1800, 9500, 4200)
        sub.SourceObject = "Table.Контактные лица"
        sub.LinkMasterFields = "Код фирмы"
        sub.LinkChildFields = "Код фирмы"

        access.DoCmd.Close(2, f_name, 1)
        access.DoCmd.Rename("Форма_Фирмы", 2, f_name)
        print("Форма_Фирмы rebuilt successfully.")
    except Exception as e:
        print("Error rebuilding Форма_Фирмы:", e)

    # -------------------------------------------------------------
    # 3. Форма_Склады (Master-Detail: Склады + Товары)
    # -------------------------------------------------------------
    print("Rebuilding Форма_Склады...")
    try:
        try:
            access.DoCmd.DeleteObject(2, "Форма_Склады")
        except:
            pass

        frm = access.CreateForm()
        f_name = frm.Name
        frm.RecordSource = "Склады"
        frm.Caption = "Склады и товары на складе"

        m_fields = [
            ("Номер склада:", "Номер склада", 400, 300, 2200, 1500),
            ("Заведующий:", "Заведующий", 4200, 300, 2400, 3500),
            ("Адрес склада:", "Адрес", 400, 800, 2200, 3500),
            ("Телефон склада:", "Телефон", 4200, 800, 2400, 3500)
        ]
        for lbl_text, col_name, lx, tx, lw, tw in m_fields:
            lbl = access.CreateControl(f_name, 100, 0, "", "", lx, tx, lw, 350)
            lbl.Caption = lbl_text
            lbl.FontBold = True
            txt = access.CreateControl(f_name, 109, 0, "", col_name, lx + lw + 100, tx, tw, 350)
            txt.ControlSource = col_name

        lbl_sub = access.CreateControl(f_name, 100, 0, "", "", 400, 1400, 5000, 350)
        lbl_sub.Caption = "Товары, хранящиеся на складе:"
        lbl_sub.FontBold = True
        lbl_sub.ForeColor = 10040119

        sub = access.CreateControl(f_name, 112, 0, "", "", 400, 1800, 9500, 4200)
        sub.SourceObject = "Table.Товары"
        sub.LinkMasterFields = "Номер склада"
        sub.LinkChildFields = "Номер склада"

        access.DoCmd.Close(2, f_name, 1)
        access.DoCmd.Rename("Форма_Склады", 2, f_name)
        print("Форма_Склады rebuilt successfully.")
    except Exception as e:
        print("Error rebuilding Форма_Склады:", e)

    # -------------------------------------------------------------
    # 4. Форма_Телевизоры (Цены, марка, примечание с Min ценой и кнопка)
    # -------------------------------------------------------------
    print("Rebuilding Форма_Телевизоры...")
    try:
        try:
            access.DoCmd.DeleteObject(2, "Форма_Телевизоры")
        except:
            pass

        frm = access.CreateForm()
        f_name = frm.Name
        frm.RecordSource = "SELECT * FROM [Товары 1] WHERE Наименование='Телевизор';"
        frm.Caption = "Мониторинг цен на телевизоры"

        # Detail fields
        top = 400
        tv_fields = [
            ("Код телевизора:", "Код товара"),
            ("Марка телевизора:", "Марка"),
            ("Отпускная цена ($):", "Цена"),
            ("Количество на складе:", "Количество"),
            ("Техническое описание:", "Описание")
        ]
        for lbl_text, col_name in tv_fields:
            lbl = access.CreateControl(f_name, 100, 0, "", "", 400, top, 2500, 350)
            lbl.Caption = lbl_text
            lbl.FontBold = True
            txt = access.CreateControl(f_name, 109, 0, "", col_name, 3000, top, 4500, 350 if col_name != "Описание" else 700)
            txt.ControlSource = col_name
            if col_name == "Цена":
                txt.Format = "Currency"
            top += 500 if col_name != "Описание" else 850

        # Footer section (Section = 2)
        # Note: In Access form, to add controls in footer, Section 2 is used
        try:
            lbl_min = access.CreateControl(f_name, 100, 2, "", "", 400, 200, 2800, 350)
            lbl_min.Caption = "Минимальная цена (=Min):"
            lbl_min.FontBold = True

            txt_min = access.CreateControl(f_name, 109, 2, "", "", 3300, 200, 2000, 350)
            txt_min.ControlSource = "=Min([Цена])"
            txt_min.Format = "Currency"
            txt_min.FontBold = True

            btn_upd = access.CreateControl(f_name, 104, 2, "", "", 5600, 150, 3200, 450)
            btn_upd.Caption = "Применить скидку 10%"
            btn_upd.FontBold = True
        except Exception as ef:
            print("Footer control notice:", ef)
            # Add to detail if footer not visible
            lbl_min = access.CreateControl(f_name, 100, 0, "", "", 400, top + 200, 2800, 350)
            lbl_min.Caption = "Минимальная цена (=Min):"
            lbl_min.FontBold = True
            txt_min = access.CreateControl(f_name, 109, 0, "", "", 3300, top + 200, 2000, 350)
            txt_min.ControlSource = "=Min([Цена])"
            txt_min.Format = "Currency"
            btn_upd = access.CreateControl(f_name, 104, 0, "", "", 5600, top + 150, 3200, 450)
            btn_upd.Caption = "Применить скидку 10%"

        access.DoCmd.Close(2, f_name, 1)
        access.DoCmd.Rename("Форма_Телевизоры", 2, f_name)
        print("Форма_Телевизоры rebuilt successfully.")
    except Exception as e:
        print("Error rebuilding Форма_Телевизоры:", e)

    # -------------------------------------------------------------
    # 5. Отчет_Товары (Простой отчет по таблице Товары)
    # -------------------------------------------------------------
    print("Rebuilding Отчет_Товары...")
    try:
        try:
            access.DoCmd.DeleteObject(3, "Отчет_Товары") # 3 = acReport
        except:
            pass

        rpt = access.CreateReport()
        r_name = rpt.Name
        rpt.RecordSource = "Товары"
        rpt.Caption = "Каталог товаров компании"

        # Report Header (Section = 1)
        try:
            lbl_hdr = access.CreateReportControl(r_name, 100, 1, "", "", 400, 200, 8000, 500)
            lbl_hdr.Caption = "КАТАЛОГ И ОСТАТКИ ТОВАРОВ"
            lbl_hdr.FontSize = 14
            lbl_hdr.FontBold = True
            lbl_hdr.ForeColor = 10040119
        except Exception as eh:
            print("Report header notice:", eh)

        # Page Header (Section = 3): Column headers
        try:
            col_heads = [
                ("Код", 400, 800),
                ("Наименование", 1300, 2200),
                ("Марка", 3600, 2800),
                ("Цена", 6500, 1400),
                ("Склад", 8000, 1000),
                ("Остаток", 9100, 1100)
            ]
            for h_text, h_left, h_w in col_heads:
                hl = access.CreateReportControl(r_name, 100, 3, "", "", h_left, 100, h_w, 350)
                hl.Caption = h_text
                hl.FontBold = True
                hl.FontSize = 9
        except Exception as ep:
            print("Page header notice:", ep)

        # Detail Section (Section = 0): Data columns
        col_data = [
            ("Код товара", 400, 800),
            ("Наименование", 1300, 2200),
            ("Марка", 3600, 2800),
            ("Цена", 6500, 1400),
            ("Номер склада", 8000, 1000),
            ("Количество", 9100, 1100)
        ]
        for c_field, c_left, c_w in col_data:
            txt = access.CreateReportControl(r_name, 109, 0, "", c_field, c_left, 50, c_w, 320)
            txt.ControlSource = c_field
            txt.FontSize = 9
            if c_field == "Цена":
                txt.Format = "Currency"

        access.DoCmd.Close(3, r_name, 1)
        access.DoCmd.Rename("Отчет_Товары", 3, r_name)
        print("Отчет_Товары rebuilt successfully.")
    except Exception as e:
        print("Error rebuilding Отчет_Товары:", e)

    # -------------------------------------------------------------
    # 6. Отчет_Музыкальные_Центры
    # -------------------------------------------------------------
    print("Rebuilding Отчет_Музыкальные_Центры...")
    try:
        try:
            access.DoCmd.DeleteObject(3, "Отчет_Музыкальные_Центры")
        except:
            pass

        rpt = access.CreateReport()
        r_name = rpt.Name
        rpt.RecordSource = """SELECT Товары.Марка, Товары.Цена, Фирмы.[Название фирмы]
FROM Фирмы INNER JOIN (Товары INNER JOIN Продажи ON Товары.[Код товара] = Продажи.[Код товара])
     ON Фирмы.[Код фирмы] = Продажи.[Код фирмы]
WHERE Товары.Наименование='Музыкальный центр';"""
        rpt.Caption = "Музыкальные центры"

        # Report Header (Section = 1)
        try:
            lbl_title = access.CreateReportControl(r_name, 100, 1, "", "", 400, 200, 9000, 500)
            lbl_title.Caption = "ОТЧЕТ ПО РЕАЛИЗАЦИИ МУЗЫКАЛЬНЫХ ЦЕНТРОВ"
            lbl_title.FontSize = 13
            lbl_title.FontBold = True
            lbl_title.ForeColor = 10040119
        except Exception as eh:
            print("Title notice:", eh)

        # Page Header (Section = 3): Date
        try:
            lbl_dt = access.CreateReportControl(r_name, 100, 3, "", "", 400, 100, 2500, 300)
            lbl_dt.Caption = "Дата формирования:"
            lbl_dt.FontSize = 8.5
            lbl_dt.FontItalic = True

            txt_dt = access.CreateReportControl(r_name, 109, 3, "", "", 3000, 100, 2200, 300)
            txt_dt.ControlSource = "=Date()"
            txt_dt.FontSize = 8.5
            txt_dt.FontItalic = True

            # Column headers
            h1 = access.CreateReportControl(r_name, 100, 3, "", "", 400, 500, 3500, 350)
            h1.Caption = "Марка центра"
            h1.FontBold = True
            h2 = access.CreateReportControl(r_name, 100, 3, "", "", 4000, 500, 3500, 350)
            h2.Caption = "Фирма-покупатель"
            h2.FontBold = True
            h3 = access.CreateReportControl(r_name, 100, 3, "", "", 7600, 500, 2000, 350)
            h3.Caption = "Цена ($)"
            h3.FontBold = True
        except Exception as ep:
            print("Date notice:", ep)

        # Detail Section (Section = 0)
        t_mark = access.CreateReportControl(r_name, 109, 0, "", "Марка", 400, 50, 3500, 320)
        t_mark.ControlSource = "Марка"
        t_firm = access.CreateReportControl(r_name, 109, 0, "", "Название фирмы", 4000, 50, 3500, 320)
        t_firm.ControlSource = "Название фирмы"
        t_price = access.CreateReportControl(r_name, 109, 0, "", "Цена", 7600, 50, 2000, 320)
        t_price.ControlSource = "Цена"
        t_price.Format = "Currency"

        # Report Footer (Section = 2): Avg Price
        try:
            lbl_avg = access.CreateReportControl(r_name, 100, 2, "", "", 400, 200, 4500, 350)
            lbl_avg.Caption = "Средняя цена музыкальных центров:"
            lbl_avg.FontBold = True

            txt_avg = access.CreateReportControl(r_name, 109, 2, "", "", 5000, 200, 2500, 350)
            txt_avg.ControlSource = "=Avg([Цена])"
            txt_avg.Format = "Currency"
            txt_avg.FontBold = True
            txt_avg.ForeColor = 255 # Red
        except Exception as ef:
            print("Avg notice:", ef)

        access.DoCmd.Close(3, r_name, 1)
        access.DoCmd.Rename("Отчет_Музыкальные_Центры", 3, r_name)
        print("Отчет_Музыкальные_Центры rebuilt successfully.")
    except Exception as e:
        print("Error rebuilding Отчет_Музыкальные_Центры:", e)

    access.CloseCurrentDatabase()
    access.Quit()
    print("All forms and reports reconstructed with rich controls!")

if __name__ == '__main__':
    recreate_all_forms_reports()


import docx
import shutil
import os
import win32com.client

base_dir = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС'
src_titul = os.path.join(base_dir, 'КИС Задание КП Шаблон Титульник.docx')

def build_kp_titulnik(student_name, student_short, cipher, theme, out_path):
    doc = docx.Document(src_titul)
    
    # Page 1:
    # P11: cipher
    doc.paragraphs[11].text = cipher
    doc.paragraphs[11].runs[0].font.size = docx.shared.Pt(14)
    doc.paragraphs[11].runs[0].bold = True
    doc.paragraphs[11].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    
    # P14: Student name and group
    doc.paragraphs[14].text = f"Обучающийся             {student_short}               группа ЭАС-514С\t"
    
    # P18: Руководитель
    doc.paragraphs[18].text = "Руководитель ________________ (Демченко М.С.) "
    
    # P22: Уфа-2026
    doc.paragraphs[22].text = "Уфа-2026"
    doc.paragraphs[22].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    
    # Page 2:
    # Table 0:
    t0 = doc.tables[0]
    t0.rows[0].cells[1].text = f"  {student_short}"
    t0.rows[0].cells[3].text = "  ЭАС-514С"
    t0.rows[0].cells[5].text = "  Демченко М.С."
    
    # Table 1:
    t1 = doc.tables[1]
    # R0: Theme
    t1.rows[0].cells[1].text = f"«{theme}»"
    t1.rows[0].cells[2].text = f"«{theme}»"
    # Set compact font size for theme to not blow up row height
    for c in [t1.rows[0].cells[1], t1.rows[0].cells[2]]:
        for p in c.paragraphs:
            p.paragraph_format.space_before = docx.shared.Pt(0)
            p.paragraph_format.space_after = docx.shared.Pt(0)
            p.paragraph_format.line_spacing = 1.0
            for r in p.runs:
                r.font.size = docx.shared.Pt(9.5)
                
    # R3: Compact spacing for main text
    for c in [t1.rows[3].cells[1], t1.rows[3].cells[2]]:
        for p in c.paragraphs:
            p.paragraph_format.space_before = docx.shared.Pt(0)
            p.paragraph_format.space_after = docx.shared.Pt(0)
            p.paragraph_format.line_spacing = 1.0
            for r in p.runs:
                r.font.size = docx.shared.Pt(8.5)

    # R4: Исходные данные
    t1.rows[4].cells[1].text = (
        "Вариант № 102. Предметная область: «Автоматизация процессов чип-тюнинга и сервисного "
        "обслуживания автопарка УУНиТ в автотехцентре BMW102». "
        "Детальный бизнес-сценарий (ДБС) 102BMW-BS-01."
    )
    for p in t1.rows[4].cells[1].paragraphs:
        p.paragraph_format.space_before = docx.shared.Pt(0)
        p.paragraph_format.space_after = docx.shared.Pt(0)
        p.paragraph_format.line_spacing = 1.0
        for r in p.runs:
            r.font.size = docx.shared.Pt(8.5)

    # P43: Dates
    doc.paragraphs[43].text = "Дата выдачи «01» сентября 2026 г.                     Дата окончания «16» декабря 2026 г."
    doc.paragraphs[43].paragraph_format.space_before = docx.shared.Pt(2)
    doc.paragraphs[43].paragraph_format.space_after = docx.shared.Pt(0)
    
    # Remove empty P44
    p44 = doc.paragraphs[44]
    p44._p.getparent().remove(p44._p)
    
    # P45: Signatures
    doc.paragraphs[44].text = f"Руководитель __________________ М.С. Демченко             Студент _________________ {student_short}"
    doc.paragraphs[44].paragraph_format.space_before = docx.shared.Pt(2)
    doc.paragraphs[44].paragraph_format.space_after = docx.shared.Pt(0)
    
    doc.save(out_path)
    print("Saved:", out_path)

f1 = os.path.join(base_dir, 'КИС_Задание_КП_Гарифуллин_К_Р.docx')
f2 = os.path.join(base_dir, 'КИС_Задание_КП_Гайнанов_Д_И.docx')

build_kp_titulnik(
    student_name="Гарифуллин Карим Рамилевич",
    student_short="Гарифуллин К.Р.",
    cipher="3211. 102BMW01 02 ПЗ",
    theme="Техническое задание на разработку отчетной формы «Отчет по количеству исполненных заявок за период»",
    out_path=f1
)

build_kp_titulnik(
    student_name="Гайнанов Данил Ильдарович",
    student_short="Гайнанов Д.И.",
    cipher="3211. 102BMW02 02 ПЗ",
    theme="Техническое задание на разработку отчетной формы «Отчет по клиентам за период»",
    out_path=f2
)

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0

for pth in [f1, f2]:
    doc = word.Documents.Open(os.path.abspath(pth))
    pgs = doc.ComputeStatistics(2)
    print(f"Verified {os.path.basename(pth)}: EXACTLY {pgs} PAGES!")
    doc.Close()

word.Quit()

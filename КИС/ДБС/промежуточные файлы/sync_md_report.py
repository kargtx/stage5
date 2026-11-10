# -*- coding: utf-8 -*-
"""
Генератор синхронизированного Markdown-отчета ДБС
"""
import os
import docx

doc_path = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\Отчет_по_ЛР_Часть_2_УУНиТ_Гайнанов_Гарифуллин.docx"
if not os.path.exists(doc_path):
    doc_path = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС\измОтчет_по_ЛР_Часть_2_УУНиТ_Гайнанов_Гарифуллин.docx"
md_path = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС\Отчет_ДБС_ЧипТюнинг.md"

doc = docx.Document(doc_path)
lines = []

lines.append("# Министерство науки и высшего образования Российской Федерации")
lines.append("## ФГБОУ ВО «Уфимский университет науки и технологий» (УУНиТ)")
lines.append("### Кафедра автоматизированных систем управления\n")
lines.append("---\n")
lines.append("# ДЕТАЛЬНЫЙ БИЗНЕС-СЦЕНАРИЙ (ДБС)")
lines.append("## 102BMW-BS-01 «Организация и исполнение услуг чип-тюнинга и технического обслуживания автопарка»\n")
lines.append("**Выполнили:** студенты гр. ЭАС-514С Гайнанов Д.И., Гарифуллин К.Р.  ")
lines.append("**Проверил:** ст. преподаватель каф. АСУ Демченко М.С.  ")
lines.append("**Год:** 2026\n")
lines.append("---\n")

body = doc._element.body
tbl_idx = 0

for child in body:
    if child.tag.endswith('p'):
        p = docx.text.paragraph.Paragraph(child, doc)
        text = p.text.strip()
        if not text:
            continue
        style = p.style.name
        if style == 'Heading 1':
            lines.append(f"\n# {text}\n")
        elif style == 'Heading 2':
            lines.append(f"\n## {text}\n")
        elif style == 'Heading 3':
            lines.append(f"\n### {text}\n")
        elif style == 'Heading 4':
            lines.append(f"\n#### {text}\n")
        elif style == 'Caption':
            lines.append(f"\n*{text}*\n")
        elif style == 'IP_20_List1':
            lines.append(f"- {text}")
        else:
            lines.append(f"{text}\n")
            
    elif child.tag.endswith('tbl'):
        t = doc.tables[tbl_idx]
        tbl_idx += 1
        if tbl_idx in [1, 2]: # Пропускаем шапку титульного листа
            continue
        # Рендерим markdown-таблицу
        if len(t.rows) > 0:
            header_cells = [c.text.replace('\n', ' ').strip() for c in t.rows[0].cells]
            lines.append("| " + " | ".join(header_cells) + " |")
            lines.append("| " + " | ".join([":---" if j > 0 else ":---" for j in range(len(header_cells))]) + " |")
            for r in t.rows[1:]:
                row_cells = [c.text.replace('\n', ' ').strip() for c in r.cells]
                lines.append("| " + " | ".join(row_cells) + " |")
            lines.append("\n")

with open(md_path, 'w', encoding='utf-8') as f:
    f.write("\n".join(lines))

print(f"Синхронизированный Markdown-отчет сохранен: {md_path}")


# -*- coding: utf-8 -*-
"""
Генератор синхронизированного Markdown-отчета ЛР КИС ЧАСТЬ 1
"""
import os
import docx

doc_path = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\лр часть 1\Отчет_по_ЛР_Часть_1_УУНиТ_Гайнанов_Гарифуллин.docx"
md_path = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\лр часть 1\Отчет_по_ЛР_Часть_1.md"
md_root = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\Отчет_по_ЛР_Часть_1.md"

doc = docx.Document(doc_path)
lines = []

lines.append("# Министерство науки и высшего образования Российской Федерации")
lines.append("## ФГБОУ ВО «Уфимский университет науки и технологий» (УУНиТ)")
lines.append("### Кафедра автоматизированных систем управления\n")
lines.append("---\n")
lines.append("# ОТЧЕТ ПО ЛАБОРАТОРНЫМ РАБОТАМ")
lines.append("## Дисциплина «Корпоративные информационные системы» — ЧАСТЬ 1")
lines.append("### Тема: «Проектирование корпоративной информационной системы автосервиса чип-тюнинга BMW102: технико-экономическая характеристика объекта автоматизации и реляционная модель данных (ERD)»\n")
lines.append("**Выполнили:** студенты гр. ЭАС-514С Гайнанов Д.И., Гарифуллин К.Р.  ")
lines.append("**Проверил:** ст. преподаватель каф. АСУ Демченко М.С.  ")
lines.append("**Год:** 2026\n")
lines.append("---\n")

body = doc._element.body
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
        else:
            lines.append(f"{text}\n")
            
    elif child.tag.endswith('tbl'):
        t = docx.table.Table(child, doc)
        if len(t.rows) == 0:
            continue
        # Render markdown table
        num_cols = len(t.rows[0].cells)
        header_vals = [c.text.strip().replace('\n', ' ') for c in t.rows[0].cells]
        lines.append("\n| " + " | ".join(header_vals) + " |")
        lines.append("| " + " | ".join([":---"] * num_cols) + " |")
        for row in t.rows[1:]:
            row_vals = [c.text.strip().replace('\n', ' ') for c in row.cells]
            lines.append("| " + " | ".join(row_vals) + " |")
        lines.append("\n")

md_content = "\n".join(lines)
with open(md_path, "w", encoding="utf-8") as f:
    f.write(md_content)
with open(md_root, "w", encoding="utf-8") as f:
    f.write(md_content)

print("Markdown сохранен в:", md_path)
print("Markdown скопирован в:", md_root)


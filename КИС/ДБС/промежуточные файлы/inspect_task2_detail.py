# -*- coding: utf-8 -*-
import docx

p = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\лр часть 1\КИС Задание на ЛР ЧАСТЬ 2.docx"
doc = docx.Document(p)

print("=== PARAGRAPHS DETAIL ===")
for i, para in enumerate(doc.paragraphs):
    print(f"P{i:02d}: align={para.alignment}, style={para.style.name}, runs={[r.text for r in para.runs]}")

print("\n=== TABLES DETAIL ===")
for i, t in enumerate(doc.tables):
    print(f"Table {i} rows={len(t.rows)}, cols={len(t.columns)}")
    for ri, r in enumerate(t.rows):
        print(f"  R{ri}: {[c.text.replace(chr(10), ' ') for c in r.cells]}")


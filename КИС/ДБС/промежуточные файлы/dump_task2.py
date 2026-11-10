# -*- coding: utf-8 -*-
import docx

p = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\лр часть 1\КИС Задание на ЛР ЧАСТЬ 2.docx"
doc = docx.Document(p)

print("=== PARAGRAPHS ===")
for i, para in enumerate(doc.paragraphs):
    if para.text.strip():
        print(f"P{i:02d} [{para.style.name}]: {repr(para.text)}")

print("\n=== TABLES ===")
for i, t in enumerate(doc.tables):
    print(f"\n--- Table {i} ({len(t.rows)}x{len(t.columns)}) ---")
    for r in t.rows:
        print([c.text.replace('\n', ' ') for c in r.cells])

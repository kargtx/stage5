# -*- coding: utf-8 -*-
import docx

p = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\промежуточные файлы\Отчет_Динамическая_модель_подпроцесса.docx"
doc = docx.Document(p)

print("=== INSPECTING PREVIOUS REPORT ===")
print("Paragraphs:", len(doc.paragraphs))
print("Tables:", len(doc.tables))
print("Sections:", len(doc.sections))

for si, s in enumerate(doc.sections):
    h = s.header
    print(f"Section {si}: header tables={len(h.tables)}, header paragraphs={len(h.paragraphs)}, text={[p.text for p in h.paragraphs if p.text.strip()]}")

print("\n--- First 30 paragraphs ---")
for i in range(min(30, len(doc.paragraphs))):
    p_text = doc.paragraphs[i].text.strip()
    if p_text:
        print(f"P{i:02d} [{doc.paragraphs[i].style.name}]: {repr(p_text[:80])}")

print("\n--- All Headings ---")
for i, p in enumerate(doc.paragraphs):
    if p.style.name.startswith("Heading"):
        print(f"P{i:02d} [{p.style.name}]: {p.text}")


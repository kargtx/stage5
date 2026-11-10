import docx

doc = docx.Document(r"C:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС\ДБС_Нигматуллин_Гайнетдинов_Рогозин_v2.docx")

print("=== PARAGRAPHS OVERVIEW ===")
for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if not txt:
        continue
    # print headings or short lines
    if p.style.name.startswith("Heading") or p.style.name.startswith("Заголовок") or p.style.name.startswith("IP_") or len(txt) < 80:
        print(f"P{i:03d} [{p.style.name:15s}]: {txt}")

print("\n=== ALL TABLES SUMMARY ===")
for t_idx, t in enumerate(doc.tables):
    rows = len(t.rows)
    cols = len(t.columns)
    r0 = [c.text.strip().replace('\n', ' ')[:35] for c in t.rows[0].cells]
    r1 = [c.text.strip().replace('\n', ' ')[:35] for c in t.rows[1].cells] if rows > 1 else []
    print(f"Table {t_idx:02d} ({rows}x{cols}):")
    print(f"   Row 0: {r0}")
    if rows > 1:
        print(f"   Row 1: {r1}")


import docx

doc = docx.Document(r"C:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС\ДБС_Нигматуллин_Гайнетдинов_Рогозин_v2.docx")

with open(r"C:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС\промежуточные файлы\sample_text_dump.txt", "w", encoding="utf-8") as out:
    out.write("=== PARAGRAPHS ===\n")
    for i, p in enumerate(doc.paragraphs):
        out.write(f"[{i:03d}][{p.style.name}] {p.text}\n")
    
    out.write("\n=== TABLES ===\n")
    for t_idx, t in enumerate(doc.tables):
        out.write(f"\n--- TABLE {t_idx} ({len(t.rows)}x{len(t.columns)}) ---\n")
        for r_idx, r in enumerate(t.rows):
            cells = [c.text.strip().replace('\n', ' ') for c in r.cells]
            out.write(f"  R{r_idx}: {cells}\n")

print("Dumped successfully to sample_text_dump.txt")


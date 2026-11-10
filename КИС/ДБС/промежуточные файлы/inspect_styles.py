import docx
from docx.shared import Pt, RGBColor

sample_path = r"C:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС\ДБС_Нигматуллин_Гайнетдинов_Рогозин_v2.docx"
doc = docx.Document(sample_path)

print("=== STYLES & FONTS IN SAMPLE ===")
for p in doc.paragraphs[:50]:
    if p.text.strip():
        runs_info = []
        for r in p.runs[:3]:
            runs_info.append(f"{r.font.name} {r.font.size.pt if r.font.size else 'None'}pt bold={r.bold} italic={r.italic} color={r.font.color.rgb if r.font.color else 'None'}")
        p_format = f"align={p.alignment} indent={p.paragraph_format.first_line_indent.pt if p.paragraph_format.first_line_indent else 'None'} sb={p.paragraph_format.space_before.pt if p.paragraph_format.space_before else 'None'} sa={p.paragraph_format.space_after.pt if p.paragraph_format.space_after else 'None'} ls={p.paragraph_format.line_spacing}"
        print(f"[{p.style.name}] text={repr(p.text[:40])}\n  runs={runs_info}\n  format={p_format}")

print("\n=== TABLE STYLES & BORDERS ===")
for t_idx in [2, 3, 4, 7, 8]:
    t = doc.tables[t_idx]
    c = t.rows[0].cells[0]
    p = c.paragraphs[0]
    r = p.runs[0] if p.runs else None
    print(f"Table {t_idx} style={t.style.name if t.style else 'None'} alignment={t.alignment}")
    if r:
        print(f"  Header cell font: {r.font.name} {r.font.size.pt if r.font.size else 'None'}pt bold={r.bold}")
    # inspect cell tcPr xml
    print(f"  Cell xml: {c._tc.xml[:200]}")
    # inspect tblPr xml
    print(f"  TblPr xml: {t._tbl.tblPr.xml[:300]}")


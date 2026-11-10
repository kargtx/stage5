import zipfile
import re
import docx

sample_path = r"C:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС\ДБС_Нигматуллин_Гайнетдинов_Рогозин_v2.docx"

with zipfile.ZipFile(sample_path) as z:
    for name in z.namelist():
        if name.endswith('.rels'):
            xml = z.read(name).decode('utf-8', errors='ignore')
            for m in re.finditer(r'Target="([^"]+)"', xml):
                target = m.group(1)
                if 'media' in target:
                    print(f"{name} -> {target}")

print("\n--- Document structure inspection ---")
doc = docx.Document(sample_path)
print(f"Total paragraphs: {len(doc.paragraphs)}")
print(f"Total tables: {len(doc.tables)}")

# Inspect all tables
for i, t in enumerate(doc.tables):
    rows = len(t.rows)
    cols = len(t.columns)
    header = [c.text.strip().replace('\n', ' ')[:40] for c in t.rows[0].cells]
    print(f"Table {i}: {rows}x{cols} -> {header}")


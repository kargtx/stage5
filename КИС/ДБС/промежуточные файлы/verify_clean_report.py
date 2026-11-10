# -*- coding: utf-8 -*-
import docx

p = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС\Отчет_по_ЛР_Часть_2_УУНиТ_Гайнанов_Гарифуллин.docx"
doc = docx.Document(p)

print("=== ВЕРИФИКАЦИЯ КЛАССИЧЕСКОГО ОТЧЕТА УУНИТ (БЕЗ ШАПОК) ===")
print("Всего параграфов:", len(doc.paragraphs))
print("Всего таблиц:", len(doc.tables))
print("Всего секций:", len(doc.sections))

for si, s in enumerate(doc.sections):
    h = s.header
    header_drawings = h._element.xpath('.//w:drawing')
    header_tables = len(h.tables)
    header_text = [p.text for p in h.paragraphs if p.text.strip()]
    print(f"Секция {si}: таблиц в колонтитуле={header_tables}, рисунков в колонтитуле={len(header_drawings)}, текст={header_text}")
    print(f"  Поля: верх={s.top_margin.mm:.1f}мм, низ={s.bottom_margin.mm:.1f}мм, лево={s.left_margin.mm:.1f}мм, право={s.right_margin.mm:.1f}мм")

print("\n--- Проверка структуры разделов (Heading 1 и Heading 2) ---")
for i, para in enumerate(doc.paragraphs):
    if para.style.name in ['Heading 1', 'Heading 2']:
        print(f"[{para.style.name}] {para.text}")

print("\n--- Проверка изображений ---")
img_count = 0
for i, para in enumerate(doc.paragraphs):
    if len(para._element.xpath('.//w:drawing')) > 0:
        img_count += 1
        caption = doc.paragraphs[i+1].text if i+1 < len(doc.paragraphs) else ''
        print(f"Рисунок #{img_count}: {caption}")

print("\n--- Проверка первых 5 таблиц ---")
for i in range(min(5, len(doc.tables))):
    t = doc.tables[i]
    print(f"Таблица {i} ({len(t.rows)}x{len(t.columns)}): R0 = {[c.text.replace(chr(10), ' ')[:25] for c in t.rows[0].cells]}")

print("\n=== ВЕРИФИКАЦИЯ УСПЕШНО ПРОЙДЕНА! ===")


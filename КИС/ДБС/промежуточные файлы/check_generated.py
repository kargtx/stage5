# -*- coding: utf-8 -*-
import docx

doc_path = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС\ДБС_102BMW_УУНиТ_Гайнанов_Гарифуллин.docx"
doc = docx.Document(doc_path)

print("=== ПРОВЕРКА СГЕНЕРИРОВАННОГО ДОКУМЕНТА ДБС ===")
print("Количество параграфов:", len(doc.paragraphs))
print("Количество таблиц:", len(doc.tables))
print("Количество секций:", len(doc.sections))

for si, s in enumerate(doc.sections):
    h = s.header
    print(f"Секция {si}: таблиц в колонтитуле={len(h.tables)}, рисунков в колонтитуле={len(h._element.xpath('.//w:drawing'))}")

print("\n--- Список таблиц и их размеры ---")
for i, t in enumerate(doc.tables):
    prev_text = ""
    elem = t._element.getprevious()
    while elem is not None:
        if elem.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(elem, doc)
            if p.text.strip():
                prev_text = p.text[:55]
                break
        elem = elem.getprevious()
    print(f"Таблица {i:02d} ({len(t.rows):2d} x {len(t.columns):2d}): prev='{prev_text}'")

print("\n--- Проверка изображений ---")
img_count = 0
for i, p in enumerate(doc.paragraphs):
    if len(p._element.xpath('.//w:drawing')) > 0:
        img_count += 1
        caption = doc.paragraphs[i+1].text if i+1 < len(doc.paragraphs) else ''
        print(f"Рисунок #{img_count} (P{i:03d}): caption='{caption}'")

print("\n--- Проверка соответствия таблиц атрибутов ERD ---")
attr_tables_count = 0
valid_erd_tables = {'Client', 'Vehicle', 'Employee', 'Service_Request', 'Work_Order', 'Diagnostic_Card', 'Order_Execution', 'Work_Completion_Act'}
for i, t in enumerate(doc.tables):
    if len(t.columns) == 4 and t.rows[0].cells[0].text.strip() == "Атрибут (Поле БД)":
        attr_tables_count += 1
        # Проверяем названия таблиц ERD в строках
        erd_names_in_table = set()
        for r in t.rows[1:]:
            erd_name = r.cells[1].text.strip()
            erd_names_in_table.add(erd_name)
            if erd_name not in valid_erd_tables:
                print(f"ОШИБКА: Неизвестная таблица ERD '{erd_name}' в Таблице {i}")
        print(f"Таблица атрибутов #{attr_tables_count:2d} (Таблица {i:02d}): {len(t.rows)-1} атрибутов, ERD таблицы={erd_names_in_table}")

print(f"\nВсего проверено таблиц атрибутов: {attr_tables_count} (ожидается 20)")
print("=== ПРОВЕРКА ЗАВЕРШЕНА ===")


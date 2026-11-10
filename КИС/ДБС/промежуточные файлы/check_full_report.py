# -*- coding: utf-8 -*-
import docx

p = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС\Отчет_ЛР_Часть2_ДБС_ЧипТюнинг_Гайнанов_Гарифуллин.docx"
doc = docx.Document(p)

print("=== ПРОВЕРКА ПОЛНОГО ОТЧЕТА ПО ЛР ЧАСТЬ 2 ===")
print("Всего параграфов:", len(doc.paragraphs))
print("Всего таблиц:", len(doc.tables))
print("Всего секций:", len(doc.sections))

print("\n--- Первые 30 параграфов (Титульный лист и Задание) ---")
for i in range(min(30, len(doc.paragraphs))):
    txt = doc.paragraphs[i].text.strip()
    if txt:
        print(f"P{i:02d} [{doc.paragraphs[i].style.name}]: {repr(txt[:75])}")

print("\n--- Последние 25 параграфов (Заключение и Список источников) ---")
for i in range(max(0, len(doc.paragraphs) - 25), len(doc.paragraphs)):
    txt = doc.paragraphs[i].text.strip()
    if txt:
        print(f"P{i:02d} [{doc.paragraphs[i].style.name}]: {repr(txt[:75])}")

print("\n--- Таблицы в отчете ---")
for i in range(min(6, len(doc.tables))):
    t = doc.tables[i]
    print(f"Таблица {i} ({len(t.rows)}x{len(t.columns)}): R0 = {[c.text.replace(chr(10), ' ')[:30] for c in t.rows[0].cells]}")

print("=== ПРОВЕРКА УСПЕШНО ЗАВЕРШЕНА ===")


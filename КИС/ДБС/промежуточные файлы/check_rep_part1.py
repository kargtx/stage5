# -*- coding: utf-8 -*-
import docx

p = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\промежуточные файлы\Отчет_Динамическая_модель_подпроцесса.docx"
try:
    doc = docx.Document(p)
    print("Paragraphs:", len(doc.paragraphs))
    print("Tables:", len(doc.tables))
    for i in range(min(25, len(doc.paragraphs))):
        if doc.paragraphs[i].text.strip():
            print(f"P{i:02d}: {repr(doc.paragraphs[i].text[:80])}")
except Exception as e:
    print("Error:", e)


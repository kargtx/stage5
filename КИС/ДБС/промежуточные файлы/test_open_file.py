# -*- coding: utf-8 -*-
import docx

paths = [
    r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС\ДБС_102BMW_УУНиТ_Гайнанов_Гарифуллин.docx",
    r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\ДБС_102BMW_УУНиТ_Гайнанов_Гарифуллин.docx",
    r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\измДБС_102BMW_УУНиТ_Гайнанов_Гарифуллин.docx"
]
for p in paths:
    try:
        with open(p, "rb") as f:
            data = f.read()
        print(f"Can read {len(data)} bytes from {p}")
        doc = docx.Document(p)
        print(f"  Successfully opened docx: {len(doc.paragraphs)} paras")
    except Exception as e:
        print(f"FAILED on {p}: {e}")


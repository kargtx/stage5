# -*- coding: utf-8 -*-
import os
import re

src_path = r'c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\МИМС\лр5\PVZ_Optimization.alp'
dst_path = r'c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\МИМС\лр5\AutoService_Optimization.alp'

with open(src_path, 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Model & Package Name
text = text.replace('<Name><![CDATA[PVZ_Optimization]]></Name>', '<Name><![CDATA[AutoService_Optimization]]></Name>')
text = text.replace('<JavaPackageName><![CDATA[pvz_optimization]]></JavaPackageName>', '<JavaPackageName><![CDATA[autoservice_optimization]]></JavaPackageName>')

# 2. Replaces in Text labels
replacements = [
    ('Пункт выдачи заказов (ПВЗ) — Вариант 5', 'Автосервис (СТО) — Оптимизация постов ремонта'),
    ('Постамат (20%)', 'Экспресс-пост / Диагностика (20%)'),
    ('Очередь (до 10 чел.)', 'Парковка ожидания (до 10 авто)'),
    ('Стойка выдачи', 'Посты ремонта и ТО'),
    ('Склад', 'Склад запчастей и ремзона'),
    ('• Загрузка кладовщиков:', '• Загрузка автомехаников:'),
    ('• Клиенты (обслужено / потеряно):', '• Автомобили (обслужено / потеряно):'),
    ('Оптимизация числа кладовщиков ПВЗ при пиковой нагрузке', 'Оптимизация числа автомехаников СТО при пиковой нагрузке'),
    ('Кладовщики (nWorkers):', 'Автомеханики (nWorkers):'),
    ('• Оптимум: N* = 5 кладовщиков', '• Оптимум: N* = 5 автомехаников (постов)'),
    ('• Отказы / уход клиентов: 0 чел. (0%)', '• Отказы / уезд клиентов: 0 авто (0%)'),
]

for old, new in replacements:
    text = text.replace(old, new)

with open(dst_path, 'w', encoding='utf-8') as f:
    f.write(text)

print("Saved adapted AnyLogic model to:", dst_path)

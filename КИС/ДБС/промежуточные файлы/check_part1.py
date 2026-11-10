# -*- coding: utf-8 -*-
import os

part1_dir = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\КИС\лр часть 1"
if os.path.exists(part1_dir):
    for f in os.listdir(part1_dir):
        fp = os.path.join(part1_dir, f)
        print(f, os.path.getsize(fp), "bytes")
else:
    print("Part 1 dir not found")


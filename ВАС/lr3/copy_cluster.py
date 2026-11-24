import os, shutil

src = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\ВАС\lr3\кластер_сеть"
dst = r"C:\cluster"

try:
    os.makedirs(dst, exist_ok=True)
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            if os.path.exists(d):
                shutil.rmtree(d)
            shutil.copytree(s, d)
        else:
            shutil.copy2(s, d)
    print("SUCCESS: Copied to C:\\cluster")
    print("Files in C:\\cluster:")
    for f in os.listdir(dst):
        print(" ", f)
except Exception as e:
    print("ERROR:", e)

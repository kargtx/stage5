import urllib.request
import ssl
import os
from PIL import Image

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

base_url = 'https://uust.ru'
targets = [
    ('/static/New_files_ugatu/images/redesign/brandbook/logo-full.png', 'uunit_logo_full.png'),
    ('/static/New_files_ugatu/images/redesign/brandbook/logo.png', 'uunit_logo.png'),
    ('/static/New_files_ugatu/images/redesign/brandbook/logo-sign.png', 'uunit_logo_sign.png'),
]

out_dir = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС'

for uri, fname in targets:
    url = base_url + uri
    out_path = os.path.join(out_dir, fname)
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    try:
        with urllib.request.urlopen(req, context=ctx, timeout=10) as r:
            data = r.read()
            with open(out_path, 'wb') as f:
                f.write(data)
        im = Image.open(out_path)
        print(f"Downloaded {fname}: size={im.size}, mode={im.mode}, bytes={len(data)}")
    except Exception as e:
        print(f"Error {fname}: {e}")

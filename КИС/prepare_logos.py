from PIL import Image
import os

uunit_path = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС\uunit_logo.png'
exec_path = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС\лого исполнителя.png'

im_uunit = Image.open(uunit_path)
print("UUNiT raw:", im_uunit.size, im_uunit.mode)

im_exec = Image.open(exec_path)
print("Exec raw:", im_exec.size, im_exec.mode)

# Let's inspect the bounding box of non-transparent pixels in both
def get_bbox(im):
    if im.mode != 'RGBA':
        im = im.convert('RGBA')
    alpha = im.split()[-1]
    return alpha.getbbox()

print("UUNiT bbox:", get_bbox(im_uunit))
print("Exec bbox:", get_bbox(im_exec))

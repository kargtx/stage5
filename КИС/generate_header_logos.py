from PIL import Image

def make_header_logo(src_path, target_w, target_h, pad_color=(255, 255, 255, 0), crop_box=None):
    im = Image.open(src_path)
    if im.mode != 'RGBA':
        im = im.convert('RGBA')
    if crop_box:
        im = im.crop(crop_box)
    else:
        # crop to non-transparent or non-white
        alpha = im.split()[-1]
        bbox = alpha.getbbox()
        if bbox:
            im = im.crop(bbox)

    # Scale to fit inside target_w x target_h preserving aspect ratio
    im_w, im_h = im.size
    scale = min(target_w / im_w, target_h / im_h)
    new_w = int(im_w * scale)
    new_h = int(im_h * scale)
    resized = im.resize((new_w, new_h), Image.Resampling.LANCZOS)

    # Create target canvas
    canvas = Image.new('RGBA', (target_w, target_h), pad_color)
    offset_x = (target_w - new_w) // 2
    offset_y = (target_h - new_h) // 2
    canvas.paste(resized, (offset_x, offset_y), resized)
    return canvas

# 1. Customer logo (UUNiT)
uunit_src = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС\uunit_logo.png'
uunit_logo = make_header_logo(uunit_src, 1182, 308, crop_box=(110, 115, 5455, 1270))
uunit_logo.save(r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС\prepared_uunit_header.png')
print("Saved prepared_uunit_header.png", uunit_logo.size)

# 2. Executor logo (BMW102)
exec_src = r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС\лого исполнителя.png'
exec_logo = make_header_logo(exec_src, 1000, 337, crop_box=(0, 160, 785, 275))
exec_logo.save(r'C:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\КИС\prepared_exec_header.png')
print("Saved prepared_exec_header.png", exec_logo.size)

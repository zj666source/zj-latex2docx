# -*- coding: utf-8 -*-
"""Generate gs.ico: a clean tech-green rounded square with white 'gs' letters."""
from PIL import Image, ImageDraw, ImageFont

FONT = r"C:\Windows\Fonts\arialbd.ttf"   # bold for crisp letters
OUT = r"D:\work文件\_latex2eq_build\gs.ico"
S = 256
GREEN = (22, 163, 74, 255)     # 清新科技绿 (#16A34A)
WHITE = (255, 255, 255, 255)

img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# rounded square background (inset a little so corners aren't clipped at small sizes)
d.rounded_rectangle([10, 10, S - 10, S - 10], radius=46, fill=GREEN)

# 'gs' text, centered
font = ImageFont.truetype(FONT, 150)
text = "gs"
bbox = d.textbbox((0, 0), text, font=font)
tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
cx, cy = (S - tw) / 2 - bbox[0], (S - th) / 2 - bbox[1]
# nudge slightly left so the visual weight is centered
d.text((cx - 4, cy - 6), text, font=font, fill=WHITE)

img.save(OUT, sizes=[(16, 16), (24, 24), (32, 32), (48, 48),
                     (64, 64), (128, 128), (256, 256)])
print("WROTE", OUT)

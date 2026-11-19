# -*- coding: utf-8 -*-
"""
make_icon.py — 生成程序图标（彩色、多尺寸 ico）。

设计：渐变底 + 白色账页 + 绿色对勾 + 一枚红点（差异）一枚绿点（已平），
远看是"核过的账"，缩小到 16px 也能认出对勾。

用法：python tools/make_icon.py
输出：assets/app.ico、assets/app.png（预览）
"""

import os

from PIL import Image, ImageDraw

S = 512                     # 先画大图再缩，边缘才干净
COLOR_TOP = (91, 141, 239)       # #5B8DEF
COLOR_BOT = (138, 92, 246)       # #8A5CF6
WHITE = (255, 255, 255)
LINE = (206, 218, 245)
CHECK = (34, 197, 94)            # #22C55E
RED = (248, 113, 113)            # #F87171


def _gradient(size):
    img = Image.new('RGB', (size, size), COLOR_TOP)
    d = ImageDraw.Draw(img)
    for y in range(size):
        t = y / max(size - 1, 1)
        c = tuple(int(COLOR_TOP[i] + (COLOR_BOT[i] - COLOR_TOP[i]) * t) for i in range(3))
        d.line([(0, y), (size, y)], fill=c)
    return img


def build():
    base = _gradient(S).convert('RGBA')

    # 圆角遮罩（做成 App 图标的样子）
    mask = Image.new('L', (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, S - 1, S - 1], radius=112, fill=255)
    base.putalpha(mask)

    d = ImageDraw.Draw(base)

    # 白色账页
    page = [116, 88, 396, 424]
    d.rounded_rectangle(page, radius=18, fill=(255, 255, 255, 246))

    # 账页上的表格横线
    for i in range(4):
        y = 150 + i * 52
        d.line([(158, y), (354, y)], fill=LINE, width=12)

    # 绿色对勾（压在右下角，带白描边保证小尺寸也清晰）
    pts = [(196, 300), (248, 352), (330, 244)]
    d.line(pts, fill=(255, 255, 255), width=44, joint='curve')
    d.line(pts, fill=CHECK, width=28, joint='curve')

    # 右下角小红点 / 左下角小绿点：一眼看出"差异 vs 已平"
    d.ellipse([318, 372, 372, 426], fill=RED, outline=(255, 255, 255), width=8)
    d.ellipse([140, 372, 194, 426], fill=CHECK, outline=(255, 255, 255), width=8)

    return base


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = os.path.join(root, 'assets')
    os.makedirs(out_dir, exist_ok=True)

    img = build()
    png = os.path.join(out_dir, 'app.png')
    img.resize((256, 256), Image.LANCZOS).save(png)

    ico = os.path.join(out_dir, 'app.ico')
    sizes = [16, 24, 32, 48, 64, 128, 256]
    frames = [img.resize((s, s), Image.LANCZOS) for s in sizes]
    frames[-1].save(ico, format='ICO', sizes=[(s, s) for s in sizes],
                    append_images=frames[:-1])
    print('已生成:', png)
    print('已生成:', ico)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())

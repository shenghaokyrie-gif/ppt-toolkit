# -*- coding: utf-8 -*-
"""
PPT Toolkit - 提取设计系统

分析一个 pptx，输出它的字号分布、颜色分布、标题/正文几何规律，
用来对齐既有模板的视觉规范，而不是凭空定参数。

用法：
    python extract_design.py <file.pptx>
"""
import sys
from collections import Counter
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation

EMU = 914400.0


def main():
    if len(sys.argv) < 2:
        print('用法: python extract_design.py <file.pptx>')
        return 2
    path = sys.argv[1]
    prs = Presentation(path)
    SW, SH = prs.slide_width, prs.slide_height
    print('画布 %.2f x %.2f 英寸' % (SW / EMU, SH / EMU))
    print('页数 %d  母版 %d  版式 %d\n' % (
        len(prs.slides), len(prs.slide_masters),
        sum(len(m.slide_layouts) for m in prs.slide_masters)))

    sizes = Counter()
    colors = Counter()
    for s in prs.slides:
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if not r.text.strip():
                        continue
                    sizes[round(r.font.size.pt, 1) if r.font.size else None] += 1
                    try:
                        colors[str(r.font.color.rgb)] += 1
                    except Exception:
                        colors['theme/继承'] += 1

    total = sum(sizes.values()) or 1
    print('=== 字号分布（按次数）===')
    for k, v in sizes.most_common():
        print('  %-10s %5d  (%.0f%%)' % (
            k, v, 100 * v / total))

    print('\n=== 颜色分布 ===')
    tc = sum(colors.values()) or 1
    for k, v in colors.most_common():
        print('  %-12s %5d  (%.0f%%)' % (k, v, 100 * v / tc))

    print('\n=== 标题框几何 ===')
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if sh.has_text_frame and sh.name.startswith('标题') and sh.text_frame.text.strip():
                print('  页%-2d L=%.2f T=%.2f W=%.2f H=%.2f' % (
                    i, sh.left / EMU, sh.top / EMU,
                    (sh.width or 0) / EMU, (sh.height or 0) / EMU))
                break

    print('\n=== 正文框几何（前 15 个）===')
    n = 0
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if (sh.has_text_frame and sh.name.startswith(('文本框', '文本占位符'))
                    and sh.text_frame.text.strip()):
                print('  页%-2d T=%.2f H=%.2f 底=%.2f  字数%d' % (
                    i, sh.top / EMU, (sh.height or 0) / EMU,
                    (sh.top + (sh.height or 0)) / EMU,
                    len(sh.text_frame.text)))
                n += 1
                break
        if n >= 15:
            break
    return 0


if __name__ == '__main__':
    sys.exit(main())

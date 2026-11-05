# -*- coding: utf-8 -*-
"""
PPT Toolkit - 文本框容量测量

对 pptx 里每个可编辑文本框，用真实换行引擎算出：
  - 当前内容需要几行、多高
  - 目标是框里放得下，还是溢出

依赖 ppt-master 的 text_measure.py（真实 CJK 换行规则）。
未找到时降级为按字符宽度粗估，并能通过 PPT_MEASURE 环境变量指定路径。

用法：
    python measure_text.py <file.pptx> [--json out.json]
"""
import sys, os, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation
from pptx.oxml.ns import qn

EMU = 914400.0
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def find_measure():
    p = os.environ.get('PPT_MEASURE')
    if p and os.path.exists(p):
        return p
    cand = [
        os.path.join(os.path.dirname(ROOT), 'ppt-master',
                     'skills', 'ppt-master', 'scripts', 'text_measure.py'),
        r'C:\Users\ASUS\Documents\deepseek-harness\default-workspace\ppt-master\skills\ppt-master\scripts\text_measure.py',
    ]
    for c in cand:
        if os.path.exists(c):
            return c
    return None


MEASURE = find_measure()


def wrap_lines(text, max_pt, size, family='微软雅黑', spacing=1.5):
    """调用真实换行引擎；找不到时按 1 汉字 ≈ 1 em 粗估"""
    if MEASURE:
        r = subprocess.run(
            [sys.executable, MEASURE, 'wrap', text,
             '--max-width', '%.2f' % max_pt, '--x', '0',
             '--dy', '%.2f' % (size * spacing),
             '--size', '%.2f' % size, '--family', family, '--json'],
            capture_output=True)
        try:
            return json.loads(r.stdout.decode('utf-8', 'replace'))
        except Exception:
            pass
    # 降级：按字宽粗估（中文按 1.0em，英文按 0.55em）
    per_line = max(1, int(max_pt / size))
    n = max(1, -(-len(text) // per_line))
    return {'lines': ['x'] * n, 'widths': [max_pt] * n}


def main():
    args = sys.argv[1:]
    if not args:
        print('用法: python measure_text.py <file.pptx> [--json out.json]')
        return 2
    src = args[0]
    outp = args[args.index('--json') + 1] if '--json' in args else None

    prs = Presentation(src)
    rows = []
    print('%-4s %-14s %-6s %-6s %-8s %-8s %s' %
          ('页', '形状', '字号', '字数', '需高"', '框高"', '结论'))
    print('-' * 72)
    bad = 0
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if not sh.has_text_frame or not sh.text_frame.text.strip():
                continue
            if sh.width is None or sh.height is None:
                continue
            # 图形内的短标签（箭头/形状名）不参与溢出判定：
            # 它们通常单行居中显示，超出会被图形裁切而非撑破版面
            short_label = (sh.name.startswith(('右箭头', '左箭头', '上箭头', '下箭头',
                                              '箭头', '矩形', '圆角矩形', '椭圆',
                                              '直线', '连接符', '三角形', '菱形'))
                           and len(sh.text_frame.text.strip()) <= 12)
            if short_label:
                continue
            bodyPr = sh.text_frame._txBody.find(qn('a:bodyPr'))
            if bodyPr is not None and (
                    bodyPr.find(qn('a:spAutoFit')) is not None
                    or bodyPr.find(qn('a:normAutofit')) is not None):
                continue  # 自适应框，不会溢出
            # wrap=none 的框不自动换行，按单行计，不判断纵向溢出
            no_wrap = bodyPr is not None and bodyPr.get('wrap') == 'none'
            l = int(bodyPr.get('lIns', 91440)) if bodyPr is not None else 91440
            r_ = int(bodyPr.get('rIns', 91440)) if bodyPr is not None else 91440
            t = int(bodyPr.get('tIns', 45720)) if bodyPr is not None else 45720
            b = int(bodyPr.get('bIns', 45720)) if bodyPr is not None else 45720
            iw = (sh.width - l - r_) / EMU * 72
            ih = (sh.height - t - b) / EMU * 72
            if iw < 6 or ih < 6:
                continue

            pts = [x.font.size.pt for p in sh.text_frame.paragraphs
                   for x in p.runs if x.font.size]
            if not pts:
                continue
            size = max(pts)
            txt = sh.text_frame.text.replace('\n', ' ')
            # 读取段落真实行距；未显式设置时按 1.15 倍估算（PowerPoint 默认单倍≈1.2）
            ls = None
            for p in sh.text_frame.paragraphs:
                if p.line_spacing is not None:
                    try:
                        ls = float(p.line_spacing)
                    except Exception:
                        ls = None
                    break
            spacing = ls if ls and ls > 0.5 else 1.2
            if no_wrap:
                nlines = 1
            else:
                res = wrap_lines(txt, iw, size, spacing=spacing)
                nlines = len(res['lines'])
            need = nlines * size * spacing / 72
            box = ih / 72
            ok = need <= box * 1.05
            if not ok:
                bad += 1
            print('%-4d %-14s %-6.0f %-6d %-8.2f %-8.2f %s' % (
                i, sh.name[:14], size, len(txt), need, box,
                '✅' if ok else '❌ 溢出'))
            rows.append(dict(page=i, shape=sh.name, size=size, chars=len(txt),
                             lines=nlines, spacing=spacing,
                             need_h=round(need, 2), box_h=round(box, 2), fits=ok))

    print('\n共 %d 个固定文本框，溢出 %d 个' % (len(rows), bad))
    if outp:
        with open(outp, 'w', encoding='utf-8') as f:
            json.dump(rows, f, ensure_ascii=False, indent=1)
        print('明细已存:', outp)
    return 0 if bad == 0 else 1


if __name__ == '__main__':
    sys.exit(main())

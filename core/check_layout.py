# -*- coding: utf-8 -*-
"""
PPT Toolkit - 版面检查

渲染 pptx 的每一页，用像素测量判断版面是否有大片空洞。

判定项：
  A. 内容占高 < 55%   → 版面过空
  B. 底部空白 > 35%   → 下半页没用上
  C. 右侧越界 > 97%   → 内容压到 logo 区

环境变量（可选）：
  DSH_NODE      node.exe 路径（默认自动探测）
  LO_CLI        libreoffice-kit cli.js 路径（默认自动探测）
  PPT_BG       背景色 R,G,B（默认 6,26,66）
"""
import sys, os, json, subprocess, glob, shutil
sys.stdout.reconfigure(encoding='utf-8')
from PIL import Image
import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, 'out')

def _bg():
    v = os.environ.get('PPT_BG', '6,26,66')
    return np.array([int(x) for x in v.split(',')])

BG = _bg()
LOGO_TOP = int(os.environ.get('PPT_LOGO_TOP', '105'))
LOGO_RIGHT = float(os.environ.get('PPT_LOGO_RIGHT', '0.88'))
FILL_MIN = float(os.environ.get('PPT_FILL_MIN', '55'))
BOTTOM_MAX = float(os.environ.get('PPT_BOTTOM_MAX', '35'))
RIGHT_MAX = float(os.environ.get('PPT_RIGHT_MAX', '97'))


def find_node():
    p = os.environ.get('DSH_NODE')
    if p and os.path.exists(p):
        return p
    for c in [r'W:\ranjian\DeepseekHarness\resources\runtime\primary-runtime\dependencies\node\bin\node.exe',
              shutil.which('node') or '']:
        if c and os.path.exists(c):
            return c
    return 'node'


def find_cli():
    p = os.environ.get('LO_CLI')
    if p and os.path.exists(p):
        return p
    for c in [r'W:\ranjian\DeepseekHarness\resources\app.asar.unpacked\dsh\node_modules\@deepseek-ai\libreoffice-kit\lib\cli.js']:
        if os.path.exists(c):
            return c
    raise SystemExit('找不到 libreoffice-kit CLI，请设置环境变量 LO_CLI')


def render(src, pages=None):
    os.makedirs(OUT, exist_ok=True)
    tmp = os.path.join(OUT, '_render_tmp.pptx')
    subprocess.run([sys.executable, os.path.join(HERE, 'strip_fonts.py'), src, tmp],
                   check=True, capture_output=True)
    outdir = os.path.join(OUT, 'render')
    if os.path.exists(outdir):
        shutil.rmtree(outdir)
    cmd = [find_node(), find_cli(), 'render', '--input', tmp, '--output-dir', outdir]
    if pages:
        cmd += ['--pages', pages]
    r = subprocess.run(cmd, capture_output=True)
    if r.returncode != 0:
        raise SystemExit('渲染失败: ' + r.stderr.decode('utf-8', 'replace')[:300])
    if os.path.exists(tmp):
        os.remove(tmp)
    return outdir


def analyse(img_path):
    im = Image.open(img_path).convert('RGB')
    a = np.asarray(im).astype(int)
    h, w, _ = a.shape
    sub = a[LOGO_TOP:, :int(w * LOGO_RIGHT)]
    mask = np.abs(sub - BG).sum(axis=2) > 60
    ys = np.where(mask.sum(axis=1) > 3)[0]
    xs = np.where(mask.sum(axis=0) > 3)[0]
    if len(ys) == 0:
        return dict(empty=True, fill=0, bottom=100, top=0, left=0, right=0)
    y0, y1 = ys[0] + LOGO_TOP, ys[-1] + LOGO_TOP
    return dict(empty=False,
                top=round(100 * y0 / h, 1),
                bottom=round(100 * (h - y1) / h, 1),
                fill=round(100 * (y1 - y0) / h, 1),
                left=round(100 * xs[0] / w, 1),
                right=round(100 * xs[-1] / w, 1))


def main():
    if len(sys.argv) < 2:
        print('用法: python check_layout.py <file.pptx> [-p 1,2,3]')
        return 2
    src = sys.argv[1]
    pages = sys.argv[sys.argv.index('-p') + 1] if '-p' in sys.argv else None

    print('渲染中...')
    outdir = render(src, pages)
    imgs = sorted(glob.glob(os.path.join(outdir, 'page-*.png')))
    print('共 %d 页\n' % len(imgs))

    print('%-4s %-8s %-10s %-10s %-10s %s' %
          ('页', '顶部%', '底部空%', '占高%', '横向%', '判定'))
    print('-' * 74)
    problems = []
    for i, p in enumerate(imgs, 1):
        m = analyse(p)
        flags = []
        if m['empty']:
            flags.append('空白页')
        else:
            if m['fill'] < FILL_MIN:
                flags.append('内容过少')
            if m['bottom'] > BOTTOM_MAX:
                flags.append('底部大片空')
            if m['right'] > RIGHT_MAX:
                flags.append('右侧越界')
        if flags:
            problems.append((i, m, flags))
        print('%-4d %-8s %-10s %-10s %-10s %s' % (
            i, m.get('top', '-'), m.get('bottom', '-'), m.get('fill', '-'),
            '%s~%s' % (m.get('left', '-'), m.get('right', '-')),
            '、'.join(flags) if flags else 'OK'))

    print('\n' + '=' * 74)
    if not problems:
        print('全部通过，无版面问题')
    else:
        print('问题页 %d 个：' % len(problems))
        for i, m, f in problems:
            print('  第%-2d页 占高 %-5s%% 底部空 %-5s%% → %s' %
                  (i, m.get('fill'), m.get('bottom'), '、'.join(f)))
    return 0 if not problems else 1


if __name__ == '__main__':
    sys.exit(main())

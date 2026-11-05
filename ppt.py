# -*- coding: utf-8 -*-
"""
PPT Toolkit - 统一入口

用法：
    python ppt.py audit   <file.pptx>              # 结构+字体交付检查
    python ppt.py design  <file.pptx>              # 提取设计系统（字号/颜色/几何）
    python ppt.py measure <file.pptx>              # 逐文本框测量容量与溢出
    python ppt.py layout  <file.pptx>              # 渲染并像素级检查版面空洞
    python ppt.py render  <file.pptx> -p 1,2,3     # 渲染指定页为 PNG
    python ppt.py text    "<文本>" --size 14       # 单段文本宽度测量

依赖：
    python-pptx, Pillow, numpy, lxml
    LibreOffice Kit（随 DSH 安装）+ ppt-master 仓库（用于 text_measure / delivery_check）
"""
import sys, os, subprocess, argparse
sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
CORE = os.path.join(HERE, 'core')
WS = os.path.dirname(HERE)

# 外部依赖：ppt-master 仓库（整仓保留，含 attribution 文件）
PM_ROOT = os.path.join(WS, 'ppt-master', 'skills', 'ppt-master')
PM_SCRIPTS = os.path.join(PM_ROOT, 'scripts')
PM_TEXT = os.path.join(PM_SCRIPTS, 'text_measure.py')
PM_DELIVERY = os.path.join(PM_SCRIPTS, 'pptx_delivery_check.py')

# LibreOffice Kit（DSH 自带）
NODE = r'W:\ranjian\DeepseekHarness\resources\runtime\primary-runtime\dependencies\node\bin\node.exe'
LO_CLI = r'W:\ranjian\DeepseekHarness\resources\app.asar.unpacked\dsh\node_modules\@deepseek-ai\libreoffice-kit\lib\cli.js'


def _run(script, args):
    p = subprocess.run([sys.executable, os.path.join(CORE, script)] + args)
    return p.returncode


def cmd_audit(args):
    """交付检查：结构完整性 + 字体可移植性"""
    if not os.path.exists(PM_DELIVERY):
        print('缺少 ppt-master，请先把仓库放到 %s' % os.path.dirname(PM_ROOT))
        return 2
    return subprocess.run([sys.executable, PM_DELIVERY, args.file]).returncode


def cmd_design(args):
    return subprocess.run([sys.executable,
                           os.path.join(CORE, 'extract_design.py'),
                           args.file]).returncode


def cmd_measure(args):
    cmd = [sys.executable, os.path.join(CORE, 'measure_text.py'), args.file]
    if getattr(args, 'json', None):
        cmd += ['--json', args.json]
    return subprocess.run(cmd).returncode


def cmd_layout(args):
    """渲染 + 像素级版面检查（调用 core/check_layout.py）"""
    cmd = [sys.executable, os.path.join(CORE, 'check_layout.py'), args.file]
    if getattr(args, 'pages', None):
        cmd += ['-p', args.pages]
    return subprocess.run(cmd).returncode


def cmd_render(args):
    """只渲染，不做判定"""
    strip = os.path.join(CORE, 'strip_fonts.py')
    os.makedirs(os.path.join(HERE, 'out'), exist_ok=True)
    tmp = os.path.join(HERE, 'out', '_render_tmp.pptx')
    subprocess.run([sys.executable, strip, args.file, tmp], check=True)
    outdir = os.path.join(HERE, 'out', 'render')
    if os.path.exists(outdir):
        import shutil
        shutil.rmtree(outdir)
    cmd = [find_node(), find_cli(), 'render', '--input', tmp, '--output-dir', outdir]
    if getattr(args, 'pages', None):
        cmd += ['--pages', args.pages]
    r = subprocess.run(cmd, capture_output=True)
    if r.returncode != 0:
        print('渲染失败:', r.stderr.decode('utf-8', 'replace')[:300])
        return 1
    print('渲染完成: %s' % outdir)
    return 0


def cmd_text(args):
    sys.path.insert(0, PM_SCRIPTS)
    r = subprocess.run([sys.executable, PM_TEXT, 'measure', args.text,
                        '--size', str(args.size),
                        '--family', args.family])
    return r.returncode


def main():
    ap = argparse.ArgumentParser(description='PPT Toolkit')
    sub = ap.add_subparsers(dest='cmd', required=True)

    for name, fn, help_text in [
        ('audit', cmd_audit, '结构+字体交付检查'),
        ('design', cmd_design, '提取设计系统'),
        ('measure', cmd_measure, '测量文本框容量'),
        ('layout', cmd_layout, '渲染并检查版面空洞'),
        ('render', cmd_render, '渲染指定页'),
    ]:
        p = sub.add_parser(name, help=help_text)
        p.add_argument('file', help='.pptx 路径')
        if name in ('layout', 'render'):
            p.add_argument('-p', '--pages', help='如 1,2,3，留空则全部')
        if name == 'measure':
            p.add_argument('--json', help='把明细写到这个 json 文件')
        p.set_defaults(func=fn)

    pt = sub.add_parser('text', help='单段文本宽度测量')
    pt.add_argument('text')
    pt.add_argument('--size', type=float, default=14)
    pt.add_argument('--family', default='微软雅黑')
    pt.set_defaults(func=cmd_text)

    args = ap.parse_args()
    sys.exit(args.func(args) or 0)


if __name__ == '__main__':
    main()

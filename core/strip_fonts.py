# -*- coding: utf-8 -*-
"""剥离嵌入字体生成渲染副本（避免 CLI 输入上限）"""
import sys, os, zipfile, re
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation

src = sys.argv[1] if len(sys.argv) > 1 else '培训课件-EMR-精修版.pptx'
dst = sys.argv[2] if len(sys.argv) > 2 else '_render2.pptx'
if os.path.exists(dst):
    os.remove(dst)

with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
    for n in zin.namelist():
        if n.startswith('ppt/fonts/'):
            continue
        d = zin.read(n)
        if n == 'ppt/presentation.xml':
            t = d.decode('utf-8')
            t = re.sub(r'<p:embeddedFontLst>.*?</p:embeddedFontLst>', '', t, flags=re.S)
            t = t.replace(' embedTrueTypeFonts="1"', '').replace(' saveSubsetFonts="1"', '')
            d = t.encode('utf-8')
        elif n == '[Content_Types].xml':
            t = d.decode('utf-8')
            t = re.sub(r'<Override[^>]*PartName="/ppt/fonts/[^>]*/>', '', t)
            d = t.encode('utf-8')
        zout.writestr(n, d)

print('副本: %s  %.1f MB  页数 %d' %
      (dst, os.path.getsize(dst) / 1024 / 1024, len(Presentation(dst).slides)))

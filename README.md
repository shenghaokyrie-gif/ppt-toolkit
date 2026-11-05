# ppt-toolkit

[English](#english) | 中文

把「改 PPT」从凭感觉改成可测量。这套工具解决的核心问题是：**你看不到渲染结果的时候，怎么知道版面有没有坏**。

## 为什么需要它

做 PPT 自动化时最容易犯的错，是只改文字不看版面。本仓库把踩过的坑固化成检查项：

| 错误 | 后果 | 对应工具 |
|---|---|---|
| 只改文字、不动几何 | 新增页只占 52% 画布，下半屏全空 | `layout` |
| 字号全用继承 | 6 行字一样大，没有层级 | `design` + `measure` |
| 强调色当正文色 | 整页全黄，没有重点 | `design` |
| 假设字体已装 | 思源黑体缺失，文字回退撑破文本框 | `audit` |
| 用 PIL 近似渲染 | 判断错、整本返工 | LibreOffice 真实渲染 |

## 安装

```bash
pip install python-pptx pillow numpy lxml
```

还需要两样外部依赖：

1. **LibreOffice Kit** —— 用于真实渲染。若你装了 DeepSeek Harness，它随附在 `app.asar.unpacked` 里；否则安装 LibreOffice 并设置 `LO_CLI`。
2. **ppt-master** —— 提供 `text_measure.py`（真实 CJK 换行）和 `pptx_delivery_check.py`（交付检查）。
   放到与本仓库同级的 `../ppt-master/skills/ppt-master/`。

> ppt-master 不能只拷脚本：`attribution_guard.py` 会校验 `LICENSE`、`SPONSORS.md`、`SPONSORS_CN.md`
> 的存在与哈希，缺一个就 `exit 78`。但**只需 `skills/ppt-master/` 子目录（36.5 MB）**，不需要整个仓库。

## 用法

```bash
# 1. 摸清一个既有模板的视觉规范
python ppt.py design deck.pptx

# 2. 检查结构完整性与字体可移植性
python ppt.py audit deck.pptx

# 3. 渲染并检查版面空洞
python ppt.py layout deck.pptx
python ppt.py layout deck.pptx -p 4,21,22

# 4. 测量文本框容量，找溢出
python ppt.py measure deck.pptx --json report.json

# 5. 单段文本宽度
python ppt.py text "培训结束后，你应当能够独立完成" --size 14
```

## 判定阈值

`layout` 的阈值按模板可调，默认值来自一套深蓝背景 + 顶部 logo 的 16:9 模板实测基线：

| 指标 | 默认 | 环境变量 |
|---|---|---|
| 内容占高下限 | 55% | `PPT_FILL_MIN` |
| 底部空白上限 | 35% | `PPT_BOTTOM_MAX` |
| 右侧越界上限 | 97% | `PPT_RIGHT_MAX` |
| 背景色 | `6,26,66` | `PPT_BG` |
| 顶部 logo 带高度(px) | 105 | `PPT_LOGO_TOP` |
| 右侧 logo 起点比例 | 0.88 | `PPT_LOGO_RIGHT` |

正常内容页的实测基线：占高 80~90%，底部余 1~10%。

## 目录

```
ppt-toolkit/
├── ppt.py                  统一入口
├── core/
│   ├── strip_fonts.py      剥离嵌入字体（大文件才能渲染）
│   ├── check_layout.py     渲染 + 像素级版面检查
│   ├── extract_design.py   提取字号/颜色/几何规范
│   └── measure_text.py     文本框容量测量
├── examples/               示例 pptx
└── out/render/             渲染输出
```

## 经验条目

1. **改文字前先量几何**。文本框高度决定能放几行，不是字号决定的。
2. **章节页不能只放大字号**。3 个要点撑不满 16:9，要用卡片/网格等结构元素。
3. **强调色只做强调**。原件里 `FFC000` 只占 4%，正文 96% 是白色或主题色。
4. **标题框要够宽**。`培训议程` 在 3.18" 下会折成两行，6.5" 才安全。
5. **渲染前剥字体**。65 MB（含 57 MB 嵌入字体）会被渲染 CLI 拒绝，剥掉后 12 MB。
6. **嵌入字体不影响交付**，只阻碍渲染。交付文件里保留嵌入字体是对的。
7. **图形内短标签不参与溢出判定**。箭头、形状里的 4-6 字标签超出会被裁切，不会撑破版面。

## English

Turn "editing a deck" from guesswork into measurement. The core problem: **when you can't see the rendered result, how do you know the layout isn't broken?**

Five commands: `design` (extract typography rules), `audit` (structure + font portability),
`layout` (render and pixel-measure empty space), `measure` (text overflow), `text` (single-line width).

Requires LibreOffice Kit for real rendering, and [ppt-master](https://github.com/hugohe3/ppt-master)
for the CJK line-breaking engine and delivery check.

## License

MIT

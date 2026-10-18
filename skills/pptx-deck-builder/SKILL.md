---
name: pptx-deck-builder
description: 用 python-pptx 生成不带「AI 味」的演示文稿。当需要创建 .pptx 文件、生成 PPT、做演示文稿、制作汇报材料，且希望避免统一背景/蓝紫渐变/每页三张卡片这类模板感时使用。提供 deck_builder.py 渲染库（含 A/B/C 三区母版、巨型数字锚点、非对称版式、5 套预置色板），以及一份把「叙事判断」前置的设计规范。适用于 Codex / Claude Code / 任何能执行 Python 的环境。不依赖任何专有 PPT 工具链。
---

# PPTX Deck Builder

## 这个 skill 解决什么

大多数 AI 生成的 PPT 之所以"有 AI 味"，不是因为审美差，而是因为**设计判断只发生了一次**——拿到主题就套模板、逐页填字，最终所有页面视觉重量相同、背景统一、版式重复。

本 skill 把两种能力打包在一起：

1. `scripts/deck_builder.py` —— 一个把设计约束**内建进 API** 的渲染库，用 python-pptx 生成真正的 .pptx
2. `references/design-rules.md` —— 一份把「每页要重新判断一次」落成硬规则的设计规范

**核心工作流：先做叙事判断（每页讲什么、什么角色），再调渲染 API。顺序反了就会出 AI 味。**

## 环境准备

```bash
pip install python-pptx Pillow
```

验证：
```bash
python scripts/deck_builder.py demo.pptx   # 生成 3 页自检文件
```

无报错即为可用。python-pptx >= 0.6.21。

## 快速开始

```python
import sys
sys.path.insert(0, "scripts")
from deck_builder import Builder, Palette

b = Builder(Palette.academic(), total_pages=12, footer_text="课题汇报 · 2026")

# 封面：左文右图
b.cover("强对流天气的雷达特征分析",
        "基于 2023–2025 年华南地区观测资料\n识别三类典型回波形态",
        eyebrow="开题报告", tag="2026 春",
        stats=[("3", "类回波形态"), ("1,247", "个样本")],
        footnote="数据来源：中国气象局雷达数据集",
        image="assets/cover.png")

# 章节扉页
b.section("01", "研究背景", "为什么要研究强对流天气的雷达识别。", chapter="CHAPTER 01")

# 内容页
s = b._slide()
b.header(s, "三类典型回波形态", chapter="CHAPTER 01", page_no=4)
b.big_number(s, 64, 156, "1,247", "个", "有效样本数")
b.bar_row(s, 640, 160, 560, "线状回波", "612 个", 0.49, growth="49.1%")
b.bullet_list(s, 640, 240, 560, [
    ("钩状回波", "与中气旋强相关，建议优先关注。"),
    ("弓形回波", "对应强风带，需结合径向速度场判断。"),
])

b.save("out.pptx")
```

## 可用 API 速查

完整签名见 `scripts/deck_builder.py` 的 docstring。坐标单位统一是**像素**（1280×720 画布）。

| 方法 | 用途 |
| :-- | :-- |
| `cover(title, sub, eyebrow, tag, stats, footnote, image)` | 封面。左文右图，`stats` 为 `[(值, 标签)]` |
| `section(number, title, desc, chapter)` | 章节扉页。巨型描边数字 |
| `header(slide, title, chapter, page_no)` | A 区标题块（内容页必调，保证母版一致） |
| `big_number(slide, x, y, num, unit, caption)` | 巨型锚点数字。**每页至少一个** |
| `card(slide, x, y, w, h, title, body, tag, head_color)` | 圆角卡片 |
| `bar_row(slide, x, y, w, label, value, ratio, growth)` | 数据条行 |
| `bullet_list(slide, x, y, w, items)` | 编号条目列表（替代等宽卡片横排） |
| `timeline(slide, x, y, w, entries)` | 竖向时间轴，`entries` 含 `(年份, 标题, 描述, 颜色)` |
| `pic(slide, path, x, y, w, h)` | 插图，自动按 cover 裁切不变形 |
| `text(slide, x, y, w, h, content, ...)` | 通用文本框，`\n` 分段 |

## 预置色板

`Palette.academic()` / `.research()` / `.cream_pink()` / `.weather()` / `.business()`

各自的适用场景、完整色值、以及为什么这样配，见 `references/design-rules.md`。

也可自定义：
```python
Palette(bg="#FFFFFF", accent="#2563EB", accent2="#0EA5E9", accent3="#1E293B")
```

## 硬规则（违反即出 AI 味）

生成前务必读 `references/design-rules.md`。最容易踩的五条：

1. **不做统一背景** —— 至少 3 页的色块占比与其他页显著不同
2. **不用等宽卡片横排超过 2 次** —— 这是最高发的 AI 排版特征
3. **每页必须有视觉锚点** —— 巨型数字（≥48px）、大图（占内容区≥40%）或大色块
4. **相邻两页版式必须不同** —— 禁止连续同构
5. **关键数字后必须跟判断** —— 写"海外收入 162.68 亿"不够，要接"——但增速来自低基数"

## 与 ppt-brief 配合

本仓库另含 `ppt-brief` skill，负责动手前对齐需求（主题/页数/色板/质感关键词/元素清单/负向约束）。

**推荐顺序：`ppt-brief` 出需求简报 → 本 skill 按简报生成。**

若已在别处确定了这 6 项约束，可跳过 ppt-brief 直接使用本 skill。

## 已知限制

- 渐变填充通过纯色近似（python-pptx 的 gradient API 不稳定），如需真渐变需手工改 XML
- 不生成图片；配图需自行准备并放入 `assets/`，或用外部图像生成能力
- 字体依赖运行环境已安装。中文字体回落到默认宋体时，需在 `Builder(...)` 里指定 `title_font`
- 深色底色板（`Palette.research()`）在投影仪环境下对比度可能不足，正式场合建议先试投

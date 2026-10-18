# ppt-skills

两个配合使用的 skill，用来生成**不带「AI 味」**的演示文稿。

## 为什么会有「AI 味」

大多数 AI 生成的 PPT 之所以一眼能看出是机器做的，根因不是审美差，而是**设计判断的密度太低**：

拿到主题 → 立刻套模板 → 逐页填字。整个过程中，设计判断只发生了 **1 次**，于是 20 页共用同一套视觉决策——统一背景、每页三张等宽卡片、蓝紫科技渐变。

本仓库的思路是把判断密度从「1 次」提到「每页 1 次」，并且**把约束问死**（具体色值、明确禁忌），逼模型离开统计平均答案。

## 两个 skill 的分工

| Skill | 职责 | 产出 |
| :-- | :-- | :-- |
| [`ppt-brief`](skills/ppt-brief/) | 动手前对齐需求：一次性问齐 6 项约束 | 一份需求简报 `.md` |
| [`pptx-deck-builder`](skills/pptx-deck-builder/) | 按约束生成 .pptx | 真正的 `.pptx` 文件 |

**推荐顺序：先 `ppt-brief` 定需求，再 `pptx-deck-builder` 生成。**

若 6 项约束已在别处确定，可跳过第一个直接用第二个。

## ppt-brief · 需求对齐

在开始制作前问齐这 6 项：

1. **主题** —— 讲什么？给谁看？什么场合？
2. **页数** —— 决定视觉节奏与章节切分
3. **色板** —— 具体 hex 色值，不是"好看的颜色"
4. **质感关键词** —— 3–5 个词
5. **元素清单** —— 必须出现什么
6. **负向约束** —— 明确禁止什么

第 6 项最容易被忽略，但对消除 AI 味最有效：**模型的默认行为是补全所有它认为合理的东西**，只写"要什么"，它会在满足要求后顺手加一堆元素。负向约束是在直接删掉这些默认选项。

用户答不上来时，skill 内置 5 套完整方案（含全部 6 项默认值）供挑选，不会卡住。

## pptx-deck-builder · 生成

基于 python-pptx，不依赖任何专有工具链。

```bash
pip install python-pptx Pillow
python skills/pptx-deck-builder/scripts/deck_builder.py demo.pptx   # 自检
```

```python
import sys
sys.path.insert(0, "skills/pptx-deck-builder/scripts")
from deck_builder import Builder, Palette

b = Builder(Palette.academic(), total_pages=12, footer_text="课题汇报")
b.cover("强对流天气的雷达特征分析", "基于 2023–2025 年华南观测资料",
        eyebrow="开题报告", stats=[("3", "类回波形态"), ("1247", "个样本")])
s = b._slide()
b.header(s, "三类典型回波形态", chapter="CHAPTER 01", page_no=3)
b.big_number(s, 64, 156, "1,247", "个", "有效样本数")
b.bar_row(s, 640, 160, 560, "线状回波", "612 个", 0.49, growth="49.1%")
b.save("out.pptx")
```

内置能力：

- **A/B/C 三区母版** —— 标题/内容/页脚位置全篇固定
- **视觉锚点 API** —— `big_number()` 强制每页有一个焦点
- **非对称版式** —— `timeline()` / `bar_row()` / `bullet_list()` 等，避免等宽卡片堆叠
- **5 套预置色板** —— 学术答辩 / 科研数据 / 品牌商业 / 气象科普 / 通用汇报，每套含完整适用说明与禁忌
- **中文友好** —— 自动设置东亚字体（`a:ea`），避免回落到宋体

## 目录结构

```
skills/
├── ppt-brief/
│   ├── SKILL.md
│   └── references/palettes.md          # 5 套色板 + 通用负向约束库
└── pptx-deck-builder/
    ├── SKILL.md
    ├── references/design-rules.md      # 设计规范（母版/节奏/色彩/版式/自检）
    └── scripts/deck_builder.py         # 渲染库
```

`demo.pptx` 未入库——它是自检产物，运行下面这句即可重新生成：

```bash
python skills/pptx-deck-builder/scripts/deck_builder.py demo.pptx
```

## 安装

**Claude Code / Codex**：把 `skills/` 下的目录复制到你的 skill 目录，或在本仓库内直接引用。

**手动使用**：只要 Python 能跑就能用，无需任何 agent 框架——`deck_builder.py` 就是个普通库。

## 已知限制

- 渐变填充用纯色近似（python-pptx 的 gradient API 不稳定）
- 不生成图片；配图需自备，或用外部图像生成能力
- 字体依赖运行环境；环境缺字体时可传 `title_font` 指定
- 深色底在投影仪上对比度可能不足，正式场合建议先试投

## License

MIT

# 配图工作流

渲染库只排版，**不生成图片**。但配图是"高级感"的主要来源之一——本文件定义如何编排外部生图能力，把图片做得和版面风格一致。

> 核心原则：**能力由运行环境提供，流程知识由本文件提供。**
> skill 不内置图像 API，而是明确"什么时候生图、prompt 怎么写、怎么验收、怎么回填"。

---

## 1. 先判断：这页的图该不该生

不是每页都要图。按**叙事角色**判断，而不是按"好不好看"判断：

| 图的角色 | 典型页面 | 处理方式 |
| :-- | :-- | :-- |
| **具象场景** | 封面、产品展示、门店、人物、空间 | ✅ 必须生图。用几何图形冒充会一眼假 |
| **证据** | 数据图、流程图、对比图 | ⚠️ 优先用代码画（`bar_row` / `timeline`），只有需要真实画面时才生图 |
| **结构表达** | 时间轴、矩阵、飞轮、地图示意、概率条 | ❌ **禁止生图**，用 SVG / 渲染库方法画。生了反而变糊 |

**判断口诀**：如果这页的图画错了内容，读者还能看懂论点 → 它是装饰，可以省；如果读者必须"看见"某样东西才能理解 → 必须生图。

---

## 2. 生图前：先出配图清单

用 `images_needed()` 生成清单，**先规划再生成**，避免边做边想要什么图：

```python
plan = b.images_needed([
    ("cover_expo.png",  "780×720", "潮玩展览空间，奶白展台，柔和射灯，柔和侧光", "P01 封面"),
    ("ip_character.png","620×500", "单个尖耳玩具角色特写，浅景深，暖调",        "P07 角色页"),
    ("blindbox.png",    "560×470", "成排未拆封方形小盒，暖光，奶白背景",        "P09 机制页"),
])
print(plan)
```

把输出写进 `assets/IMAGES.md`，作为交付附件——使用者一眼知道哪几张图待补。

---

## 3. Prompt 配方（决定成败的一步）

**核心要求：全篇用同一组风格词，禁止混搭。** 摄影 + 插画 + 3D 混用是最常见的失败原因。

### 标准骨架（按顺序拼）

```
[主体与场景] + [构图] + [材质/风格] + [色彩限制] + [光线] + [技术参数] + [禁忌]
```

### 各风格的标准风格词

| 风格 | 风格词串（直接复用） |
| :-- | :-- |
| **编辑摄影**（品牌/产品/空间） | `editorial photography, cream white background, soft diffused side light, shallow depth of field, warm tone, minimal, high-end product editorial` |
| **科研/技术** | `clean technical illustration, isometric, flat shading, high contrast, no gradients, precise geometry` |
| **学术/示意** | `minimal line diagram, monochrome with single accent color, generous whitespace, no texture` |
| **写实场景**（人物/城市） | `documentary photography, natural light, muted tones, shallow depth of field, candid` |

### 必带的禁忌词（漏了就会出 AI 味）

```
no text, no logo, no watermark, no faces, no humans
```

- `no text` —— 图片自带文字会和版面文字打架，且常常是乱码
- `no logo` / `no watermark` —— 侵权风险 + 破坏质感
- `no faces` / `no humans` —— 除非明确要人物，否则人脸最容易暴露 AI 感

### 色彩必须锚定到色板

不要让模型自己选色。把色板 hex 写进 prompt：

```
Color palette strictly limited to: cream white #F7F3EE, blush pink #F4A8C4,
soft coral red #E94752, pale aqua #64C9D4. No green, no purple, no other hues.
```

**实测教训**：不锚定色彩时，模型会自行引入薄荷绿、玫瑰紫等"看起来高级"但破坏色板的颜色——这类图必须重生。

---

## 4. 生图后：逐张验收（不可跳过）

生成的文件名通常是 prompt 前 30 字符 + 时间戳，**顺序常与请求错位**。必须：

1. **逐张打开看真实内容**，不能按时间戳盲映射
2. 核对四项：

| 检查项 | 不合格的表现 | 处理 |
| :-- | :-- | :-- |
| 内容正确 | 画的是别的场景 | 重生 |
| **材质统一** | 同批图里一张哑光陶土、一张亮面树脂 | **重生**（最容易忽略） |
| 色彩落板 | 出现色板外的颜色 | 重生并在 prompt 强调 "strictly limited to" |
| 比例可用 | 主体被裁掉 / 留白失衡 | 裁切调整，或重生 |

3. 重命名成英文短名（`cover_expo.png`），**禁止中文、空格、特殊字符**
4. 删除原名临时文件，避免 assets 目录堆积

**止损规则**：同一张图**最多重试 2 次**。仍不合格就换方案——改用 SVG 抽象表达，或这页干脆不放图（用巨型数字撑版面）。无休止重试只会浪费时间。

---

## 5. 落地时的三个技术细节

### ① 尺寸按最终显示比例给

不要把 1:1 的图塞进 16:9 的位置。按实际显示区域算：
- 封面右侧全高：`780×720`（约 1.08:1）
- 上半横幅：`1150×290`（约 4:1，**必须特别指定 ultra-wide**）
- 方形卡片位：`540×420`（约 1.29:1）

比例差太多时，用 `pic()` 的 cover 裁切会切掉主体。

### ② 生成图常带水印，需要裁掉

多数生图服务会在右下角加水印。用 Pillow 裁掉底部约 8–9%：

```python
from PIL import Image
im = Image.open(src).convert("RGB")
w, h = im.size
im.crop((0, 0, w, int(h * 0.91))).save(dst, "PNG", optimize=True)
```

**这一步要逐个文件处理，不要写通配符批量删除**。

### ③ 图片放进 `assets/` 后用相对路径引用

```python
b.pic(s, "assets/cover_expo.png", x, y, w, h)   # ✅ 相对路径
```

**禁止** `http://` 在线地址（构建后图会丢）和 `image://` 占位符。

---

## 6. 素材未就绪时：用占位框而不是空着

```python
b.pic_placeholder(s, 700, 160, 520, 320,
                  hint="此页需要门店外景照片",
                  spec="建议 16:9，暖色调，夜景橱窗")
```

占位框会在 PPT 里画出虚线框和说明，交付后使用者一眼看到"这里待补图"。补图时把这一行换成 `b.pic(...)` 即可，坐标不用改。

---

## 7. 完整流程（一句话版）

```
① 判断图属于「具象场景/证据/结构」→ 只有具象场景必须生图
② 用 images_needed() 出清单，写进 assets/IMAGES.md
③ 按标准骨架写 prompt，全篇复用同一组风格词 + 色板 hex + 禁忌词
④ 逐张验收内容/材质/色彩/比例，不合格最多重试 2 次
⑤ 裁掉水印（底部 8–9%），重命名为英文短名
⑥ 放进 assets/，用相对路径 b.pic() 引用
```

---

## 附：一个真实案例的 9 张图规划

品牌分析类 PPT（20 页），实际生图 9 张：

| 文件 | 尺寸 | 角色 | 内容 |
| :-- | :-- | :-- | :-- |
| `cover_expo.png` | 780×720 | 场域 | 潮玩展览空间，展台阵列 |
| `brand_store.png` | 1536×931 | 证据 | 明亮门店内景，陈列墙 |
| `ip_figures.png` | 1536×931 | 论点 | 多个角色列队，展示 IP 矩阵 |
| `ip_character.png` | 1024×1397 | 论点 | 单角色特写 |
| `blindbox.png` | 1536×931 | 论点 | 未拆封盒阵列 |
| `users_lifestyle.png` | 1536×931 | 场景 | 用户在房间整理收藏 |
| `channel_store.png` | 1536×931 | 证据 | 门店外立面夜景 |
| `overseas_city.png` | 1536×931 | 证据 | 海外街头排队 |
| `ending_hall.png` | 1536×931 | 收束 | 空旷展厅，远处光晕 |

其中 `ip_figures` 和 `ip_character` **各重试过 2 次**（一次材质偏陶土、一次色彩偏薄荷绿），说明"材质统一 + 色彩锚定"是最高频的两个失败点。

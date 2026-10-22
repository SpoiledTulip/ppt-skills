# -*- coding: utf-8 -*-
"""
deck_builder.py —— 用 python-pptx 构建「不带 AI 味」的演示文稿

把以下设计原则内建进 API，而不是留给调用者自觉：
  1. A/B/C 三区母版：标题块 24–124px / 内容区 124–656px / 页脚 656–696px
     全篇位置完全一致——这是"专业"与"模板感"的共同来源，关键在一致性
  2. 每页必须有一个视觉锚点（巨型数字 / 大图 / 大色块），不允许所有元素趋近中间大小
  3. 内容页 60% 以上用非对称版式（宽窄比 55:45 / 60:40 / 62:38）
  4. 色板 ≤4 色；强调色面积 ≤10%，hero 页可放宽到 18%
  5. 字号阶梯固定；巨型数字必须换字体（与正文区分）

坐标一律用「像素」，内部自动换算 EMU（1280×720 @96dpi）。
所有 API 的 x/y/w/h 单位都是 px，可直接对照设计稿。

依赖：python-pptx>=0.6.21
用法：
    from deck_builder import Builder, Palette
    b = Builder(Palette.cream_pink())
    s = b.content_slide("一超多强", chapter="CHAPTER 02")
    b.big_number(s, 64, 160, "141.6", "亿元", "占公司总营收 38%")
    b.save("out.pptx")
"""

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

CANVAS_W = 1280
CANVAS_H = 720
PAD_X = 64          # 左右安全边距
ZONE_A_H = 100      # 标题块高度
ZONE_C_H = 40       # 页脚高度
MARGIN_TOP = 24


def px(v):
    """像素 → EMU"""
    return Emu(int(round(v * 9525)))


def _rgb(hexstr):
    hexstr = hexstr.lstrip('#')
    return RGBColor(int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16))


def _relative_luminance(hexstr):
    """感知亮度（0–1）。用于自动选前景色，避免深底深字。"""
    h = hexstr.lstrip('#')
    rgb = [int(h[i:i + 2], 16) / 255.0 for i in (0, 2, 4)]
    # sRGB → 线性
    lin = [(c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4) for c in rgb]
    return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2]


def on_color(bg_hex, dark="#1E1A17", light="#FFFFFF"):
    """给定背景色，返回可读的前景色（亮底给深字，暗底给白字）。"""
    return dark if _relative_luminance(bg_hex) > 0.45 else light


def _set_font(run, name):
    """同时设置西文与东亚字体——中文若不设 ea typeface 会回落到默认宋体"""
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = parse_xml('<a:ea xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>')
        rPr.append(ea)
    ea.set('typeface', name)


class Palette:
    """四色板 + 派生中性色。强调色不要超过两个。"""

    def __init__(self, bg, accent, accent2, accent3,
                 text="#2A2320", text2="#6B5F58", text3="#A89C93",
                 line="#E8DFD6", card="#FFFFFF"):
        self.bg = bg
        self.accent = accent      # 主强调色：巨型数字、焦点元素
        self.accent2 = accent2    # 次强调色：第二系列、分区
        self.accent3 = accent3    # 第三色：仅用于点缀
        self.text = text
        self.text2 = text2
        self.text3 = text3
        self.line = line
        self.card = card

    # ---- 预置色板（对应 ppt-brief/references/palettes.md）----
    @classmethod
    def academic(cls):
        """学术答辩：白底 + 深蓝，公式图表友好"""
        return cls(bg="#FFFFFF", accent="#2563EB", accent2="#0EA5E9", accent3="#1E293B",
                   text="#1E293B", text2="#475569", text3="#94A3B8", line="#E2E8F0")

    @classmethod
    def research(cls):
        """科研数据：深底 + 亮色，图表对比度最高"""
        return cls(bg="#0F172A", accent="#38BDF8", accent2="#34D399", accent3="#F8FAFC",
                   text="#F8FAFC", text2="#CBD5E1", text3="#94A3B8",
                   line="#1E293B", card="#1E293B")

    @classmethod
    def cream_pink(cls):
        """品牌商业：奶油白 + 粉红蓝（泡泡玛特那份用的就是这套）"""
        return cls(bg="#F7F3EE", accent="#E94752", accent2="#64C9D4", accent3="#F4A8C4",
                   text="#2A2320", text2="#6B5F58", text3="#A89C93", line="#E8DFD6")

    @classmethod
    def weather(cls):
        """气象科普：浅蓝底，教室投影友好"""
        return cls(bg="#F0F7FA", accent="#2C7DA0", accent2="#F2C14E", accent3="#1D3557",
                   text="#1D3557", text2="#3D5A73", text3="#8AA5B8", line="#D6E4EC")

    @classmethod
    def business(cls):
        """通用汇报：中性灰蓝，绝不抢戏"""
        return cls(bg="#FAFAF9", accent="#F97316", accent2="#475569", accent3="#0F172A",
                   text="#0F172A", text2="#475569", text3="#94A3B8", line="#E7E5E4")


class Builder:
    """构建器。所有方法返回 slide，便于链式组织。"""

    def __init__(self, palette, title_font="微软雅黑", num_font="Arial Black",
                 note_font="Georgia", total_pages=None, footer_text=""):
        self.p = palette
        self.title_font = title_font
        self.num_font = num_font
        self.note_font = note_font
        self.total_pages = total_pages
        self.footer_text = footer_text

        self.prs = Presentation()
        self.prs.slide_width = px(CANVAS_W)
        self.prs.slide_height = px(CANVAS_H)
        self._blank = self.prs.slide_layouts[6]

    # ================= 底层图元 =================

    def _slide(self, bg=None):
        s = self.prs.slides.add_slide(self._blank)
        self.rect(s, 0, 0, CANVAS_W, CANVAS_H, fill=bg or self.p.bg)
        return s

    def rect(self, slide, x, y, w, h, fill=None, line=None, radius=None,
             line_w=1.0, opacity=None):
        """矩形。radius 给值则用圆角矩形（px）"""
        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
        shp = slide.shapes.add_shape(shape_type, px(x), px(y), px(w), px(h))
        if radius:
            # adjustment 值 = 圆角半径 / 短边
            adj = min(0.5, float(radius) / max(1.0, min(w, h)))
            shp.adjustments[0] = adj
        if fill:
            shp.fill.solid()
            shp.fill.fore_color.rgb = _rgb(fill)
        else:
            shp.fill.background()
        if line:
            shp.line.color.rgb = _rgb(line)
            shp.line.width = Pt(line_w)
        else:
            shp.line.fill.background()
        shp.shadow.inherit = False
        if opacity is not None:
            self._set_opacity(shp, opacity)
        return shp

    def _set_opacity(self, shape, opacity):
        """设置形状填充透明度（python-pptx 无原生 API，直接改 XML）"""
        solidFill = shape.fill._xPr.find(qn('a:solidFill'))
        if solidFill is None:
            return
        clr = solidFill.find(qn('a:srgbClr'))
        if clr is None:
            return
        alpha = parse_xml(
            f'<a:alpha xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            f'val="{int(opacity * 100000)}"/>')
        clr.append(alpha)

    def text(self, slide, x, y, w, h, content, size=19, bold=False,
             color=None, font=None, align="left", line_spacing=1.5,
             letter_spacing=None, anchor="top"):
        """文本框。content 可含 \n 分段。

        size 用 px —— 内部按 0.75 换算成 pt（96dpi → 72pt/inch）
        """
        tb = slide.shapes.add_textbox(px(x), px(y), px(w), px(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                              "bottom": MSO_ANCHOR.BOTTOM}[anchor]

        lines = str(content).split('\n')
        for i, ln in enumerate(lines):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                              "right": PP_ALIGN.RIGHT}[align]
            para.line_spacing = line_spacing
            run = para.add_run()
            run.text = ln
            run.font.size = Pt(size * 0.75)
            run.font.bold = bold
            run.font.color.rgb = _rgb(color or self.p.text)
            _set_font(run, font or self.title_font)
            if letter_spacing:
                run._r.get_or_add_rPr().set('spc', str(int(letter_spacing * 75)))
        return tb

    def pic(self, slide, path, x, y, w, h):
        """插图。按 cover 方式铺满给定区域（先等比放大再居中裁切）"""
        from PIL import Image
        iw, ih = Image.open(path).size
        target = w / h
        src = iw / ih
        pic = slide.shapes.add_picture(path, px(x), px(y), px(w), px(h))
        if src > target:      # 图更宽 → 裁左右
            frac = (1 - target / src) / 2
            pic.crop_left = frac
            pic.crop_right = frac
        elif src < target:    # 图更高 → 裁上下
            frac = (1 - src / target) / 2
            pic.crop_top = frac
            pic.crop_bottom = frac
        return pic

    def pic_placeholder(self, slide, x, y, w, h, hint="", spec=""):
        """配图占位框。素材未就绪时用，标清「该放什么图」。

        产出物在 PowerPoint 里一眼可见，方便交付后按清单补图；
        补图时把本方法调用换成 pic(slide, "assets/xxx.png", x, y, w, h) 即可。
        """
        self.rect(slide, x, y, w, h, fill=None, line=self.p.text3,
                  radius=10, line_w=1.5)
        # 内部浅色填充，暗示"这里待填充"
        self.rect(slide, x + 2, y + 2, w - 4, h - 4, fill=self.p.card,
                  radius=9, opacity=0.55)
        cy = y + h / 2 - 46
        self.text(slide, x + 24, cy, w - 48, 26, "［ 待补配图 ］",
                  size=16, bold=True, color=self.p.text3, align="center")
        if hint:
            self.text(slide, x + 24, cy + 32, w - 48, 30, hint,
                      size=13, color=self.p.text3, align="center", line_spacing=1.4)
        if spec:
            self.text(slide, x + 24, cy + 68, w - 48, 24, spec,
                      size=12, color=self.p.text3, align="center")

    def images_needed(self, items):
        """配图需求清单。items: [(文件名, 建议尺寸, prompt 要点, 所属页)]

        返回用法：把这行打印到控制台 / 写进 assets/IMAGES.md，
        交给外部生图能力逐张产出，再回填到 assets/。
        """
        lines = ["| 文件名 | 建议尺寸 | prompt 要点 | 所属页 |",
                 "| :-- | :-- | :-- | :-- |"]
        for name, size, prompt, page in items:
            lines.append(f"| `{name}` | {size} | {prompt} | {page} |")
        return "\n".join(lines)

    # ================= 母版骨架 =================

    def header(self, slide, title, chapter=None, page_no=None):
        """A 区：左侧粉色竖条 + 主标题 +（可选）右侧章节贴纸"""
        self.rect(slide, PAD_X, MARGIN_TOP + 12, 6, 34, fill=self.p.accent3, radius=3)
        self.text(slide, PAD_X + 18, MARGIN_TOP + 14, 800, 40, title,
                  size=34, bold=True, color=self.p.text, letter_spacing=1)
        if chapter:
            self.rect(slide, CANVAS_W - PAD_X - 108, MARGIN_TOP + 18, 108, 28,
                      fill=self.p.accent3, radius=6)
            # 贴纸文字色随底色亮度自适应——深色 accent3 上必须给白字，否则黑底黑字
            self.text(slide, CANVAS_W - PAD_X - 108, MARGIN_TOP + 24, 108, 20, chapter,
                      size=13, bold=True, color=on_color(self.p.accent3), align="center")
        if page_no is not None:
            self.footer(slide, page_no)

    def footer(self, slide, page_no):
        """C 区：左侧项目名（footer_text） + 右侧页码"""
        y = CANVAS_H - MARGIN_TOP - ZONE_C_H
        if self.footer_text:
            self.text(slide, PAD_X, y + 12, 460, 20, self.footer_text,
                      size=13, color=self.p.text3)
        total = self.total_pages or len(self.prs.slides._sldIdLst)
        self.text(slide, CANVAS_W - PAD_X - 120, y + 12, 120, 20,
                  f"{page_no:02d} / {total}", size=13, color=self.p.text3, align="right")

    # ================= 页面版式 =================

    def cover(self, main_title, sub_title, eyebrow=None, tag=None,
              stats=None, footnote=None, image=None):
        """封面：左侧文字栏纵向撑满 + 右侧全高主视觉"""
        s = self._slide()
        img_w = 780
        if image:
            self.pic(s, image, CANVAS_W - img_w, 0, img_w, CANVAS_H)
            # 左缘羽化：叠多层递减不透明度的色块，模拟渐变过渡
            steps = 12
            for i in range(steps):
                self.rect(s, CANVAS_W - img_w + i * 8, 0, 9, CANVAS_H,
                          fill=self.p.bg, opacity=round(1 - (i + 1) / steps, 3))

        if eyebrow:
            self.rect(s, PAD_X, 96, 268, 34, fill=self.p.accent3, radius=17)
            self.text(s, PAD_X + 16, 104, 240, 20, eyebrow,
                      size=14, bold=True, color=on_color(self.p.accent3), letter_spacing=1.6)

        self.text(s, PAD_X, 236, 620, 100, main_title,
                  size=82, bold=True, color=self.p.text, line_spacing=1.12, letter_spacing=3)
        self.rect(s, PAD_X + 4, 366, 68, 5, fill=self.p.accent2, radius=3)
        self.text(s, PAD_X, 396, 600, 80, sub_title,
                  size=19, color=self.p.text2, line_spacing=1.62)

        if stats:
            x = PAD_X
            for val, label in stats:
                self.text(s, x, 540, 200, 62, val, size=54, bold=True,
                          color=self.p.accent, font=self.num_font, line_spacing=0.95)
                self.text(s, x, 606, 200, 20, label, size=12.5, color=self.p.text3)
                x += 190

        if footnote:
            self.text(s, PAD_X, 664, 700, 20, footnote, size=12, color=self.p.text3)
        if tag:
            self.rect(s, CANVAS_W - PAD_X - 130, 96, 130, 32, fill=self.p.bg,
                      line=self.p.accent, radius=6)
            self.text(s, CANVAS_W - PAD_X - 130, 104, 130, 20, tag,
                      size=13, bold=True, color=self.p.accent, align="center")
        return s

    def section(self, number, title, desc, chapter=None):
        """章节扉页：巨型描边数字 + 右侧标题组（版式自定义，不含页脚）"""
        s = self._slide()
        self.rect(s, 118, 96, 300, 260, fill=self.p.accent3, radius=40, opacity=0.9)
        self.rect(s, 330, 268, 190, 190, fill=self.p.accent2, radius=95, opacity=0.35)

        tb = s.shapes.add_textbox(px(40), px(150), px(480), px(320))
        tf = tb.text_frame
        tf.word_wrap = False
        para = tf.paragraphs[0]
        run = para.add_run()
        run.text = number
        run.font.size = Pt(225)
        run.font.color.rgb = _rgb(self.p.bg)
        _set_font(run, self.num_font)
        # 描边字：无填充 + 描边
        rPr = run._r.get_or_add_rPr()
        rPr.set('ln', '50800')
        rPr.insert(0, parse_xml(
            '<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="25400">'
            f'<a:solidFill><a:srgbClr val="{self.p.accent.lstrip("#")}"/></a:solidFill></a:ln>'))

        x = 640
        self.rect(s, x, 232, 44, 3, fill=self.p.accent2, radius=2)
        if chapter:
            self.text(s, x + 56, 224, 400, 20, chapter, size=15, bold=True,
                      color=self.p.text2, letter_spacing=3.5)
        self.text(s, x, 258, 560, 90, title, size=68, bold=True,
                  color=self.p.text, line_spacing=1.12, letter_spacing=2)
        self.text(s, x, 372, 520, 100, desc, size=19,
                  color=self.p.text2, line_spacing=1.7)

        self.rect(s, CANVAS_W - 84, CANVAS_H - 56, 10, 10,
                  fill=self.p.accent3, radius=5)
        return s

    def big_number(self, slide, x, y, number, unit="", caption="",
                   num_size=96, unit_size=22, cap_size=13):
        """巨型锚点数字——每页至少一个，hero 页的核心视觉"""
        self.text(slide, x, y, 460, num_size * 1.0, number,
                  size=num_size, bold=True, color=self.p.accent,
                  font=self.num_font, line_spacing=0.9)
        if unit:
            self.text(slide, x + 320, y + num_size * 0.42, 140, unit_size * 1.6, unit,
                      size=unit_size, bold=True, color=self.p.text)
        if caption:
            self.text(slide, x + 320, y + num_size * 0.42 + unit_size * 1.3, 220,
                      cap_size * 2.4, caption, size=cap_size,
                      color=self.p.text3, font=self.note_font)

    def card(self, slide, x, y, w, h, title, body, tag=None,
             tag_color=None, head_color=None, radius=14,
             title_size=17, body_size=14):
        """内容卡片。head_color 为非白色块时，文字色自动适配底色亮度"""
        bg = head_color or self.p.card
        line = None if head_color else self.p.line
        # 深色卡片必须给浅色字，否则标题/正文与底色同色、等于隐形
        if head_color:
            t_color = b_color = on_color(head_color)
        else:
            t_color, b_color = self.p.text, self.p.text2
        self.rect(slide, x, y, w, h, fill=bg, line=line, radius=radius)
        cy = y + 14
        if tag:
            tw = max(56, len(tag) * 15 + 22)
            self.rect(slide, x + 15, cy, tw, 26,
                      fill=tag_color or self.p.accent2, radius=7)
            self.text(slide, x + 15, cy + 4, tw, 18, tag, size=13, bold=True,
                      color=on_color(tag_color or self.p.accent2), align="center")
            cy += 36
        self.text(slide, x + 15, cy, w - 30, 26, title,
                  size=title_size, bold=True, color=t_color)
        self.text(slide, x + 15, cy + title_size + 12, w - 30, h - (cy - y) - title_size - 24,
                  body, size=body_size, color=b_color, line_spacing=1.55)

    def bar_row(self, slide, x, y, w, label, value_text, ratio,
                color=None, growth=None, label_size=14):
        """数据条行：标签 + 数值 + 按比例填充的横条"""
        color = color or self.p.accent2
        self.text(slide, x, y, w * 0.55, 22, label, size=label_size, color=self.p.text)
        self.text(slide, x + w * 0.55, y, w * 0.3, 22, value_text,
                  size=label_size + 1, bold=True, color=self.p.text,
                  font=self.num_font, align="right")
        if growth:
            self.text(slide, x + w * 0.85, y, w * 0.15, 22, growth,
                      size=label_size - 1, bold=True, color=self.p.accent, align="right")
        self.rect(slide, x, y + 26, w, 5, fill=self.p.line, radius=3)
        self.rect(slide, x, y + 26, max(6, w * ratio), 5, fill=color, radius=3)

    def bullet_list(self, slide, x, y, w, items, gap=10,
                    num_color=None, title_size=16.5, body_size=13.5):
        """编号条目列表（替代「等宽卡片横排」的常见 AI 味写法）"""
        num_color = num_color or self.p.accent2
        cy = y
        for i, (t, d) in enumerate(items):
            self.rect(slide, x, cy, 36, 36, fill=num_color, radius=10)
            self.text(slide, x, cy + 8, 36, 20, f"{i+1:02d}", size=14, bold=True,
                      color=self.p.card, font=self.num_font, align="center")
            self.text(slide, x + 50, cy - 2, w - 50, 24, t,
                      size=title_size, bold=True, color=self.p.text)
            self.text(slide, x + 50, cy + title_size + 8, w - 50,
                      body_size * 3.2, d, size=body_size,
                      color=self.p.text2, line_spacing=1.5)
            cy += 36 + title_size + body_size * 2.6 + gap

    def timeline(self, slide, x, y, w, entries, label_w=92):
        """竖向时间轴：年份 + 轴点 + 文本卡"""
        n = len(entries)
        rowh = 116
        self.rect(slide, x + label_w + 7, y + 10, 2, rowh * n - 30, fill=self.p.line)
        for i, (year, title, desc, color) in enumerate(entries):
            cy = y + i * rowh
            self.text(slide, x, cy + 12, label_w, 32, year, size=26, bold=True,
                      color=color, font=self.num_font, align="right")
            self.rect(slide, x + label_w, cy + 20, 16, 16, fill=self.p.bg,
                      line=color, radius=8, line_w=3)
            self.rect(slide, x + label_w + 26, cy, w - label_w - 26, 100,
                      fill=self.p.card, line=self.p.line, radius=14)
            self.text(slide, x + label_w + 44, cy + 12, w - label_w - 70, 24, title,
                      size=20, bold=True, color=self.p.text)
            self.text(slide, x + label_w + 44, cy + 42, w - label_w - 70, 50, desc,
                      size=15, color=self.p.text2, line_spacing=1.6)

    def save(self, path):
        self.prs.save(path)
        return path


# ================= 自检 demo =================

if __name__ == "__main__":
    import sys, os

    out = sys.argv[1] if len(sys.argv) > 1 else "demo.pptx"
    b = Builder(Palette.cream_pink(), total_pages=3, footer_text="DEMO · 自检")

    b.cover("示例品牌", "这是封面副标题\n第二行说明文字",
            eyebrow="DEMO · 自检", tag="2025 年度",
            stats=[("630", "全球门店（家）"), ("43.8%", "海外营收占比")],
            footnote="本文件由 deck_builder.py 自动生成，用于验证渲染链路")

    s = b.section("01", "示例章节", "这是章节扉页的一句话说明。", chapter="CHAPTER 01")

    s = b._slide()
    b.header(s, "示例内容页", chapter="CHAPTER 01", page_no=3)
    b.big_number(s, PAD_X, 156, "141.6", "亿元", "单 IP 收入")
    b.bar_row(s, 640, 160, 560, "SKULLPANDA", "35.40 亿", 0.25, growth="+170.6%")
    b.bar_row(s, 640, 220, 560, "CRYBABY", "29.29 亿", 0.21, growth="+151.5%")
    b.bar_row(s, 640, 280, 560, "MOLLY", "28.97 亿", 0.20, growth="+38.4%")
    b.bullet_list(s, 640, 340, 560, [
        ("可复制的孵化能力", "不是单点抽中彩票，而是矩阵式产出。"),
        ("艺术家 IP 占九成", "自有产权，利润留在自己手里。"),
    ])

    b.save(out)
    print(f"OK -> {os.path.abspath(out)}")

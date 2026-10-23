# -*- coding: utf-8 -*-
"""Build a native 16:9 PPTX master (55 x 30.92 cm) - editable, not image slides."""
import os
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "report-master-16x9.pptx")

W, H = Cm(54.98), Cm(30.92)
INK = RGBColor(0x2F, 0x2F, 0x2F)
SUB = RGBColor(0x6F, 0x6F, 0x6F)
BLUE = RGBColor(0x1F, 0x5F, 0xBF)
HEAD_BG = RGBColor(0xEA, 0xF1, 0xFE)
ZEBRA = RGBColor(0xF7, 0xF9, 0xFC)
RULE = RGBColor(0xE5, 0xE7, 0xEB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COVER_TOP = RGBColor(0x8C, 0xC2, 0xF8)
COVER_BOT = RGBColor(0x5A, 0x9A, 0xE8)
FONT = "微软雅黑"

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

L = Cm(3.17)                # left margin
CW = W - Cm(6.34)           # content width
COLW = Cm(24.0)             # one column
GAP = Cm(0.64)


def tb(slide, l, t, w, h, text, size, color=INK, bold=False, align=PP_ALIGN.LEFT,
       spacing=1.0, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = spacing
    r = p.add_run()
    r.text = text
    f = r.font
    f.name = FONT
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    r._r.get_or_add_rPr().find(qn("a:latin")).set("typeface", FONT)
    ea = r._r.get_or_add_rPr().makeelement(qn("a:ea"), {})
    ea.set("typeface", FONT)
    r._r.get_or_add_rPr().append(ea)
    return box


def rule(slide, top, color=RULE, weight=Cm(0.05)):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, L, top, CW, weight)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    s.shadow.inherit = False
    return s


# ============ slide 2 first (content), then cover moved to front ============
# ---------- cover ----------
cover = prs.slides.add_slide(blank)
bg = cover.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
bg.line.fill.background()
bg.shadow.inherit = False
bg.fill.gradient()
bg.fill.gradient_angle = 90
stops = bg.fill.gradient_stops
stops[0].color.rgb = COVER_TOP
stops[0].position = 0.0
stops[1].color.rgb = COVER_BOT
stops[1].position = 1.0

# 白色圆角卡片承载 logo
card = cover.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(6.2), Cm(8.4), Cm(6.6), Cm(6.6))
card.fill.solid()
card.fill.fore_color.rgb = WHITE
card.line.fill.background()
card.shadow.inherit = False
card.adjustments[0] = 0.22

logo_path = os.path.join(HERE, "logo.jpg")
if os.path.exists(logo_path):
    cover.shapes.add_picture(logo_path, Cm(7.2), Cm(9.4), Cm(4.6), Cm(4.6))

tb(cover, Cm(15.6), Cm(8.9), Cm(33.0), Cm(3.0), "1037 拼拼", 80, WHITE, True)
tb(cover, Cm(15.8), Cm(13.4), Cm(33.0), Cm(1.6), "拼拼一下，就差你了！", 28, WHITE, False, PP_ALIGN.LEFT, 1.0)

# 分隔线（封面）
s = cover.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(15.8), Cm(16.4), Cm(12.0), Cm(0.06))
s.fill.solid()
s.fill.fore_color.rgb = WHITE
s.line.fill.background()
s.shadow.inherit = False

# 目录：两栏
toc = [
    ("一、项目概述", "二、项目背景与痛点", "三、用户研究与洞察", "四、产品定位与核心功能", "五、设计流程与方法"),
    ("六、关键界面设计说明", "七、商业模式", "八、项目成果与反思", "九、个人成长总结"),
]
x = Cm(15.8)
for col in toc:
    body = "\n".join(col)
    box = tb(cover, x, Cm(17.6), Cm(15.0), Cm(9.0), body, 16, WHITE, False, PP_ALIGN.LEFT, 1.9)
    x += Cm(16.5)

# 底部白条
band = cover.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, H - Cm(5.6), W, Cm(5.6))
band.fill.solid()
band.fill.fore_color.rgb = WHITE
band.line.fill.background()
band.shadow.inherit = False
tb(cover, L, H - Cm(4.9), Cm(30.0), Cm(1.0), "1037 拼拼 · 校园拼单平台 · 项目报告", 16, INK, True)
tb(cover, L, H - Cm(3.6), Cm(30.0), Cm(1.0), "Product + UI Design 全流程复盘", 12, SUB)
tb(cover, L, H - Cm(2.2), Cm(48.0), Cm(1.0),
   "Pivot Studio　｜　2023.03 – 2024.03　｜　产品设计 · UI 设计", 12, RGBColor(0xC8, 0x7A, 0x28), True)

# ---------- content page ----------
s2 = prs.slides.add_slide(blank)

tb(s2, L, Cm(2.2), Cm(30.0), Cm(1.6), "一、项目概述", 24, INK, True)
rule(s2, Cm(4.1))

tb(s2, L, Cm(4.9), CW, Cm(3.0),
   "「1037 拼拼」是由 Pivot Studio 团队孵化的校园平台类产品，面向华中科技大学学生的招募、"
   "拼团平台：可发布比赛队友招募、游戏拼团、剧本杀拼车等需求，双方通过私聊与个人主页"
   "快速了解，进行双向选择——满意就拼，效率较高。", 16, INK, False, PP_ALIGN.LEFT, 1.75)

# 左栏：表格
x1 = L
tb(s2, x1, Cm(8.4), COLW, Cm(1.1), "项目基本信息", 18, INK, True)

rows = [
    ("维度", "内容"),
    ("项目周期", "2023.03 – 2024.03（约 12 个月）"),
    ("担任角色", "Product + UI Design（主导）"),
    ("协作团队", "Pivot Studio 跨职能"),
    ("交付范围", "产品定义 · 用户研究 · 交互 · 高保真视觉"),
    ("项目转折", "App 完成后因运营不足转向 QQ 社群"),
]
tbl_w, tbl_h = COLW, Cm(0.95) * len(rows)
tbl_shape = s2.shapes.add_table(len(rows), 2, x1, Cm(9.6), tbl_w, tbl_h)
table = tbl_shape.table
table.columns[0].width = Cm(6.4)
table.columns[1].width = COLW - Cm(6.4)
table.rows[0].height = Cm(1.0)
for i in range(1, len(rows)):
    table.rows[i].height = Cm(0.95)

for i, (a, b) in enumerate(rows):
    for j, txt in enumerate((a, b)):
        cell = table.cell(i, j)
        cell.margin_left = Cm(0.25)
        cell.margin_right = Cm(0.2)
        cell.margin_top = Cm(0.1)
        cell.margin_bottom = Cm(0.1)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        if i == 0:
            cell.fill.fore_color.rgb = HEAD_BG
        elif i % 2 == 0:
            cell.fill.fore_color.rgb = ZEBRA
        else:
            cell.fill.fore_color.rgb = WHITE
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.line_spacing = 1.15
        r = p.add_run()
        r.text = txt
        r.font.name = FONT
        r.font.size = Pt(12)
        r.font.bold = (i == 0)
        r.font.color.rgb = BLUE if i == 0 else INK

tb(s2, x1, Cm(16.6), COLW, Cm(0.8), "表 1　项目基本信息", 10, SUB)

# 右栏：痛点
x2 = L + COLW + GAP
tb(s2, x2, Cm(8.4), COLW, Cm(1.1), "三大核心痛点", 18, INK, True)
bullets = [
    ("招募困难", "仅 26.5% 能在自己社交圈组好队伍，73.5% 需拓宽圈子找队友，64.1% 的需求长期未被满足。"),
    ("拼单分散", "拼车 / 拼团 / 剧本杀等需求散落在朋友圈、表白墙，信息不对等、筛选低效。"),
    ("社交焦虑", "39.3% 的学生渴望社交却缺途径或勇气；陌生拼单又存在安全与诈骗隐患。"),
]
box = s2.shapes.add_textbox(x2, Cm(9.9), COLW, Cm(9.0))
tf = box.text_frame
tf.word_wrap = True
tf.margin_left = 0
tf.margin_right = 0
tf.margin_top = 0
tf.margin_bottom = 0
for idx, (head, body) in enumerate(bullets):
    p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
    p.line_spacing = 1.6
    p.space_after = Pt(18)
    r1 = p.add_run()
    r1.text = head + "："
    r1.font.name = FONT
    r1.font.size = Pt(16)
    r1.font.bold = True
    r1.font.color.rgb = INK
    r2 = p.add_run()
    r2.text = body
    r2.font.name = FONT
    r2.font.size = Pt(16)
    r2.font.color.rgb = INK

# 底部强调条
acc = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, x2, Cm(19.4), COLW, Cm(1.7))
acc.fill.solid()
acc.fill.fore_color.rgb = HEAD_BG
acc.line.fill.background()
acc.shadow.inherit = False
tb(s2, x2 + Cm(0.4), Cm(19.75), COLW - Cm(0.8), Cm(1.2),
   "机会点：以「拼单事件」为载体，做安全、高效、信息对等的双向选择平台。", 12, BLUE, True)

# 页脚
rule(s2, H - Cm(3.4))
tb(s2, L, H - Cm(2.9), Cm(30.0), Cm(0.9), "© Pivot Studio　｜　林丽丽 Lili Lin　｜　2023.03 – 2024.03", 10, SUB)
tb(s2, W - L - Cm(3.0), H - Cm(2.9), Cm(3.0), Cm(0.9), "02", 10, SUB, False, PP_ALIGN.RIGHT)

prs.save(OUT)
print("saved", OUT, os.path.getsize(OUT) // 1024, "KB")

# -*- coding: utf-8 -*-
"""Build the 9:16 vertical master (1080x1920) - social/short-video pages.

产物：
  vertical-9x16.pptx      2 个版式母版：L1 竖封面 + L2 竖内容页
  vertical-9x16-16p.pptx  16 页竖版正式稿（v2：按第 9 节 + 设计顾问四页改法重排）
  vertical-9x16-13p.pptx  旧 13 页（保留，不覆盖）

v2 相对 v1 的改动（全部有实测依据，写在 _docs/master-layout-spec.md 第 9 节）：
  1. P03 2×2 四小图 → 拆两页，每页 2 张并排（间距 56、居中、等比不拉伸）
  2. P05 三张并排（409/270/225）→ 拆三页，最大那张单独一页（media-full 936 宽）
  3. P09 4 联 + 11 连缩略图 → 切出单机，2 联并排；缩略图条删除
  4. P10 同 P09 处理（切出最大单图，缩略图条删除）
  5. 媒体页新增「高媒体带」600→1448（848px 高），下方要点 3 条 pitch 110
  6. 修 v1 的拉伸 bug：mode="width" 原来把任意源图强制成 936×560（等比失真最大 23%）
"""
import io
import os
import numpy as np
import pymupdf
from PIL import Image
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn, nsdecls
from pptx.oxml import parse_xml

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.abspath(os.path.join(HERE, "..", "..", "..", "_out", "master"))
MASTER = os.path.join(OUT_DIR, "vertical-9x16.pptx")
OUT16 = os.path.join(OUT_DIR, "vertical-9x16-18p.pptx")
SRC_PDF = os.path.abspath(os.path.join(HERE, "..", "..", "..", "_out", "pinpin", "deck", "pdf",
                                       "report_vector.pdf"))
FIGDIR = os.path.join(OUT_DIR, "vertical-figs")
HIFIGDIR = os.path.join(OUT_DIR, "vertical-figs-hi")
CROPDIR = os.path.join(OUT_DIR, "vertical-figs-crop")

PX = 9525  # EMU per px at 96dpi
W, H = 1080 * PX, 1920 * PX

INK = RGBColor(0x2F, 0x2F, 0x2F)
SUB = RGBColor(0x6F, 0x6F, 0x6F)
BLUE = RGBColor(0x1F, 0x5F, 0xBF)
DEEP = RGBColor(0x1F, 0x5F, 0xBF)
HEAD_BG = RGBColor(0xEA, 0xF1, 0xFE)
ZEBRA = RGBColor(0xF7, 0xF9, 0xFC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ORANGE = RGBColor(0xF0, 0x8A, 0x24)
COVER_TOP = RGBColor(0x8C, 0xC2, 0xF8)
COVER_BOT = RGBColor(0x5A, 0x9A, 0xE8)
FONT = "微软雅黑"

MARGIN = 72 * PX          # 左右 72px
TOP_H = 560 * PX          # 顶部品牌块
BODY_H = 608 * PX         # 常规媒体带高（v1 布局）
BOTTOM_Y = TOP_H + BODY_H  # 1168

BX, BY = MARGIN, TOP_H + 24 * PX          # 常规媒体框 左上 72,584
BW, BH = W - 2 * MARGIN, BODY_H - 48 * PX  # 936 × 560

# v2：高媒体带（用于 4 联切图 / media-full 大图）
# 844 而不是 848/840：手机竖截判据是「高 ≥840」，844 留 4px 余量；
# 1440 起点 + pitch 112 → 第 3 条要点文字底 1810 < 页脚 1830，不压页脚。
TALL_Y = 584 * PX
TALL_H = 844 * PX
POINTS_Y_TALL = 1440 * PX
POINTS_PITCH_TALL = 112 * PX
POINTS_PITCH = 168 * PX
POINTS_Y = BOTTOM_Y + 40 * PX
PAIR_GAP = 32 * PX        # 并排间距：设计顾问定档 32（56 会让窄的那张掉到 419px＜420 判据）


def blank_deck():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    return prs


def text(slide, x, y, w, h, s, px_size, color=INK, bold=False,
         align=PP_ALIGN.LEFT, spacing=1.0, anchor=MSO_ANCHOR.TOP, runs=None):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    if runs is not None:                       # 单段多 run（粗体小标题 + 正文同一行）
        p = tf.paragraphs[0]
        p.alignment = align
        p.line_spacing = spacing
        for seg, b in runs:
            r = p.add_run()
            r.text = seg
            r.font.name = FONT
            r.font.size = Pt(px_size * 0.75)
            r.font.bold = b
            r.font.color.rgb = color if b else INK
            r._r.get_or_add_rPr().append(_ea())
        return box
    for i, line in enumerate(str(s).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = line
        r.font.name = FONT
        r.font.size = Pt(px_size * 0.75)
        r.font.bold = bold
        r.font.color.rgb = color
        r._r.get_or_add_rPr().append(_ea())
    return box


def _ea():
    return parse_xml('<a:ea %s typeface="%s"/>' % (nsdecls("a"), FONT))


def rect(slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE, radius=None):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    s.shadow.inherit = False
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    return s


def fit(rw, rh, max_w, max_h):
    """等比缩放到 max_w × max_h 之内（只缩不放，绝不拉伸）。"""
    k = min(max_w / rw, max_h / rh)
    return int(round(rw * k)), int(round(rh * k))


def nat(p):
    with Image.open(p) as im:
        return im.size


# ================= L1 竖封面 =================
def cover_slide(prs):
    c = prs.slides.add_slide(prs.slide_layouts[6])
    bg = rect(c, 0, 0, W, H, COVER_TOP)
    bg.fill.gradient()
    bg.fill.gradient_angle = 90
    gs = bg.fill.gradient_stops
    gs[0].color.rgb = COVER_TOP
    gs[0].position = 0.0
    gs[1].color.rgb = COVER_BOT
    gs[1].position = 1.0

    rect(c, (W - 320 * PX) // 2, 300 * PX, 320 * PX, 320 * PX, WHITE,
         MSO_SHAPE.ROUNDED_RECTANGLE, 0.22)
    logo = next((p for p in (os.path.join(OUT_DIR, "logo.jpg"), os.path.join(HERE, "logo.jpg"))
                 if os.path.exists(p)), None)
    if logo:
        add_pic(c, logo, (W - 220 * PX) // 2, 350 * PX, 220 * PX, 220 * PX,
                alt="1037 拼拼 App 图标")
    else:
        print("[vertical] WARN: 未找到 logo.jpg，封面 logo 卡为空")

    text(c, 0, 720 * PX, W, 140 * PX, "1037 拼拼", 96, WHITE, True, PP_ALIGN.CENTER)
    text(c, 0, 880 * PX, W, 70 * PX, "拼拼一下，就差你了！", 44, WHITE, False, PP_ALIGN.CENTER)
    rect(c, (W - 180 * PX) // 2, 1010 * PX, 180 * PX, 6 * PX, WHITE)
    text(c, MARGIN, 1058 * PX, W - 2 * MARGIN, 190 * PX,
         "① 项目概述 · 背景与痛点 · 用户研究\n"
         "② 产品定位 · 核心功能 · 设计流程\n"
         "③ 界面设计 · 商业模式 · 成果反思",
         30, RGBColor(0xE4, 0xEF, 0xFC), False, PP_ALIGN.CENTER, 1.55)

    rect(c, 0, H - 420 * PX, W, 420 * PX, WHITE)
    text(c, MARGIN, H - 340 * PX, W - 2 * MARGIN, 60 * PX,
         "校园拼单平台 · 项目报告", 44, INK, True)
    text(c, MARGIN, H - 260 * PX, W - 2 * MARGIN, 50 * PX,
         "Product + UI Design 全流程复盘", 30, SUB)
    text(c, MARGIN, H - 180 * PX, W - 2 * MARGIN, 50 * PX,
         "Pivot Studio　｜　2023.03 – 2024.03", 30, ORANGE, True)


def add_pic(container, path, x, y, w, h, alt="", cols=1):
    """插图统一入口，按 10-19 三方约定落元数据（各管一件事，不混用）：

    - title  = image_cols=N  → 校验器读（verify_vertical.py 判 936/N ≥ 420）
    - descr  = 人类描述       → 屏幕阅读器读，不许放机器元数据
    - name   = python-pptx 默认按文件名给，供人和脚本查

    N = 这张图「内部」并排的联数；本脚本在抽图阶段就把 4 联切成独立图片，
    所以放上去的每一张都是 1 联。若哪天直接放未切的联图条，这里必须写真联数。
    """
    pic = container.shapes.add_picture(str(path), x, y, w, h)
    cNvPr = pic._element.nvPicPr.cNvPr
    cNvPr.set("title", f"image_cols={cols}")
    if alt:
        cNvPr.set("descr", alt)
    return pic


# ================= L2 竖内容页 =================
def render_body(s, body, by, bh, alt=""):
    """把一页的「主体」画进 936 × bh 的框里（by 为框顶 y）；alt 为该页标题，做人类描述前缀。"""
    kind = body[0]
    if kind == "figs":
        figs, mode = body[1], body[2]
        rs = [nat(p) for p in figs]
        ratios = [a / b for a, b in rs]
        if mode == "width":                       # 单图：等比 fit 进 936×bh，居中（v1 此处会拉伸）
            w, h = fit(rs[0][0], rs[0][1], BW, bh)
            add_pic(s, figs[0], BX + (BW - w) // 2, by + (bh - h) // 2, w, h, alt=alt)
        elif mode == "pair":                      # v2：2 张并排、等高、居中、等比不拉伸
            gap = PAIR_GAP
            h = min(bh, int((BW - gap) / sum(ratios)))
            ws = [int(round(h * r)) for r in ratios]
            tot = sum(ws) + gap
            x = BX + (BW - tot) // 2
            for i, p in enumerate(figs[:2]):
                add_pic(s, p, x, by + (bh - h) // 2, ws[i], h,
                        alt=f"{alt}（并排 {i + 1}/{len(figs[:2])}）")
                x += ws[i] + gap
        elif mode == "row":                       # 并排等高，整排按宽铺满（高度不得超过 bh）
            gap = 16 * PX
            h = min(bh, int((BW - gap * (len(figs) - 1)) / sum(ratios)))
            ws = [int(round(h * r)) for r in ratios]
            tot = sum(ws) + gap * (len(figs) - 1)
            x = BX + (BW - tot) // 2
            for i, p in enumerate(figs):
                add_pic(s, p, x, by + (bh - h) // 2, ws[i], h,
                        alt=f"{alt}（并排 {i + 1}/{len(figs)}）")
                x += ws[i] + gap
        elif mode == "stack":                     # 竖排，整组按宽或按高铺满
            gap = 16 * PX
            w = BW
            hs = [int(round(w / r)) for r in ratios]
            if sum(hs) + gap * (len(figs) - 1) > bh:
                w = int((bh - gap * (len(figs) - 1)) / sum(1 / r for r in ratios))
                hs = [int(round(w / r)) for r in ratios]
            y = by + (bh - (sum(hs) + gap * (len(figs) - 1))) // 2
            for i, p in enumerate(figs):
                add_pic(s, p, BX + (BW - w) // 2, y, w, hs[i],
                        alt=f"{alt}（竖排 {i + 1}/{len(figs)}）")
                y += hs[i] + gap
    elif kind == "card":                          # 纯文字页：淡蓝信息卡，保持可编辑
        rect(s, BX, by, BW, bh, ZEBRA)
        rows, y = body[1], by + 56 * PX
        for label, value in rows:
            rect(s, BX + 40 * PX, y + 12 * PX, 8 * PX, 44 * PX, BLUE)
            text(s, BX + 72 * PX, y, BW - 112 * PX, 50 * PX, label, 30, SUB, True)
            text(s, BX + 72 * PX, y + 48 * PX, BW - 112 * PX, 60 * PX, value, 34, INK)
            y += 112 * PX
    elif kind == "tree":
        # 原生三级清单：重绘功能模块图（源图 1564×811，缩到 936 宽后模块名只剩 9px、
        # 子项 10-14px 不可读；2 拆即使占满整页也只有 17.2px＜18px 判据，故按规格 9.7 重绘）
        root, modules = body[1], body[2]
        rw, rh = 380 * PX, 58 * PX
        n = len(modules)
        gap = 10 * PX
        row_h = min(96 * PX, max(60 * PX, (bh - rh - 24 * PX - gap * (n - 1)) // n))
        rect(s, BX + (BW - rw) // 2, by, rw, rh, DEEP, MSO_SHAPE.ROUNDED_RECTANGLE, 0.18)
        text(s, BX + (BW - rw) // 2, by + 13 * PX, rw, 38 * PX, root, 26, WHITE, True,
             PP_ALIGN.CENTER)
        y = by + rh + 24 * PX
        for name, subs in modules:
            rect(s, BX, y, BW, row_h, ZEBRA, MSO_SHAPE.ROUNDED_RECTANGLE, 0.08)
            rect(s, BX, y + row_h // 2 - 22 * PX, 8 * PX, 44 * PX, BLUE)
            text(s, BX + 28 * PX, y + (row_h - 34 * PX) // 2, 250 * PX, 40 * PX, name, 24, INK, True)
            text(s, BX + 292 * PX, y + (row_h - 34 * PX) // 2, BW - 330 * PX, row_h - 16 * PX,
                 "　·　".join(subs) if subs else "—", 18, SUB, spacing=1.3)
            y += row_h + gap
    elif kind == "steps":
        # 原生五阶段 + 色彩规范：重绘设计规范整图（源图 2476×1725，色值小字 8-9px，
        # 整块裁切后 k=0.489 → 视觉 3px，按规格 9.7 只能重绘）
        steps, colors = body[1], body[2]
        y = by
        for i, st in enumerate(steps):
            rect(s, BX, y, 44 * PX, 44 * PX, HEAD_BG, MSO_SHAPE.ROUNDED_RECTANGLE, 0.2)
            text(s, BX, y + 7 * PX, 44 * PX, 34 * PX, f"{i + 1:02d}", 22, BLUE, True, PP_ALIGN.CENTER)
            text(s, BX + 68 * PX, y + 3 * PX, BW - 68 * PX, 44 * PX, st, 26, INK)
            y += 62 * PX
        y = by + 5 * 62 * PX + 24 * PX
        rect(s, BX, y, BW, 2 * PX, RGBColor(0xE5, 0xE7, 0xEB))
        y += 26 * PX
        sw, sh, gap = 104 * PX, 104 * PX, 20 * PX
        x = BX
        for label, hexs, col in colors:
            rect(s, x, y, sw, sh, col, MSO_SHAPE.ROUNDED_RECTANGLE, 0.12)
            text(s, x, y + sh + 10 * PX, sw + 56 * PX, 34 * PX, label, 22, INK, True)
            text(s, x, y + sh + 44 * PX, sw + 56 * PX, 30 * PX, hexs, 20, SUB)
            x += sw + gap


def content_slide(prs, no, title, lead, body, points, page_label, tall=False):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, TOP_H, DEEP)
    text(s, MARGIN, 90 * PX, W - 2 * MARGIN, 120 * PX, no, 96, RGBColor(0x5A, 0x9A, 0xE8), True)
    text(s, MARGIN, 250 * PX, W - 2 * MARGIN, 80 * PX, title, 52, WHITE, True)
    text(s, MARGIN, 350 * PX, W - 2 * MARGIN, 176 * PX, lead, 30, RGBColor(0xDC, 0xE8, 0xFA),
         spacing=1.3)

    by, bh = (TALL_Y, TALL_H) if tall else (BY, BH)
    if body:
        render_body(s, body, by, bh, alt=title)

    y = POINTS_Y_TALL if tall else POINTS_Y
    pitch = POINTS_PITCH_TALL if tall else POINTS_PITCH
    for head, txt in points[:3]:
        rect(s, MARGIN, y + 14 * PX, 8 * PX, 44 * PX, BLUE)
        text(s, MARGIN + 24 * PX, y, W - 2 * MARGIN - 24 * PX, 40 * PX, head, 30, SUB, True)
        text(s, MARGIN + 24 * PX, y + 44 * PX, W - 2 * MARGIN - 24 * PX, 100 * PX, txt, 34, INK,
             spacing=1.25)
        y += pitch

    text(s, MARGIN, H - 90 * PX, W - 2 * MARGIN, 40 * PX,
         "© Pivot Studio　｜　林丽丽 Lili Lin", 22, SUB)
    text(s, W - MARGIN - 160 * PX, H - 90 * PX, 160 * PX, 40 * PX, page_label, 22, SUB,
         False, PP_ALIGN.RIGHT)


# ================= 插图：从源报告 PDF 抽出并按需降采样 =================
def extract_figs(max_px=1200, quality=92, subdir="vertical-figs"):
    outdir = os.path.join(OUT_DIR, subdir)
    os.makedirs(outdir, exist_ok=True)
    if not os.path.exists(SRC_PDF):
        raise SystemExit(f"[vertical] 源报告 PDF 不存在: {SRC_PDF}")
    doc = pymupdf.open(SRC_PDF)
    out = {}
    for i, page in enumerate(doc, 1):
        seen, k = set(), 0
        for im in page.get_images(full=True):
            if im[0] in seen:
                continue
            seen.add(im[0])
            d = doc.extract_image(im[0])
            k += 1
            path = os.path.join(outdir, f"p{i:02d}_{k}.jpg")
            src = Image.open(io.BytesIO(d["image"])).convert("RGB")
            if max(src.size) > max_px:
                r = max_px / max(src.size)
                src = src.resize((max(1, int(src.width * r)), max(1, int(src.height * r))),
                                 Image.LANCZOS)
            src.save(path, "JPEG", quality=quality, optimize=True)
            out.setdefault(i, []).append(path)
    return out


def split_panels(path, min_gap_px=6, bg_tol=18, min_panel_px=40):
    """按列背景间隙切出并排面板（像素证据，不靠肉眼）。返回 [(x0, x1, w)]。"""
    a = np.asarray(Image.open(path).convert("L"), dtype=np.float32)
    h, w = a.shape
    hist = np.bincount((a.astype(np.uint8).ravel() // 4), minlength=64)
    bg = int(np.argmax(hist)) * 4
    is_bg = (a.std(axis=0) < bg_tol) & (np.abs(a.mean(axis=0) - bg) < bg_tol * 2)
    runs, start = [], None
    for x, b in enumerate(is_bg):
        if b and start is None:
            start = x
        elif not b and start is not None:
            runs.append((start, x - 1))
            start = None
    if start is not None:
        runs.append((start, w - 1))
    cuts = [(s, e) for s, e in runs if e - s + 1 >= min_gap_px and s > 0 and e < w - 1]
    bounds, prev = [], 0
    for s, e in cuts:
        bounds.append((prev, s - 1))
        prev = e + 1
    bounds.append((prev, w - 1))
    return [(s, e, e - s + 1) for s, e in bounds if e - s + 1 >= min_panel_px]


def crop_panel(path, x0, x1, out_path):
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    im = Image.open(path).convert("RGB")
    im.crop((x0, 0, x1 + 1, im.height)).save(out_path, "JPEG", quality=92, optimize=True)
    return out_path


# ================= 16 页竖版内容表（v2） =================
def build_16(figs, hi, crops):
    F = lambda n, k: figs[n][k - 1]
    pages = [
        dict(no="01", title="项目概述",
             lead="从 0 到 1 的校园拼单产品，12 个月独立完成产品定义与高保真交付",
             body=("card", [("项目周期", "2023.03 – 2024.03，约 12 个月"),
                            ("担任角色", "Product + UI Design，主导全流程"),
                            ("交付范围", "产品定义 · 用户研究 · 交互流程 · 高保真视觉")]),
             points=[("项目起点", "本科进入 Pivot Studio 后负责的主要项目，也是做 UI Design 的起点"),
                     ("独立完成", "从低保真原型到高保真设计稿，全程 0→1 独立完成"),
                     ("项目转折", "App 搭建完成后因运营人手不足，转向更易维护的 QQ 社群")]),
        dict(no="03", title="用户研究与洞察（一）",
             lead="面向华科在校生随机发放 120 份问卷，拼单需求满足度评分分布",
             body=("figs", [F(3, 1)], "width"),
             points=[("需求满足度", "64.1% 的组队需求无法被满足，缺靠谱、可追溯的渠道"),
                     ("安全顾虑", "62% 担心陌生拼单的诈骗与纠纷，需要身份可溯源")], tall=True),
        dict(no="03", title="用户研究与洞察（二）",
             lead="找队友经历分布：有过成功组队的只占少数，多数停留在「有想法没行动」",
             body=("figs", [F(3, 2)], "width"),
             points=[("找队友情况", "仅 26.5% 能在圈子内组好队伍，73.5% 有强烈找队友需求"),
                     ("用户画像", "竞赛驱动型 / 娱乐社交型 / 效率拒绝型三类核心人群")], tall=True),
        dict(no="03", title="用户研究与洞察（三）",
             lead="拼单场景偏好：剧本杀、出去干饭、约拍拍照等高频社交活动最集中",
             body=("figs", [F(3, 3)], "width"),
             points=[("场景集中", "拼车、拼团、剧本杀三场景占 71%，适合场景化模板")], tall=True),
        dict(no="03", title="用户研究与洞察（四）",
             lead="社交偏好分布：与好朋友同行、认识新朋友是两类正向体验",
             body=("figs", [F(3, 4)], "width"),
             points=[("独处偏好", "享受独自一人 / 想和别人一起但只能独自一人 / 和好朋友一起 / 几乎每次都能认识新朋友"),
                     ("结论", "需求高频、场景集中，适合做垂直校园拼单平台")], tall=True),
        dict(no="04", title="产品定位与核心功能",
             lead="四种拼单情景、六大功能模块，搭建高效沟通、双向选择、信息对等的平台",
             body=("tree", "拼拼功能模块图",
                   [("登录／注册模块", ["账号 + 密码随行登录／注册", "找回密码"]),
                    ("搜索模块", []),
                    ("浏览模块", ["浏览广场页面", "浏览我关注的标签相关", "选择添加／删除关注标签"]),
                    ("我的模块", ["我参与过的／我发布过的", "建议与反馈"]),
                    ("消息模块", ["接收用户评论、点赞、官方通知"])]),
             points=[("安全可靠", "校园邮箱 + 学号认证，拼单安全可靠、可溯源"),
                     ("真需求", "到截止日期拼拼贴自动删除，保证首页均为真需求"),
                     ("高效匹配", "标签筛选精准匹配需求，节省时间、提高成功率")]),
        dict(no="04", title="核心流程（一）",
             lead="关键路径：登录 / 注册 → 发布、编辑拼拼贴 → 浏览拼拼贴并拼单",
             body=("figs", [F(5, 1)], "width"),
             points=[("登录 / 注册", "校园邮箱验证 + 学号认证，建立可信身份"),
                     ("发布拼拼贴", "选类型、填需求、设截止时间，三步发布"),
                     ("浏览并拼单", "标签筛选后私聊确认，双向选择后拼单")], tall=True),
        dict(no="04", title="核心流程（二）",
             lead="发布路径：选类型、填需求、设截止时间，三步完成",
             body=("figs", [F(5, 2)], "width"),
             points=[("发布入口", "底部 Tab 居中主按钮，一步进入发布"),
                     ("表单降门槛", "默认值与模板减少重复输入"),
                     ("截止时间", "到期自动下架，保证信息新鲜")], tall=True),
        dict(no="04", title="核心流程（三）",
             lead="浏览与拼单：标签筛选后私聊确认，双向选择后拼单",
             body=("figs", [F(5, 3)], "width"),
             points=[("标签筛选", "按类型与时间筛选，快速定位真需求"),
                     ("详情确认", "发起人信息、标签、截止时间一屏可见"),
                     ("私聊闭环", "加入后自动进入私聊，交易路径最短")], tall=True),
        dict(no="05", title="设计流程与方法",
             lead="Slogan：拼拼一下，就差你了！从想法到像素，全程独立设计，逐层收敛",
             body=("steps",
                   ["用户研究", "信息架构", "低保真", "中保真", "高保真"],
                   [("品牌蓝主色", "#2E7CF6", RGBColor(0x2E, 0x7C, 0xF6)),
                    ("强调橙", "#FF8A3D", RGBColor(0xFF, 0x8A, 0x3D)),
                    ("次级灰", "#7E7E7E", RGBColor(0x7E, 0x7E, 0x7E)),
                    ("浅底灰", "#F3F3F3", RGBColor(0xF3, 0xF3, 0xF3)),
                    ("纯白", "#FFFFFF", RGBColor(0xFF, 0xFF, 0xFF))]),
             points=[("五阶段收敛", "用研 → 信息架构 → 低保真 → 中保真 → 高保真"),
                     ("设计系统", "品牌蓝 #2E7CF6 主色、橙 #FF8A3D 强调色"),
                     ("组件沉淀", "拼拼贴卡片、标签、主按钮、私聊气泡、底部 Tab")]),
        dict(no="05", title="设计系统与组件",
             lead="三级字阶 + 核心组件库，保证视觉一致性与可交付性",
             body=("figs", [F(7, 1)], "width"),
             points=[("字号三级", "标题 / 正文 / 辅助三级字阶，全站统一"),
                     ("色彩规范", "品牌蓝主色 + 橙色强调 + 深蓝与浅灰中性色"),
                     ("可交付", "组件化输出，设计稿可直接交给开发落地")]),
        dict(no="06", title="关键界面设计说明",
             lead="高保真界面围绕「首页—发布—我的—个人主页」主线展开",
             body=("figs", [F(8, 1), F(8, 2)], "row"),
             points=[("首页", "信息流呈现附近 / 热门拼拼贴，标签筛选与实时参与人数强化社区感"),
                     ("发布", "发布流程低门槛，支持选择拼单类型与截止时间"),
                     ("个人主页 / 私聊", "标签与动态帮助建立信任，私聊承接交易闭环")]),
        dict(no="06", title="首页（高保真）",
             lead="信息流 + 标签筛选 + 实时参与人数，强化社区感",
             body=("figs", [crops["p09_ph1"], crops["p09_ph2"]], "pair"),
             points=[("信息流", "附近 / 热门双入口，拼拼贴卡片一屏浏览"),
                     ("社区感", "实时参与人数与截止倒计时强化紧迫感"),
                     ("入口", "底部 Tab 四宫格，发布按钮居中突出")], tall=True),
        dict(no="06", title="发布与加入（高保真）",
             lead="发布低门槛，加入一步到位，闭环路径最短",
             body=("figs", [crops["p10_main"]], "width"),
             points=[("发布", "选类型、填标题与描述、设截止时间，三步完成"),
                     ("加入", "详情页一键加入，自动进入私聊建立连接"),
                     ("降门槛", "默认值与模板减少重复输入，降低发布意愿门槛")], tall=True),
        dict(no="06", title="加入页（高保真）",
             lead="详情页承接决策与加入动作，信息一屏可见",
             body=("figs", [F(11, 1)], "width"),
             points=[("详情页", "发起人信息、标签、截止时间一屏可见"),
                     ("一键加入", "加入后自动进入私聊，交易闭环最短"),
                     ("安全感", "校园邮箱 + 学号认证标识全程可见，抑制诈骗")]),
        dict(no="06", title="我的页（高保真）",
             lead="个人主页展示标签与动态，帮助快速建立信任、找到同好",
             body=("figs", [F(12, 1)], "width"),
             points=[("个人主页", "标签 + 动态 + 历史拼单，一眼看清是否同好"),
                     ("信任感", "校园认证标识与信用记录可视化呈现"),
                     ("管理", "我发布的 / 我加入的，状态与倒计时清晰")]),
        dict(no="07", title="商业模式与项目反思",
             lead="四条变现路径兼顾短期增值与长期私域沉淀，0→1 完整复盘",
             body=("card", [("付费增值", "提升个人主页标签上限；帖子付费置顶 / 加推荐值"),
                            ("IP 周边", "拼拼主题贴纸、书签、钥匙扣等低成本周边"),
                            ("应用内广告", "广告横栏与商家合作接单，按曝光收费"),
                            ("私域矩阵", "公众号 / QQ 空间号沉淀私域，商单盈利")]),
             points=[("项目成果", "0→1 独立完成产品定义、交互流程与高保真视觉设计"),
                     ("最大遗憾", "运营闭环未跑通，架构调整后转向 QQ 社群"),
                     ("经验", "早期应前置验证运营与增长路径，而非只聚焦设计侧")]),
    ]

    prs = blank_deck()
    cover_slide(prs)
    for i, p in enumerate(pages, 2):
        content_slide(prs, p["no"], p["title"], p["lead"], p["body"], p["points"],
                      f"{i:02d} / {len(pages) + 1}", tall=p.get("tall", False))
    prs.save(OUT16)
    return prs, OUT16


# ================= 主流程 =================
figs = extract_figs(1200, 92, "vertical-figs")
hi = extract_figs(3000, 95, "vertical-figs-hi")
print(f"[vertical] 抽出插图 {sum(len(v) for v in figs.values())} 张（1200px）"
      f" + {sum(len(v) for v in hi.values())} 张（3000px 高清，用于切图）")

# --- P09：4 联 → 切出单机，取前两台；11 连缩略图条删除 ---
p9 = hi[9][0]
pan = split_panels(p9)
print(f"[vertical] P09 源图 {nat(p9)} 检出 {len(pan)} 联: "
      f"{[(s, e, w) for s, e, w in pan]}")
if len(pan) != 4:
    print("[vertical] WARN: P09 未检出 4 联，切图按检出结果处理")
ph = [crop_panel(p9, s, e, os.path.join(CROPDIR, f"p09_phone{i+1}.jpg"))
      for i, (s, e, w) in enumerate(pan[:2])]
print(f"[vertical] P09 切机: " + "  ".join(f"{os.path.basename(p)} {nat(p)}" for p in ph))

# --- P10：只有 1 张真图 + 缩略图条 → 取最大面板 ---
p10 = hi[10][0]
pan10 = split_panels(p10)
main = max(pan10, key=lambda t: t[2]) if pan10 else (0, nat(p10)[0] - 1, nat(p10)[0])
print(f"[vertical] P10 源图 {nat(p10)} 检出 {len(pan10)} 块，最大块 {main}")
p10_main = crop_panel(p10, main[0], main[1], os.path.join(CROPDIR, "p10_main.jpg"))
print(f"[vertical] P10 取图: p10_main.jpg {nat(p10_main)}（p10_2 缩略图条删除）")

crops = {"p09_ph1": ph[0], "p09_ph2": ph[1], "p10_main": p10_main}

# 2 个版式母版（L1 竖封面 + L2 竖内容页）
m = blank_deck()
cover_slide(m)
content_slide(m, "02", "项目基本信息", "12 个月，从产品定义做到高保真交付",
              ("card", [("项目周期", "2023.03 – 2024.03，约 12 个月"),
                        ("担任角色", "Product + UI Design，主导全流程")]),
              [("项目周期", "2023.03 – 2024.03，约 12 个月"),
               ("担任角色", "Product + UI Design，主导全流程"),
               ("协作方式", "Pivot Studio 产品 / 前端 / 移动端跨职能")],
              "02 / 16")
m.save(MASTER)
print(f"[vertical] saved {MASTER}  {os.path.getsize(MASTER)//1024} KB  {len(m.slides)} layouts")

prs16, out16 = build_16(figs, hi, crops)
print(f"[vertical] saved {out16}  {os.path.getsize(out16)/1048576:.2f} MB  {len(prs16.slides)} pages")

# --- 每页实测：图片尺寸 / 是否越界 / 是否拉伸 ---
print("\n[vertical] 每页图位实测（PX=px@96dpi）：")
bad = 0
for i, s in enumerate(prs16.slides, 1):
    pics = [sh for sh in s.shapes if sh.shape_type == 13]
    if not pics:
        continue
    line = []
    for sh in pics:
        iw, ih = sh.image.size
        w, h = sh.width // PX, sh.height // PX
        ar_src, ar_put = iw / ih, w / h
        stretch = abs(ar_src - ar_put) / ar_src
        flag = ""
        if stretch > 0.01:
            flag = f" 拉伸{stretch*100:.1f}%!"
            bad += 1
        if sh.top < 0 or sh.top + sh.height > H or sh.left < 0 or sh.left + sh.width > W:
            flag += " 越界!"
            bad += 1
        line.append(f"{w}x{h}(src {iw}x{ih}, r{ar_put:.3f}){flag}")
    print(f"  P{i:02d}: " + " | ".join(line))
print(f"[vertical] 拉伸/越界问题: {bad} 处")

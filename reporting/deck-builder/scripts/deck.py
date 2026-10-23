#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""deck-builder: fixed-canvas deck pipeline (PDF -> JPG -> PPTX).

  render   PDF  -> JPG pages at a target width (single lossy step)
  assemble JPG folder -> PPTX at the deck canvas size (one full-bleed image per slide)
  check    verify images / pptx against the canvas spec
             --mode images (default): assembled deck, full-bleed picture per slide
             --mode native: editable native master, geometry checks (verify_master)

Canvas spec (Pivot Studio report deck):
  54.98 x 30.92 cm  =  1558.2 x 876.48 pt  =  16:9  =  3247 x 1826 px @150dpi

Examples
  python deck.py render   --pdf report.pdf --out jpg --width 2560 --quality 82
  python deck.py render   --pdf report.pdf --out jpg --dpi 150
  python deck.py assemble --images jpg --out report_deck.pptx
  python deck.py check    --images jpg
  python deck.py check    --pptx report_deck.pptx

Dependencies: pymupdf, pillow (render); python-pptx (assemble/check-pptx).
  pip install pymupdf pillow python-pptx
"""
from __future__ import annotations

import argparse
import glob
import os
import sys

CANVAS_CM = (54.98, 30.92)
CANVAS_PT = (1558.2, 876.48)
SLIDE_EMU = (19789775, 11131550)  # exact value from the existing deck, ratio 1.7778
RATIO = 16.0 / 9.0
EMU_PER_CM = 360000


def _need(module: str, hint: str):
    try:
        return __import__(module)
    except ImportError:
        sys.exit(f"[deck] missing python package '{module}'. Install: {hint}")


def _page_files(folder: str):
    exts = (".jpg", ".jpeg", ".png")
    files = [f for f in glob.glob(os.path.join(folder, "*")) if f.lower().endswith(exts)]
    return sorted(files)


def cmd_render(a) -> int:
    pymupdf = _need("pymupdf", "pip install pymupdf")
    _need("PIL", "pip install pillow")
    from PIL import Image

    doc = pymupdf.open(a.pdf)
    os.makedirs(a.out, exist_ok=True)
    width = a.width
    if a.dpi:
        width = int(round(CANVAS_CM[0] / 2.54 * a.dpi))
    total = 0
    first = None
    for i, page in enumerate(doc, 1):
        zoom = width / page.rect.width
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
        im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        path = os.path.join(a.out, f"{a.prefix}{i:02d}.jpg")
        im.save(path, "JPEG", quality=a.quality, optimize=True, progressive=True, subsampling=2)
        total += os.path.getsize(path)
        first = first or (pix.width, pix.height)
    print(f"[deck] rendered {doc.page_count} pages -> {a.out}")
    print(f"[deck] size {first[0]}x{first[1]}px  q{a.quality}  "
          f"total {total / 1048576:.2f} MB  avg {total / doc.page_count / 1024:.0f} KB/page")
    pages = [p for p in doc]
    ratios = {(round(p.rect.width / p.rect.height, 4)) for p in pages}
    print(f"[deck] page ratios: {sorted(ratios)}  (canvas {RATIO:.4f})")
    return 0


def _emu_from_cm(v: float) -> int:
    return int(round(v * EMU_PER_CM))


def cmd_assemble(a) -> int:
    pptx = _need("pptx", "pip install python-pptx")
    from pptx import Presentation
    from pptx.util import Emu

    files = _page_files(a.images)
    if not files:
        sys.exit(f"[deck] no images in {a.images}")

    if a.width_cm is None and a.height_cm is None:
        w, h = SLIDE_EMU
    else:
        w = _emu_from_cm(a.width_cm if a.width_cm else CANVAS_CM[0])
        h = _emu_from_cm(a.height_cm if a.height_cm else CANVAS_CM[1])
    prs = Presentation()
    prs.slide_width = Emu(w)
    prs.slide_height = Emu(h)
    blank = prs.slide_layouts[6]

    for f in files:
        slide = prs.slides.add_slide(blank)
        pic = slide.shapes.add_picture(f, 0, 0, width=Emu(w), height=Emu(h))
        stem = os.path.splitext(os.path.basename(f))[0]
        # alt text: keeps the deck searchable and screen-reader friendly
        pic._element.nvPicPr.cNvPr.set("descr", a.alt_prefix + stem)
        if a.speaker_notes:
            slide.notes_slide.notes_text_frame.text = stem

    os.makedirs(os.path.dirname(os.path.abspath(a.out)) or ".", exist_ok=True)
    prs.save(a.out)
    print(f"[deck] {len(files)} slides -> {a.out}  "
          f"{os.path.getsize(a.out) / 1048576:.2f} MB  canvas {w / EMU_PER_CM:.2f}x{h / EMU_PER_CM:.2f} cm")
    return 0


def cmd_check(a) -> int:
    ratio = 9.0 / 16.0 if a.canvas == "9:16" else RATIO
    rc = 0
    if a.images:
        _need("PIL", "pip install pillow")
        from PIL import Image
        files = _page_files(a.images)
        if not files:
            sys.exit(f"[deck] no images in {a.images}")
        sizes, total, worst = set(), 0, ("", 0)
        for f in files:
            im = Image.open(f)
            sizes.add(im.size)
            n = os.path.getsize(f)
            total += n
            if n > worst[1]:
                worst = (f, n)
            r = im.size[0] / im.size[1]
            if abs(r - ratio) / ratio > 0.005:
                print(f"  ! ratio {im.size} ({r:.4f}) in {os.path.basename(f)}")
                rc = 1
        print(f"[deck] images: {len(files)}  sizes {sorted(sizes)}  "
              f"total {total / 1048576:.2f} MB  avg {total / len(files) / 1024:.0f} KB  "
              f"max {worst[1] / 1024:.0f} KB  canvas {a.canvas}")
        if len(sizes) > 1:
            print("  ! mixed pixel sizes across pages")
            rc = 1
    if a.pptx:
        if a.mode == "native":
            # 原生母版（可编辑形状/文本框）：几何校验，横版判据在 verify_master.verify()，
            # 竖版在 verify_vertical.check()；判据与各自独立命令行完全一致
            sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
            if a.canvas == "9:16":
                from verify_vertical import check as verify_vertical
                rc = verify_vertical(a.pptx) or rc
            else:
                from verify_master import verify
                rc = verify(a.pptx) or rc
        else:
            # 图片拼装模式：画布比例 + 每页一张满幅图
            pptx = _need("pptx", "pip install python-pptx")
            from pptx import Presentation
            prs = Presentation(a.pptx)
            w, h = prs.slide_width, prs.slide_height
            r = w / h
            print(f"[deck] {a.pptx}: {len(prs.slides)} slides  "
                  f"{w / 360000:.2f}x{h / 360000:.2f} cm  ratio {r:.4f}  "
                  f"{os.path.getsize(a.pptx) / 1048576:.2f} MB  canvas {a.canvas}")
            if abs(r - ratio) / ratio > 0.005:
                print(f"  ! slide ratio is not {a.canvas}")
                rc = 1
            for i, s in enumerate(prs.slides, 1):
                pics = [sh for sh in s.shapes if sh.shape_type == 13]
                cover = [p for p in pics if abs(p.width - w) < 36000 and abs(p.height - h) < 36000]
                if not cover:
                    print(f"  ! slide {i}: no full-bleed picture ({len(pics)} pictures)")
                    rc = 1
    print("[deck] check", "FAILED" if rc else "OK")
    return rc


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="deck", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    r = sub.add_parser("render", help="PDF -> JPG pages")
    r.add_argument("--pdf", required=True)
    r.add_argument("--out", required=True)
    r.add_argument("--width", type=int, default=2560, help="pixel width (default 2560)")
    r.add_argument("--dpi", type=int, default=0, help="or compute width from dpi (e.g. 150)")
    r.add_argument("--quality", type=int, default=82)
    r.add_argument("--prefix", default="")
    r.set_defaults(func=cmd_render)

    s = sub.add_parser("assemble", help="JPG folder -> PPTX")
    s.add_argument("--images", required=True)
    s.add_argument("--out", required=True)
    s.add_argument("--width-cm", type=float, default=None)
    s.add_argument("--height-cm", type=float, default=None)
    s.add_argument("--alt-prefix", default="")
    s.add_argument("--speaker-notes", action="store_true")
    s.set_defaults(func=cmd_assemble)

    c = sub.add_parser("check", help="verify images / pptx")
    c.add_argument("--images")
    c.add_argument("--pptx")
    c.add_argument("--mode", choices=("images", "native"), default="images",
                   help="pptx check mode: images = assembled deck (full-bleed per slide), "
                        "native = editable master (geometry via verify_master/verify_vertical)")
    c.add_argument("--canvas", choices=("16:9", "9:16"), default="16:9",
                   help="target canvas ratio (default 16:9; use 9:16 for vertical decks)")
    c.set_defaults(func=cmd_check)

    a = ap.parse_args(argv)
    if a.cmd == "check" and not (a.images or a.pptx):
        ap.error("check needs --images and/or --pptx")
    return a.func(a)


if __name__ == "__main__":
    raise SystemExit(main())

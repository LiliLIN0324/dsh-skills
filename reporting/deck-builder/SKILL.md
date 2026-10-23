---
name: deck-builder
description: 把 Word/PPT/PDF 报告稳定产出一套「同尺寸、同体积、同质感」的图片集 + PPTX。固定画布 54.98×30.92cm(16:9)，一次栅格化，python-pptx 组装，可机器校验。用于 Pivot Studio 个人项目报告（1037 拼拼这类）的批量交付。
---

# deck-builder

固定画布出图管线：**docx → 矢量 PDF → JPG（一次压缩）→ PPTX（脚本组装）→ 自动校验**。

## 什么时候用

- 用户要把一份 Word / PPT / PDF 报告导成「一页一图」的 PPTX 或图片集
- 多个项目要做得尺寸、页边距、字号、文件体积完全一致
- 现有图太大（>500KB/页）或 PPTX 太重（>10MB），要压到能发的体积
- 交付前要检查：页比例、像素尺寸、单页体积、是否每页都满幅

## 画布规格（锁死，所有项目共用）

| 项 | 值 |
|---|---|
| 页面 | 54.98 × 30.92 cm = 1558.2 × 876.48 pt = 16:9（比例 1.7778） |
| 分享级图 | 2560 × 1440 px，JPEG **q82** progressive，4:2:0 → 约 200KB/页 |
| 存档级图 | 3247 × 1826 px（150dpi），JPEG q85 → 约 320KB/页 |
| PPTX 画布 | 19789775 × 11131550 EMU（54.97 × 30.92 cm） |
| Word 页设置 | 纸张自定义 54.98 × 30.92 cm；边距 左右 3.17cm、上下 2.54cm |

## 三步跑完

### 1. docx → 矢量 PDF（WPS COM，不开界面）

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/export_pdf.ps1 -Docx "C:\path\报告.docx" -Out "C:\path\out\report.pdf"
```

验证是矢量：`python -c "import pymupdf;d=pymupdf.open('report.pdf');print(len(d[0].get_text()))"` → 应 > 0。
若为 0，说明 PDF 是图片打包出来的，整条链都会糊，先解决导出这一步。

### 2. PDF → JPG

```powershell
python scripts/deck.py render --pdf report.pdf --out jpg --width 2560 --quality 82
# 存档： --dpi 150 --quality 85
```

### 3. JPG → PPTX + 校验

```powershell
python scripts/deck.py assemble --images jpg --out report_deck.pptx --speaker-notes
python scripts/deck.py check --images jpg
python scripts/deck.py check --pptx report_deck.pptx                # 图片拼装模式（默认）
python scripts/deck.py check --pptx master.pptx --mode native       # 原生母版（几何校验）
```

**check 的两种 pptx 模式别用混**：
- `--mode images`（默认）：拼装出来的 deck，判据 = 画布比例 + 每页一张满幅图。拿它查原生母版必报 FAILED，属正常。
- `--mode native`：原生母版（可编辑形状/文本框）。横版判据在 `verify_master.verify()`（比例、越界、文本估行溢出、无整页图充版面）；竖版在 `verify_vertical.check()`，除上述外还有规格第 9 节版式判据：**一屏视觉主体 ≤1、并排 ≤2 张且每张宽 ≥420 或高 ≥840**（媒体图聚类：纵向重叠并排算同主体，同宽全宽竖叠也算 1 个主体）、**图内联数**（机器标注 `cNvPr@title` 写 `image_cols=N`，descr/alt 留给人类描述，判 936/N ≥ 420）。同一套判据也有独立命令行 `python scripts/verify_master.py <pptx>`、`python scripts/verify_vertical.py <pptx>`。
- 竖版（1080×1920）必须加 `--canvas 9:16`，否则 16:9 判据会误报；横版默认即可。
- 规格第 9.7 条「源图缩到 936 宽后图内最小字号 ≥18px」是**素材级规则，校验器判不了**（机器看不了图内容）：check 全绿 ≠ 图可读，这层靠人审素材；后续若加标注通道（如图内最小字号注记）再进机器。
两种模式都要求最终输出 `[deck] check OK` / `[verify] OK` / `[verify-v] OK`。

依赖：`pip install pymupdf pillow python-pptx`（WPS 需已安装，用 `Kwps.Application`）。

**用哪个 python**：PATH 里的用户级 `python`（`C:\Users\13266\AppData\Local\Programs\Python\Python312\python.exe`，3.12.10），已装好 pymupdf 1.28.2 / python-pptx 1.0.2 / Pillow 13.0.0，**不需要设 PYTHONPATH**。dsh runtime 自带的 python 没有 pymupdf，别用它跑本技能；也**不要往 dsh runtime 的 site-packages 装东西**（随更新重置，见 `_docs/install-pymupdf.md`）。

### 4（可选）PPTX → PDF，作品质核对

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/export_pptx_pdf.ps1 -Pptx deck.pptx -Out deck.pdf
```

WPS 的 `Kwpp.Application` 只认 `Presentation.SaveAs($out, 32)`（32 = ppSaveAsPDF）；`ExportAsFixedFormat` 会报 `HRESULT E_FAIL`。
`-Pptx`/`-Out` 传相对路径也行（脚本入口统一转绝对；以前直接传 COM 会 E_FAIL，已修）。

## 原生 PPTX 母版（比「一页一图」更好的交付形态）

一页一图只适合「已经定稿、不再改」的交付。**还在改内容的报告，母版要做成原生 PPTX**：真实文本框 + 形状 + 表格，改字双击即改，导出即所见。
版式规格见 `_docs/master-layout-spec.md`（6 个横版 + 1 个竖版 9:16），实现在 `scripts/build_master_16x9.py` 与 `scripts/build_vertical_9x16.py`（python-pptx 生成，可直接改）。
铁律：母版里**不许拿整页图片充版面**；要出图时走上面的 render，不要反向拿 JPG 拼页。

### 视频（成品，不是中间件）

页面 JPG → MP4，ffmpeg：`D:\dsh-tools\bin\ffmpeg.cmd`。
横 1920×1080 / 竖 1080×1920，30fps，h264 crf21 yuv420p faststart；每页停留 4s（封面/结尾 5s），转场 0.4s 交叉淡入；BGM −18dB。
实测 13 页 56s：横 2.05MB、竖 1.37MB。竖版**必须**用 9:16 单独版式渲染后再编码，不能把横版缩放塞进去。

## 验收阈值

- 单页 ≤ 350KB（2560 q82）；总图集 ≤ 5MB / 15页
- 所有页像素尺寸一致，比例 1.7778 ±0.5%
- PPTX ≤ 4MB / 20页，每页一张满幅图，且带 alt 文本
- 交付版 PDF 有文字层（可搜索、可复制）
- `check` 输出必须是 `OK`

## 已知坑

- **不要 JPG→JPG**：二次压缩会把文字边缘糊掉，永远从 PDF 重新栅格。
- **Word 页尺寸必须是 16:9**：`55.01×30.8cm`（1.786）比画布宽 0.4%，进 PPTX 会被纵向拉伸。
- **pptx 全图铺满 = 没有文字层**：要可搜索就同时交 PDF；PPTX 里用 alt 文本兜底。
- **WPS COM 导出**：文档先确保已保存且未被占用；导出后 `$doc.Close(0); $app.Quit()` 收尾。
- 页面里若混入竖版（A4 1.414）图，塞进 16:9 会被拉伸/裁切，先按画布重排再出图。
- **export_pdf.ps1 打印 `OK: ...` 后进程可能卡在 COM 清理几分钟不退出**（10-19 实测）：导出已成功，自动化调用设超时，看到 OK 即成功，不必等进程退出。
- 用户开着 WPS 时慎跑 export_pdf.ps1：COM 可能附着已开实例，`Quit()` 有连带风险。
- **原生母版里满幅底矩形的 EMU 必须与画布完全相等**：用 `Cm(54.98)`（=19792800）会比锁死画布 19789775 大 3025 EMU，被 `check --mode native` 判 OUT-OF-CANVAS。母版脚本里画布尺寸用同一常量，别手写 Cm()。
- **WPS COM（Kwps/Kwpp）Open 不认相对路径**，报 HRESULT E_FAIL——必须传绝对路径；脚本 `export_pdf.ps1`/`export_pptx_pdf.ps1` 入口已统一转绝对，直接调 COM 时自己转。
- 完整验收记录（四环节真实输出、卸载方式）见 `_docs/skills-deck-builder.md`。

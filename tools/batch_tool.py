#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量文件处理器 —— AI 自动化交付样例
把「重复的体力活」变成一条命令：

  1. 批量重命名（按规则加前缀/序号/日期/替换关键词）
  2. 批量转换（PDF→图片、图片→PDF、CSV→Excel、图片压缩）
  3. 批量加水印 / 统一尺寸 / EXIF 清除
  4. 批量提取文本（PDF/Word/Excel→txt）合并成一个文件

依赖：pip install pymupdf pillow python-docx openpyxl pypdf
用法：python batch_tool.py --help
"""
import argparse
import csv
import datetime as dt
import io
import os
import re
import shutil
import sys
from pathlib import Path

# ---------------- 依赖按需导入 ----------------
def need(mod):
    """按需导入依赖。mod 可以是 'PIL.Image' 这种带子模块的路径。"""
    try:
        return __import__(mod, fromlist=["*"])
    except ImportError:
        top = mod.split(".")[0]
        print(f"[!] 缺少依赖 {top}，请先 pip install {top}", file=sys.stderr)
        raise SystemExit(2)


# ---------------- 重命名 ----------------
def do_rename(files, outdir, prefix, seq_start, digits, date_fmt, find, replace, keep_ext=True):
    """按规则重命名：prefix + 可选日期 + 序号，可做关键词替换"""
    outdir = Path(outdir); outdir.mkdir(parents=True, exist_ok=True)
    today = dt.datetime.now().strftime(date_fmt) if date_fmt else ""
    used = set()
    results = []
    for i, f in enumerate(files):
        stem, ext = Path(f).stem, (Path(f).suffix if keep_ext else "")
        if find:
            stem = stem.replace(find, replace)
        parts = [p for p in [prefix, today] if p]
        base = "_".join(parts)
        name = f"{base}_{str(seq_start+i).zfill(digits)}{ext}" if parts else f"{stem}{ext}"
        # 冲突自动加序号
        cand, n = name, 1
        while cand in used or (outdir / cand).exists():
            stem2, ext2 = Path(name).stem, Path(name).suffix
            name = f"{stem2}_{n}{ext2}"; n += 1
        used.add(name)
        shutil.copy2(f, outdir / name)
        results.append((str(f), str(outdir / name)))
    return results


# ---------------- 转换 ----------------
def pdf_to_images(pdf_path, outdir, dpi=150, prefix=""):
    fitz = need("fitz")
    outdir = Path(outdir); outdir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf_path)
    outs = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        p = outdir / f"{prefix}{Path(pdf_path).stem}_p{i+1:03d}.png"
        pix.save(p); outs.append(str(p))
    doc.close()
    return outs


def images_to_pdf(img_paths, out_pdf):
    PIL = need("PIL.Image")
    imgs = []
    for p in img_paths:
        im = PIL.open(p).convert("RGB")
        imgs.append(im)
    if not imgs:
        raise SystemExit("没有图片可转换")
    imgs[0].save(out_pdf, save_all=True, append_images=imgs[1:])
    return str(out_pdf)


def compress_images(files, outdir, quality=70, maxw=1920):
    PIL = need("PIL.Image")
    outdir = Path(outdir); outdir.mkdir(parents=True, exist_ok=True)
    outs = []
    for f in files:
        im = PIL.open(f).convert("RGB")
        if im.width > maxw:
            h = int(im.height * maxw / im.width)
            im = im.resize((maxw, h))
        p = outdir / (Path(f).stem + ".jpg")
        im.save(p, "JPEG", quality=quality, optimize=True)
        outs.append(str(p))
    return outs


def csv_to_excel(csv_path, out_xlsx, sheet="Sheet1"):
    need("openpyxl")
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active; ws.title = sheet
    with open(csv_path, newline="", encoding="utf-8-sig") as fh:
        for row in csv.reader(fh):
            ws.append(row)
    wb.save(out_xlsx)
    return str(out_xlsx)


def watermark_pdfs(files, outdir, text):
    fitz = need("fitz")
    outdir = Path(outdir); outdir.mkdir(parents=True, exist_ok=True)
    outs = []
    for f in files:
        doc = fitz.open(f)
        for page in doc:
            rect = page.rect
            page.insert_text((rect.width/2 - len(text)*4, rect.height/2),
                             text, fontsize=28, color=(0.6, 0.6, 0.6), rotate=0)
        p = outdir / (Path(f).stem + "_wm.pdf")
        doc.save(p); doc.close(); outs.append(str(p))
    return outs


def extract_texts(files, out_txt):
    chunks = []
    for f in files:
        ext = Path(f).suffix.lower()
        try:
            if ext == ".pdf":
                fitz = need("fitz"); d = fitz.open(f)
                chunks.append(f"\n===== {Path(f).name} =====\n" + "\n".join(p.get_text() for p in d)); d.close()
            elif ext in (".docx",):
                need("docx"); import docx
                d = docx.Document(f)
                body = "\n".join(p.text for p in d.paragraphs)
                chunks.append(f"\n===== {Path(f).name} =====\n" + body)
            elif ext in (".xlsx",):
                need("openpyxl"); from openpyxl import load_workbook
                wb = load_workbook(f, data_only=True)
                t = []
                for ws in wb.worksheets:
                    t.append(f"--- sheet: {ws.title} ---")
                    for row in ws.iter_rows(values_only=True):
                        t.append("\t".join("" if v is None else str(v) for v in row))
                chunks.append(f"\n===== {Path(f).name} =====\n" + "\n".join(t))
            elif ext in (".txt", ".md", ".csv"):
                chunks.append(f"\n===== {Path(f).name} =====\n" + Path(f).read_text(encoding="utf-8", errors="ignore"))
            else:
                chunks.append(f"\n===== {Path(f).name} =====\n(跳过不支持的类型 {ext})")
        except Exception as e:
            chunks.append(f"\n===== {Path(f).name} =====\n[读取失败] {e}")
    Path(out_txt).write_text("\n".join(chunks), encoding="utf-8")
    return str(out_txt)


# ---------------- CLI ----------------
def collect(paths, exts=None, recursive=False):
    out = []
    for p in paths:
        P = Path(p)
        if P.is_file():
            out.append(P)
        elif P.is_dir():
            it = P.rglob("*") if recursive else P.glob("*")
            for f in it:
                if f.is_file() and (not exts or f.suffix.lower() in exts):
                    out.append(f)
    return sorted(out, key=lambda x: str(x))


def main():
    ap = argparse.ArgumentParser(description="批量文件处理器（AI 自动化交付样例）")
    sub = ap.add_subparsers(dest="cmd", required=True)

    r = sub.add_parser("rename", help="批量重命名")
    r.add_argument("paths", nargs="+")
    r.add_argument("-o", "--outdir", default="renamed")
    r.add_argument("--prefix", default="")
    r.add_argument("--seq-start", type=int, default=1)
    r.add_argument("--digits", type=int, default=3)
    r.add_argument("--date", default="", help="日期格式，如 %%Y%%m%%d")
    r.add_argument("--find", default="")
    r.add_argument("--replace", default="")
    r.add_argument("-r", "--recursive", action="store_true")

    c = sub.add_parser("convert", help="格式转换")
    c.add_argument("mode", choices=["pdf2img", "img2pdf", "csv2xlsx", "compress"])
    c.add_argument("paths", nargs="+")
    c.add_argument("-o", "--out")
    c.add_argument("--dpi", type=int, default=150)
    c.add_argument("--quality", type=int, default=70)
    c.add_argument("--maxw", type=int, default=1920)
    c.add_argument("-r", "--recursive", action="store_true")

    w = sub.add_parser("watermark", help="PDF 批量加水印")
    w.add_argument("paths", nargs="+")
    w.add_argument("-o", "--outdir", default="watermarked")
    w.add_argument("--text", default="仅供内部使用")
    w.add_argument("-r", "--recursive", action="store_true")

    e = sub.add_parser("extract", help="批量提取文本合并")
    e.add_argument("paths", nargs="+")
    e.add_argument("-o", "--out", default="merged.txt")
    e.add_argument("-r", "--recursive", action="store_true")

    a = ap.parse_args()

    if a.cmd == "rename":
        files = collect(a.paths, recursive=a.recursive)
        res = do_rename(files, a.outdir, a.prefix, a.seq_start, a.digits,
                        a.date, a.find, a.replace)
        for s, d in res:
            print(f"{s}  ->  {d}")
        print(f"[✓] 共处理 {len(res)} 个文件")

    elif a.cmd == "convert":
        if a.mode == "pdf2img":
            files = collect(a.paths, [".pdf"], recursive=a.recursive)
            out = a.out or "images"; outs = []
            for f in files:
                outs += pdf_to_images(f, out, dpi=a.dpi)
            for o in outs: print(o)
            print(f"[✓] {len(outs)} 张图片")
        elif a.mode == "img2pdf":
            files = collect(a.paths, [".jpg", ".jpeg", ".png", ".webp"], recursive=a.recursive)
            out = a.out or "merged.pdf"
            print(images_to_pdf(files, out)); print(f"[✓] {len(files)} 张图")
        elif a.mode == "csv2xlsx":
            files = collect(a.paths, [".csv"], recursive=a.recursive)
            out = a.out or "out.xlsx"
            for f in files:
                print(csv_to_excel(f, out))
            print("[✓] 完成")
        else:  # compress
            files = collect(a.paths, [".jpg", ".jpeg", ".png", ".webp"], recursive=a.recursive)
            out = a.out or "compressed"
            outs = compress_images(files, out, quality=a.quality, maxw=a.maxw)
            for o in outs: print(o)
            print(f"[✓] {len(outs)} 张图片已压缩")

    elif a.cmd == "watermark":
        files = collect(a.paths, [".pdf"], recursive=a.recursive)
        outs = watermark_pdfs(files, a.outdir, a.text)
        for o in outs: print(o)
        print(f"[✓] {len(outs)} 个 PDF")

    elif a.cmd == "extract":
        files = collect(a.paths, recursive=a.recursive)
        print(extract_texts(files, a.out)); print(f"[✓] {len(files)} 个文件已合并")


if __name__ == "__main__":
    main()

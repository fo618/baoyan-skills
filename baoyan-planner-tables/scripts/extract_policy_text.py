# -*- coding: utf-8 -*-
"""
推免政策 PDF → 纯文本（形态 C 的第 0 步）

把院校推免/保研政策 PDF 抽成纯文本并打印到终端，供解读撰写前逐条核对原文。
支持扫描件兜底提示（若抽不出文字，会提示改用 OCR 或找网页版原文）。

用法:
  python extract_policy_text.py <政策.pdf> [--out 文本.txt] [--pages 1-6] [--quiet]

选项:
  --out FILE   同时把文本写成 UTF-8 文件（默认 policy_text.txt）
  --pages RANGE 只抽指定页，如 1-6 或 2,4,7（默认全部）
  --quiet      只写文件，不打印正文（仍打印页数与字数统计）

依赖:
  pip install pypdf
"""
import argparse
import os
import sys

try:
    from pypdf import PdfReader
except ImportError:  # 兼容旧包名
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        sys.exit("缺少依赖：请先执行  pip install pypdf")


def parse_pages(spec, total):
    """'1-6' / '2,4,7' → 0-based 页码列表"""
    if not spec:
        return list(range(total))
    out = []
    for part in str(spec).split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.extend(range(int(a) - 1, int(b)))  # 闭区间
        else:
            out.append(int(part) - 1)
    return [p for p in out if 0 <= p < total]


def main():
    ap = argparse.ArgumentParser(description="推免政策 PDF 抽文本")
    ap.add_argument("pdf")
    ap.add_argument("--out", nargs="?", const="policy_text.txt", default=None)
    ap.add_argument("--pages", default=None)
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    if not os.path.isfile(args.pdf):
        sys.exit("找不到文件：%s" % args.pdf)

    reader = PdfReader(args.pdf)
    total = len(reader.pages)
    pages = parse_pages(args.pages, total)

    chunks = []
    for i in pages:
        try:
            text = reader.pages[i].extract_text() or ""
        except Exception as e:
            text = "[第 %d 页抽取失败：%s]" % (i + 1, e)
        chunks.append("===== PAGE %d =====\n%s" % (i + 1, text))

    body = "\n\n".join(chunks).strip()
    chars = len(body.replace("\n", ""))

    print("文件：%s" % os.path.basename(args.pdf))
    print("总页数：%d ｜ 已抽取：%d 页 ｜ 字符数：%d" % (total, len(pages), chars))
    if chars < 200:
        print("⚠ 文字极少 —— 多半是扫描件/图片版 PDF，请改用 OCR，或去院校官网/学信网找网页版原文。")
    if not args.quiet:
        print("-" * 60)
        print(body)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(body)
        print("已写入：%s" % args.out)


if __name__ == "__main__":
    main()

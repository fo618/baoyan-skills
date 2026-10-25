# -*- coding: utf-8 -*-
"""
从中国研究生招生信息网（研招网 yz.chsi.com.cn）抓取**当年最新**的《全国硕士研究生招生工作管理规定》
（即包含推荐免试/保研条款的全国性政策文件），**默认生成自包含 HTML 交付件**落到桌面。

**执行约定（2026-09-29 八次定版）：本脚本已下线，默认流程里不要跑。**
用户明确要求"以后只需要生成本校和本学院的政策文件就行，不要《全国硕士研究生招生工作管理规定》"——
出保研 PPT 时政策交付物只有院校级 HTML（见 SKILL.md 铁律 19e，脚本 fetch_school_tuimian_policy.py）。
**只有用户在当次对话里单独点名要全国规定时才跑本脚本**（见 SKILL.md 铁律 19d）。
（本文件仍被 fetch_school_tuimian_policy.py 依赖：复用其中的 write_html() / desk_dir() / safe_name()，勿删。）

**交付形态（2026-09-28 用户定版）**：只交付 HTML，用户自己在浏览器里 Ctrl+P 另存为 PDF。
原因：自研重排 PDF 会破坏原文的表格与缩进，用户明确反馈"排版太乱"。HTML 原样保留原文富文本，
打印出来的才是最接近原件的版本。加 `--pdf` 才会额外生成重排版 PDF（仅作备选，不推荐交付）。

用法：
    python fetch_yanzhaowang_policy.py                 # 自动定位最新版本 → 输出 HTML 到桌面
    python fetch_yanzhaowang_policy.py --year 2027     # 指定年份
    python fetch_yanzhaowang_policy.py --url <页面URL> # 直接指定页面
    python fetch_yanzhaowang_policy.py --dir D:\\交付   # 指定输出目录
    python fetch_yanzhaowang_policy.py --dry-run       # 只看抓到了什么，不写文件
    python fetch_yanzhaowang_policy.py --pdf           # 额外再出一份重排 PDF（不推荐）

依赖：仅 urllib（内置）。加 --pdf 时才需要 pymupdf。

注意：教育部/研招网的政策多为 HTML，官方通常不提供 PDF 附件。本脚本抓的是 HTML 全文，
      按原段落原样嵌入交付件，首页写明来源 URL、文号与抓取时间以便核真。
      若某天官方提供了 .pdf 附件，脚本会优先直接下载该原件。
"""

import os
import re
import sys
import html as htmlmod
import datetime
import urllib.request
import argparse

BASE = "https://yz.chsi.com.cn"
HOME = BASE + "/"
COLUMN = BASE + "/kyzx/jybzc/"          # 教育部政策栏目页
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/122.0 Safari/537.36")

TITLE_RE = re.compile(r"(20\d{2})年全国硕士研究生招生工作管理规定")


# --------------------------------------------------------------------------- 抓取
def fetch(url, timeout=25):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Language": "zh-CN,zh"})
    raw = urllib.request.urlopen(req, timeout=timeout).read()
    for enc in ("utf-8", "gb18030"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "ignore")


def abs_url(href):
    if href.startswith("http"):
        return href
    if href.startswith("//"):
        return "https:" + href
    return BASE + ("" if href.startswith("/") else "/") + href


def find_latest_url(year=None):
    """从首页/栏目页里找最新的《XXXX年全国硕士研究生招生工作管理规定》正文页。"""
    cands = {}
    for page in (HOME, COLUMN):
        try:
            html = fetch(page)
        except Exception as e:
            print(f"  ⚠ 抓取失败 {page}：{e}")
            continue
        for m in re.finditer(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', html, re.S):
            href, text = m.group(1), htmlmod.unescape(re.sub(r"<[^>]+>", "", m.group(2))).strip()
            ym = TITLE_RE.search(text)
            if not ym:
                continue
            y = int(ym.group(1))
            url = abs_url(href)
            if url not in cands:
                cands[url] = y
    if not cands:
        return None, None
    pool = cands.items()
    if year:
        pool = [(u, y) for u, y in pool if y == year] or cands.items()
    url, y = sorted(pool, key=lambda kv: -kv[1])[0]
    return url, y


# --------------------------------------------------------------------------- 抽取
TAG_P = re.compile(r"<p\b[^>]*>(.*?)</p>", re.S)
TAG_ANY = re.compile(r"<[^>]+>")


def clean(s):
    s = re.sub(r"&nbsp;|&#160;", " ", s)
    s = TAG_ANY.sub("", s)
    return htmlmod.unescape(s).replace("\u3000", " ").strip()


def extract(html, year):
    """抽出通知抬头 + 规定全文段落。"""
    # 同时保留「原始 HTML 段」与「纯文本段」，前者用于生成高保真 HTML 交付件
    pairs = []
    for raw in TAG_P.findall(html):
        txt = clean(raw)
        if txt:
            pairs.append((raw, txt))
    ps = [t for _, t in pairs]

    # 起点：优先「文号」段；其次「各省、自治区」抬头段；最后标题段
    start = None
    for i, p in enumerate(ps):
        if re.search(r"教学〔[^〕]+〕\d+号", p) or p.startswith("各省、自治区、直辖市"):
            start = i
            break
    if start is None:
        for i, p in enumerate(ps):
            if f"{year}年全国硕士研究生招生工作管理规定" in p:
                start = i
                break
    if start is None:
        # 退化方案：从第一条起的正文
        for i, p in enumerate(ps):
            if p.startswith("第一条"):
                start = i
                break
    if start is None:
        return None

    # 终点：含「自印发之日起施行」的段落（含）；找不到则到尾
    end = len(ps)
    for i in range(start, len(ps)):
        if "自印发之日起施行" in ps[i]:
            end = i + 1
            break

    body = pairs[start:end]

    # 过滤掉明显不属于正文的成分（导航、版权、推荐阅读）
    drop_kw = ("上一篇", "下一篇", "相关推荐", "版权所有", "扫一扫", "返回顶部", "责编")
    body = [q for q in body if not any(k in q[1] for k in drop_kw)]

    title = f"{year}年全国硕士研究生招生工作管理规定"
    docno = ""
    for _, p in body[:6]:
        m = re.search(r"教学〔[^〕]+〕\d+号", p)
        if m:
            docno = m.group(0)
            break
    return {"title": title, "docno": docno, "year": year,
            "paras": [t for _, t in body],
            "paras_html": [r for r, _ in body]}


# --------------------------------------------------------------------------- HTML（默认交付件）
HTML_SHELL = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>《{title}》</title>
<style>
 @page {{ size: A4; margin: 18mm 16mm 16mm; }}
 * {{ box-sizing: border-box; }}
 body {{ font-family:"Microsoft YaHei","微软雅黑","SimSun","宋体",serif; font-size:11.5pt;
        line-height:1.8; color:#1a1a1a; background:#fff; max-width:900px; margin:0 auto;
        padding:28px 24px 60px; }}
 h1 {{ font-size:18pt; text-align:center; margin:0 0 8px; letter-spacing:1px; }}
 .meta {{ text-align:center; color:#5a5a5a; font-size:9.5pt; margin-bottom:10px; }}
 .srcbar {{ border-top:1px solid #d0d0d0; border-bottom:1px solid #d0d0d0; padding:7px 0;
            margin:0 0 18px; font-size:8.5pt; color:#666; word-break:break-all; line-height:1.6; }}
 .tip {{ background:#fbf6f6; border-left:3px solid #9e2f2f; padding:8px 12px; font-size:9pt;
         color:#6b2c2c; margin-bottom:18px; }}
 .content p {{ margin:0 0 7px; text-align:justify; }}
 .content p.doctitle {{ font-size:15pt; text-align:center; font-weight:bold; margin:16px 0 12px; }}
 .content p.chapter  {{ font-size:12.5pt; text-align:center; font-weight:bold; margin:16px 0 8px; }}
 .content p.article  {{ margin:0 0 7px; }}
 .content p.sign     {{ text-align:right; margin:12px 0 0; }}
 .content table {{ border-collapse:collapse; width:100% !important; margin:10px 0; font-size:10pt; }}
 .content td, .content th {{ border:1px solid #7a7a7a !important; padding:5px 7px;
                             word-break:break-word; }}
 .content img {{ max-width:100%; }}
 .content a {{ color:#9e2f2f; }}
 @media print {{
   body {{ padding:0; max-width:none; font-size:11pt; }}
   .noprint {{ display:none !important; }}
   .content a {{ color:#000; text-decoration:none; }}
 }}
</style>
</head>
<body>
<h1>《{title}》</h1>
<div class="meta">{meta}</div>
<div class="srcbar">来源：{source_url}<br>抓取时间：{stamp}　|　条文以教育部 / 研招网官方页面原文为准</div>
<div class="tip noprint">打印建议：Ctrl（Mac 为 ⌘）+ P → 目标选「另存为 PDF」→ 纸张 A4、边距默认、勾选「背景图形」。</div>
<div class="content">
{body}
</div>
</body>
</html>
"""


def to_body_html(data):
    """把抽取结果拼成高保真正文 HTML：优先用原文段落 HTML（保留加粗/表格），
    并给「文件标题 / 章 / 条 / 落款」加样式类，避免交付件变成一坨连排文字。"""
    raws = data.get("paras_html") or []
    out = []
    for i, txt in enumerate(data["paras"]):
        inner = raws[i] if i < len(raws) else txt
        if not re.search(r"^\s*<(p|div|h[1-6]|table|ul|ol|blockquote)\b", inner or "", re.I):
            cls = ""
            if f"{data['year']}年全国硕士研究生招生工作管理规定" in txt and len(txt) < 40:
                cls = "doctitle"
            elif re.match(r"^第[一二三四五六七八九十]+章", txt) and len(txt) < 40:
                cls = "chapter"
            elif re.match(r"^第[一二三四五六七八九十百零〇\d]+条", txt):
                cls = "article"
            elif len(txt) < 26 and re.match(r"^20\d{2}年\d{1,2}月\d{1,2}日$", txt):
                cls = "sign"
            out.append(f'<p class="{cls}">{inner}</p>' if cls else f"<p>{inner}</p>")
        else:
            out.append(inner)
    return "\n".join(out)


def write_html(out_path, title, meta, source_url, body_html):
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    doc = HTML_SHELL.format(title=title, meta=meta, source_url=source_url,
                            stamp=stamp, body=body_html)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(doc)
    return os.path.getsize(out_path)


# --------------------------------------------------------------------------- PDF
def desk_dir():
    for k in ("USERPROFILE", "HOME"):
        base = os.environ.get(k)
        if not base:
            continue
        for cand in (os.path.join(base, "Desktop"), os.path.join(base, "OneDrive", "Desktop"),
                     os.path.join(base, "桌面")):
            if os.path.isdir(cand):
                return cand
    for k in ("DESKTOP", "OneDriveDesktop"):
        p = os.environ.get(k)
        if p and os.path.isdir(p):
            return p
    return os.getcwd()


def safe_name(s):
    return re.sub(r'[\\/:*?"<>|]', "_", s).strip()


def build_pdf(data, source_url, out_path):
    try:
        import pymupdf as fitz          # PyMuPDF >= 1.22 推荐写法
    except ImportError:
        import fitz

    W, H = fitz.paper_rect("a4").width, fitz.paper_rect("a4").height
    ML, MR, MT, MB = 56, 56, 62, 58
    FONT = "china-s"                    # PyMuPDF 内置 CJK 字体，无需外挂字体文件
    LINE_H = 1.68
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    footer = f"来源：{source_url}    抓取时间：{stamp}    仅供参酌，务必以教育部/研招网原文为准"

    state = {"page": None, "y": MT}

    def newpage():
        page = fitz_doc.new_page(width=W, height=H)
        page.insert_text((ML, H - MB + 26), footer, fontname=FONT, fontsize=7.0, color=(0.45, 0.45, 0.45))
        page.draw_line((ML, H - MB + 14), (W - MR, H - MB + 14), color=(0.75, 0.75, 0.75), width=0.5)
        page.insert_text((ML, MT - 22), f"《{data['title']}》", fontname=FONT, fontsize=7.4,
                         color=(0.35, 0.35, 0.35))
        page.draw_line((ML, MT - 12), (W - MR, MT - 12), color=(0.75, 0.75, 0.75), width=0.5)
        state["page"] = page
        state["y"] = MT

    def wrap(text, maxw, size):
        """按字符宽度贪心折行；中英混排均可用（禁则处理从简）。"""
        lines, cur = [], ""
        for ch in text:
            if ch == "\n":
                lines.append(cur)
                cur = ""
                continue
            probe = cur + ch
            if cur and fitz.get_text_length(probe, fontname=FONT, fontsize=size) > maxw:
                lines.append(cur)
                cur = ch
            else:
                cur = probe
        if cur:
            lines.append(cur)
        return lines

    def write_block(text, size, align="left", gap=5, indent=0, bold=False, gray=None):
        """逐行写入文本块，自动跨页；宋君壕式行写法，规避 insert_textbox 返回值版本差异。"""
        if not text:
            return
        color = (0.14, 0.14, 0.14) if gray is None else gray
        top = state["y"]
        maxw = W - ML - MR - indent
        lh = size * LINE_H
        for line in wrap(text, maxw, size):
            if top + lh > H - MB:
                newpage()
                top = state["y"]
            w = fitz.get_text_length(line, fontname=FONT, fontsize=size)
            if align == "center":
                x = ML + indent + (maxw - w) / 2
            elif align == "right":
                x = ML + indent + (maxw - w)
            else:
                x = ML + indent
            state["page"].insert_text((x, top + size * 0.86), line, fontname=FONT,
                                      fontsize=size, color=color,
                                      render_mode=2 if bold else 0, border_width=0.3 if bold else 0)
            top += lh
        state["y"] = top + gap

    fitz_doc = fitz.open()
    newpage()

    # 封面区：标题 + 文号 + 来源
    write_block(f"《{data['title']}》", 15, align="center", gap=6, bold=True)
    head = []
    if data["docno"]:
        head.append(data["docno"])
    head.append(f"来源：中国研究生招生信息网（{BASE}）")
    head.append(f"抓取时间：{stamp}")
    write_block("      ".join(head), 8.2, align="center", gap=8, gray=(0.4, 0.4, 0.4))
    state["page"].draw_line((ML + 60, state["y"]), (W - MR - 60, state["y"]),
                            color=(0.7, 0.7, 0.7), width=0.6)
    state["y"] += 14

    # 正文
    for p in data["paras"]:
        if f"{data['year']}年全国硕士研究生招生工作管理规定" in p and len(p) < 40:
            write_block(p, 13, align="center", gap=8, bold=True)
        elif re.match(r"^第[一二三四五六七八九十]+章", p) and len(p) < 40:
            write_block(p, 11.4, align="center", gap=7, bold=True)
        elif re.match(r"^第[一二三四五六七八九十百零〇\d]+条", p):
            write_block(p, 9.5, gap=5)
        elif len(p) < 26 and ("教　育　部" in p or re.match(r"^20\d{2}年\d{1,2}月\d{1,2}日$", p)):
            write_block(p, 9.5, align="right", gap=6)
        else:
            write_block(p, 9.5, gap=5)

    fitz_doc.save(out_path, garbage=4, deflate=True)
    fitz_doc.close()
    return os.path.getsize(out_path)


# --------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="抓取研招网最新全国硕士招生（含推免/保研）政策并生成 PDF")
    ap.add_argument("--url", help="正文页面 URL（不给则自动查找最新版）")
    ap.add_argument("--year", type=int, help="指定年份，如 2027")
    ap.add_argument("--dir", help="输出目录（默认桌面）")
    ap.add_argument("--out", help="输出文件名（默认自动生成）")
    ap.add_argument("--pdf", action="store_true",
                    help="额外再生成一份重排 PDF（默认只交付 HTML，重排 PDF 排版不如原文保真）")
    ap.add_argument("--dry-run", action="store_true", help="只预览，不写文件")
    a = ap.parse_args()

    url = a.url
    year = a.year
    if url:
        m = TITLE_RE.search(fetch(url))
        year = int(m.group(1)) if m else (year or datetime.datetime.now().year + 1)
        print(f"指定页面  : {url}")
    else:
        url, year = find_latest_url(year)
        if not url:
            print("❌ 未在研招网找到《XX年全国硕士研究生招生工作管理规定》正文页，请手动指定 --url")
            sys.exit(1)
        print(f"自动定位  : 《{year}年全国硕士研究生招生工作管理规定》")

    page_html = fetch(url)
    print(f"正文页    : {url}")

    # 官方若提供 PDF 附件，优先下载原件
    pdfs = re.findall(r'href="([^"]+\.pdf[^"]*)"', page_html)
    if pdfs and not a.dry_run:
        pdf_url = abs_url(pdfs[0])
        outdir = a.dir or desk_dir()
        os.makedirs(outdir, exist_ok=True)
        name = a.out or safe_name(f"《{year}年全国硕士研究生招生工作管理规定》.pdf")
        dst = os.path.join(outdir, name)
        data = urllib.request.urlopen(urllib.request.Request(pdf_url, headers={"User-Agent": UA}),
                                      timeout=30).read()
        open(dst, "wb").write(data)
        print(f"✅ 官方 PDF 原件已下载：{dst}（{len(data)/1024:.0f} KB）")
        return

    data = extract(page_html, year)
    if not data:
        print("❌ 正文抽取失败：页面结构可能已改版，请人工核对")
        sys.exit(2)

    body_len = sum(len(p) for p in data["paras"])
    print(f"文号      : {data['docno'] or '（未识别到）'}")
    print(f"段落数    : {len(data['paras'])}，正文 {body_len} 字")

    outdir = a.dir or desk_dir()
    base = a.out or safe_name(
        f"《{data['title']}》" + (f"（{data['docno']}）" if data["docno"] else ""))
    dst_html = os.path.join(outdir, base + ".html")

    if a.dry_run:
        print(f"(dry-run) 将输出到: {dst_html}")
        print("---- 前 400 字 ----")
        print("\n".join(data["paras"])[:400])
        return

    os.makedirs(outdir, exist_ok=True)
    meta = "　|　".join(x for x in (data["docno"], "中国研究生招生信息网", "教育部印发") if x)
    size = write_html(dst_html, data["title"], meta, url, to_body_html(data))
    print(f"✅ 已生成 HTML：{dst_html}")
    print(f"   来源：{url}")
    print(f"   大小：{size/1024:.0f} KB（浏览器打开后 Ctrl+P 另存为 PDF 即可）")

    if a.pdf:
        dst_pdf = os.path.join(outdir, base + ".pdf")
        build_pdf(data, url, dst_pdf)
        print(f"✅ 已生成重排 PDF：{dst_pdf}（排版不如 HTML 保真，仅作备选）")


if __name__ == "__main__":
    main()

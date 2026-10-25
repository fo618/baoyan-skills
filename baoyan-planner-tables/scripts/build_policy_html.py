# -*- coding: utf-8 -*-
"""
推免政策解读 HTML 生成器（形态 C：政策文件 → 面向家长和学生的分析表）

输入：政策解读数据 JSON（policy_data_*.json）
输出：单文件、自包含、可直接打印/转 PDF 的深红主题 HTML

适用场景：拿到某校《推荐优秀应届本科毕业生免试攻读研究生工作规定》等政策 PDF/网页原文后，
把它翻译成家长和学生看得懂的"一张明白账"——门槛、一票否决、成绩算法、奖励分全景、
行动路线图、名额、流程、红线、FAQ、待确认清单。

用法:
  python build_policy_html.py <policy_data.json> [输出.html] [--title 覆盖标题]

数据 JSON 结构见 assets/policy_sample_data.json（天津工业大学实例，可直接复制改写）。
主题色与 CSS 固化在本文件 CSS 常量中，改排版改这里，不要手写 HTML。

铁律提醒（脚本只做渲染，内容由撰写方负责）：
  1. "政策原文要求 / 分值 / 否决情形 / 流程"等栏目必须忠实原文，不得臆造；
  2. "白话解读 / 难度 / 性价比 / 战略含义 / 行动路线图"属分析意见，必须在页脚声明非官方口径；
  3. 涉及各学院自行制定的规则，一律写进「待确认清单」，不得写成确定结论。
"""
import argparse
import json
import os
import sys

# ---------------- 主题常量（与 Excel / PPT 形态保持一致，勿随意改动） ----------------
RED = "#b01c2e"
RED_D = "#8a1424"
RED_L = "#fdeef0"
RED_LL = "#fff7f8"
GOLD = "#a8751a"
GREEN = "#1f7a4d"
BLUE = "#1c5a99"

CN_NUM = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
          "十一", "十二", "十三", "十四", "十五"]

# 模块默认标题（JSON 里给了 title 就用 JSON 的）
DEFAULT_TITLES = {
    "highlights": "一页速览",
    "thresholds": "资格门槛：谁有资格报名？",
    "vetoes": "一票否决",
    "formula": "综合成绩怎么算",
    "bonus": "奖励分全景对照表",
    "roadmap": "行动路线图",
    "quota": "名额是怎么分的？",
    "flow": "推免工作流程",
    "risks": "红线与风险清单",
    "faq": "家长最常问的问题",
    "todos": "需要向学院确认的问题",
}


# ---------------- 通用小工具 ----------------
def tag_class(text):
    """把中文难度/性质标签映射成 CSS 类"""
    t = str(text or "")
    if "免初试" in t:
        return "t-hard"
    if "中高" in t or "中等" in t:
        return "t-mid"
    if "较易" in t or "容易" in t or "最易" in t:
        return "t-easy"
    if "特殊" in t:
        return "t-gray"
    if "极高" in t or "很高" in t or "硬" in t:
        return "t-hard"
    if "高" in t:
        return "t-hard"
    if "软性" in t or "参考" in t or "信息" in t:
        return "t-info"
    return "t-gray"


def tag(text, cls=None):
    if not text:
        return ""
    return '<span class="tag %s">%s</span>' % (cls or tag_class(text), text)


def box(kind, title, html="", items=None):
    inner = html or ""
    if items:
        inner += "<ul>" + "".join("<li>%s</li>" % i for i in items) + "</ul>"
    return '<div class="box box-%s"><span class="bt">%s</span>%s</div>' % (kind, title, inner)


def h2(idx, title, note=""):
    s = '<h2>%s、%s' % (CN_NUM[idx] if idx < len(CN_NUM) else str(idx + 1), title)
    if note:
        s += '<small>%s</small>' % note
    s += "</h2>"
    return s


def raw(v):
    return "" if v is None else str(v)


def group_runs(rows):
    """把奖励分行按 group 连续分组；group 为空时不合并"""
    runs = []
    for r in rows:
        g = raw(r.get("group"))
        if g and runs and runs[-1]["group"] == g:
            runs[-1]["rows"].append(r)
        else:
            runs.append({"group": g, "rows": [r], "merge": bool(g)})
    if runs and not runs[-1]["group"]:
        runs[-1]["merge"] = False
    for run in runs:
        if not run["group"]:
            run["merge"] = False
    return runs


# ---------------- 各模块渲染 ----------------
def render_hero(meta):
    school = raw(meta.get("school"))
    name = raw(meta.get("policy_name")) or "推免政策"
    audience = raw(meta.get("audience")) or "家长 · 学生 共读版"
    subtitle = raw(meta.get("subtitle"))
    metas = []
    if name:
        metas.append(("政策名称", name))
    if meta.get("published"):
        metas.append(("发布/公示", raw(meta["published"])))
    if meta.get("source"):
        metas.append(("来源", raw(meta["source"])))
    if meta.get("effective"):
        metas.append(("适用", raw(meta["effective"])))
    meta_html = "".join("<span><b>%s：</b>%s</span>" % (k, v) for k, v in metas)

    title = raw(meta.get("title")) or ("%s 推免（保研）政策解读分析表" % school if school else name)
    sub = subtitle or "把政策原文，翻译成“孩子现在该做什么”的一张明白账"

    return (
        '<header class="hero"><div class="wrap">'
        '<span class="kicker">%s</span>'
        '<h1>%s</h1>'
        '<div class="sub">%s</div>'
        '<div class="meta">%s</div>'
        '</div></header>' % (audience, title, sub, meta_html)
    )


def render_highlights(items):
    cards = []
    for it in items:
        cards.append(
            '<div class="card"><div class="num">%s<span>%s</span></div>'
            '<div class="lbl">%s</div><div class="desc">%s</div></div>'
            % (raw(it.get("num")), raw(it.get("unit")), raw(it.get("label")), raw(it.get("desc")))
        )
    return '<div class="cards">%s</div>' % "".join(cards)


def render_thresholds(sec):
    rows = []
    for r in sec.get("rows", []):
        rows.append(
            "<tr><td class=\"k\">%s</td><td>%s</td><td>%s</td><td class=\"c\">%s</td></tr>"
            % (raw(r.get("dim")), raw(r.get("require")), raw(r.get("plain")), tag(r.get("tag")))
        )
    return (
        '<table><thead><tr><th style="width:130px">审核维度</th><th>政策原文要求</th>'
        '<th style="width:230px">给家长/学生的白话解读</th><th class="c" style="width:90px">性质</th>'
        '</tr></thead><tbody>%s</tbody></table>' % "".join(rows)
    )


def render_vetoes(sec):
    rows = []
    for i, r in enumerate(sec.get("rows", []), 1):
        rows.append(
            "<tr><td class=\"c k\">%s</td><td>%s</td><td>%s</td></tr>"
            % (raw(r.get("no") or i), raw(r.get("item")), raw(r.get("plain")))
        )
    html = (
        '<table><thead><tr><th class="c" style="width:60px">序号</th><th>否决情形</th>'
        '<th style="width:340px">家长需要知道的潜台词</th></tr></thead><tbody>%s</tbody></table>' % "".join(rows)
    )
    w = sec.get("warn")
    if w:
        html += box("danger", raw(w.get("title")), raw(w.get("html")), w.get("items"))
    return html


def render_formula(sec):
    html = '<div class="formula"><div class="eq">%s</div>' % raw(sec.get("eq"))
    if sec.get("ex"):
        html += '<div class="ex">%s</div>' % raw(sec["ex"])
    html += "</div>"
    rows = []
    for r in sec.get("rows", []):
        rows.append(
            "<tr><td class=\"k\">%s</td><td class=\"c\"><span class=\"score\">%s</span></td>"
            "<td>%s</td><td>%s</td></tr>"
            % (raw(r.get("part")), raw(r.get("scale")), raw(r.get("rule")), raw(r.get("meaning")))
        )
    html += (
        '<table><thead><tr><th style="width:110px">组成</th><th class="c" style="width:80px">量级</th>'
        '<th>计算口径（政策原文）</th><th style="width:250px">战略含义</th></tr></thead>'
        '<tbody>%s</tbody></table>' % "".join(rows)
    )
    w = sec.get("warn")
    if w:
        html += box("warn", raw(w.get("title")), raw(w.get("html")), w.get("items"))
    return html


def render_bonus(sec):
    rows = sec.get("rows", [])
    out = []
    for run in group_runs(rows):
        rs = run["rows"]
        n = len(rs)
        notes = [raw(r.get("note")) for r in rs]
        merge_note = sum(1 for x in notes if x) <= 1 and bool(notes[0])

        gcell = ""
        if run["group"]:
            gtxt = run["group"]
            if rs[0].get("group_note"):
                gtxt += '<br><span style="font-weight:400;font-size:12px;color:var(--ink3)">%s</span>' % raw(rs[0]["group_note"])
            if run["merge"] and n > 1:
                gcell = '<td class="k" rowspan="%d">%s</td>' % (n, gtxt)
            else:
                gcell = '<td class="k">%s</td>' % gtxt

        for i, r in enumerate(rs):
            cells = []
            if i == 0:
                cells.append(gcell)
            elif run["merge"] and n > 1:
                pass  # 已被上一行的 rowspan 覆盖
            else:
                cells.append('<td class="k">%s</td>' % raw(r.get("group")))

            cells.append("<td>%s</td>" % raw(r.get("item")))
            if r.get("score_tag"):          # 特殊行（如"免初试"）用徽标而非分值
                cells.append('<td class="c">%s</td>' % tag(r.get("score"), "t-hard"))
            else:
                cells.append('<td class="c"><span class="score">%s</span></td>' % raw(r.get("score")))
            cells.append('<td class="c">%s</td>' % tag(r.get("tag")))
            if i == 0 and merge_note:
                cells.append('<td rowspan="%d" style="vertical-align:middle">%s</td>' % (n, notes[0]))
            elif not merge_note:
                cells.append("<td>%s</td>" % notes[i])
            out.append("<tr>%s</tr>" % "".join(c for c in cells if c))
    html = (
        '<table><thead><tr><th style="width:110px">类别</th><th>具体项目</th>'
        '<th class="c" style="width:120px">分值</th><th class="c" style="width:80px">难度</th>'
        '<th style="width:220px">关键点 / 性价比</th></tr></thead><tbody>%s</tbody></table>' % "".join(out)
    )
    d = sec.get("danger")
    if d:
        html += box("danger", raw(d.get("title")), raw(d.get("html")), d.get("items"))
    return html


def render_roadmap(sec):
    items = []
    for it in sec.get("items", []):
        pts = "".join("<li>%s</li>" % p for p in it.get("points", []))
        sub = '<em>%s</em>' % raw(it["sub"]) if it.get("sub") else ""
        items.append(
            '<div class="tl-item"><div class="tl-h">%s%s</div><ul>%s</ul></div>'
            % (raw(it.get("head")), sub, pts)
        )
    return '<div class="timeline">%s</div>' % "".join(items)


def render_quota(sec):
    rows = []
    for r in sec.get("rows", []):
        rows.append(
            "<tr><td class=\"k\">%s</td><td>%s</td><td>%s</td></tr>"
            % (raw(r.get("type")), raw(r.get("rule")), raw(r.get("plain")))
        )
    return (
        '<table><thead><tr><th style="width:170px">名额类型</th><th>分配规则</th>'
        '<th style="width:250px">家长如何理解</th></tr></thead><tbody>%s</tbody></table>' % "".join(rows)
    )


def render_flow(sec):
    steps = []
    for s in sec.get("steps", []):
        steps.append('<div class="f"><b>%s</b><span>%s</span></div>' % (raw(s.get("t")), raw(s.get("d"))))
    return '<div class="flow">%s</div>' % "".join(steps)


def render_risks(sec):
    rows = []
    for r in sec.get("rows", []):
        rows.append(
            "<tr><td class=\"k\">%s</td><td>%s</td><td class=\"c\">%s</td></tr>"
            % (raw(r.get("stage")), raw(r.get("item")),
               tag(r.get("result"), r.get("tag_cls") or tag_class(r.get("result"))))
        )
    html = (
        '<table><thead><tr><th style="width:150px">风险阶段</th><th>情形</th>'
        '<th class="c" style="width:110px">后果</th></tr></thead><tbody>%s</tbody></table>' % "".join(rows)
    )
    w = sec.get("warn")
    if w:
        html += box("warn", raw(w.get("title")), raw(w.get("html")), w.get("items"))
    return html


def render_faq(items):
    out = []
    for it in items:
        out.append('<div class="qa"><div class="q">%s</div><div class="a">%s</div></div>'
                   % (raw(it.get("q")), raw(it.get("a"))))
    return "".join(out)


def render_todos(sec):
    rows = []
    for i, r in enumerate(sec.get("rows", []), 1):
        rows.append(
            "<tr><td class=\"c k\">%s</td><td>%s</td><td>%s</td><td>%s</td></tr>"
            % (raw(r.get("no") or i), raw(r.get("item")), raw(r.get("why")), raw(r.get("who")))
        )
    return (
        '<table><thead><tr><th class="c" style="width:50px">#</th><th>待确认事项</th>'
        '<th style="width:220px">为什么重要</th><th style="width:130px">咨询对象</th></tr></thead>'
        '<tbody>%s</tbody></table>' % "".join(rows)
    )


def render_footer(data, meta):
    lines = data.get("footer")
    if not lines:
        school = raw(meta.get("school")) or "本校"
        lines = [
            "<b>数据与口径说明</b>",
            "· 本表依据%s公开发布的推免（保研）政策文件整理。" % school,
            "· 表中“政策原文要求”“分值”“否决情形”“流程”等栏目均直接对应政策条款；"
            "“白话解读”“难度”“性价比”“战略含义”“行动路线图”为本解读版的分析与建议，不代表学校官方口径。",
            "· 涉及各学院自行制定的加分方案与认定清单，一律以孩子所在学院当年公布的推免工作办法为准；政策可能按年度调整，请以最新版本为准。",
            "· 本表仅供家庭内部规划参考，不构成任何承诺或保证。",
        ]
    return "<footer>%s</footer>" % "<br>".join(lines)


# ---------------- CSS（改样式改这里） ----------------
CSS = """
  :root{
    --red:#b01c2e; --red-d:#8a1424; --red-l:#fdeef0; --red-ll:#fff7f8;
    --gold:#a8751a; --ink:#22201f; --ink2:#57514e; --ink3:#8b847f;
    --line:#e6dcd8; --bg:#faf7f5; --green:#1f7a4d; --blue:#1c5a99;
  }
  *{box-sizing:border-box;}
  body{
    margin:0;background:var(--bg);color:var(--ink);
    font-family:"PingFang SC","Microsoft YaHei","Hiragino Sans GB",system-ui,-apple-system,sans-serif;
    font-size:15px;line-height:1.75;-webkit-font-smoothing:antialiased;
  }
  .wrap{max-width:1180px;margin:0 auto;padding:0 24px 80px;}
  header.hero{
    background:linear-gradient(135deg,var(--red-d) 0%,var(--red) 55%,#c8402f 100%);
    color:#fff;padding:44px 0 38px;margin-bottom:34px;border-bottom:5px solid #7d1120;
  }
  .hero .wrap{padding-bottom:0;}
  .kicker{
    display:inline-block;font-size:12.5px;letter-spacing:2px;
    background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.35);
    padding:4px 12px;border-radius:20px;margin-bottom:16px;
  }
  .hero h1{margin:0 0 10px;font-size:32px;line-height:1.35;font-weight:800;letter-spacing:.5px;}
  .hero .sub{font-size:16px;opacity:.94;margin-bottom:20px;font-weight:300;}
  .hero .meta{
    display:flex;flex-wrap:wrap;gap:10px 26px;font-size:13px;opacity:.9;
    border-top:1px solid rgba(255,255,255,.25);padding-top:14px;
  }
  .hero .meta b{font-weight:600;}
  h2{font-size:22px;margin:46px 0 6px;padding-left:14px;border-left:6px solid var(--red);
     line-height:1.4;font-weight:800;}
  h2 small{display:block;font-size:13px;font-weight:400;color:var(--ink3);letter-spacing:.5px;margin-top:2px;}
  h3{font-size:17px;margin:28px 0 10px;color:var(--red-d);font-weight:700;}
  p{margin:10px 0;}
  .lead{color:var(--ink2);font-size:14.5px;margin:6px 0 18px;}
  table{width:100%;border-collapse:collapse;background:#fff;font-size:14px;margin:14px 0 8px;
        box-shadow:0 1px 3px rgba(0,0,0,.06);border:1px solid var(--line);}
  thead th{background:var(--red);color:#fff;font-weight:600;text-align:left;padding:11px 13px;
           font-size:13.5px;border:1px solid #96182a;white-space:nowrap;}
  thead th.c{text-align:center;}
  tbody td{padding:11px 13px;border:1px solid var(--line);vertical-align:top;}
  tbody tr:nth-child(even){background:var(--red-ll);}
  tbody td.c{text-align:center;}
  tbody td.k{font-weight:600;color:var(--red-d);white-space:nowrap;}
  .cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:14px;margin:18px 0 10px;}
  .card{background:#fff;border:1px solid var(--line);border-top:4px solid var(--red);
        border-radius:4px;padding:16px 18px;box-shadow:0 1px 3px rgba(0,0,0,.05);}
  .card .num{font-size:30px;font-weight:800;color:var(--red);line-height:1.1;}
  .card .num span{font-size:15px;font-weight:600;}
  .card .lbl{font-size:14px;font-weight:700;margin-top:4px;}
  .card .desc{font-size:12.5px;color:var(--ink2);margin-top:5px;line-height:1.6;}
  .tag{display:inline-block;padding:1px 9px;border-radius:3px;font-size:12px;font-weight:700;
       white-space:nowrap;line-height:1.7;}
  .t-hard{background:var(--red);color:#fff;}
  .t-mid{background:#fdf0d8;color:var(--gold);border:1px solid #ecc882;}
  .t-easy{background:#e7f4ec;color:var(--green);border:1px solid #b9ddc8;}
  .t-info{background:#e9f0f8;color:var(--blue);border:1px solid #bcd2e8;}
  .t-gray{background:#efeceb;color:var(--ink2);border:1px solid #ddd7d4;}
  .score{font-weight:800;color:var(--red);font-size:15px;white-space:nowrap;}
  .formula{background:var(--red-l);border:1px dashed #e3b7bd;border-radius:5px;
           padding:18px 20px;margin:16px 0;text-align:center;}
  .formula .eq{font-size:20px;font-weight:800;color:var(--red-d);letter-spacing:1px;}
  .formula .ex{font-size:13px;color:var(--ink2);margin-top:8px;}
  .timeline{margin:18px 0 6px;position:relative;padding-left:26px;}
  .timeline::before{content:"";position:absolute;left:8px;top:6px;bottom:6px;width:2px;background:#e3c3c7;}
  .tl-item{position:relative;margin-bottom:18px;background:#fff;border:1px solid var(--line);
           border-radius:4px;padding:14px 18px;}
  .tl-item::before{content:"";position:absolute;left:-22px;top:20px;width:11px;height:11px;border-radius:50%;
                   background:var(--red);border:2px solid #fff;box-shadow:0 0 0 2px var(--red);}
  .tl-item .tl-h{font-size:15px;font-weight:800;color:var(--red-d);margin-bottom:6px;}
  .tl-item .tl-h em{font-style:normal;font-size:12.5px;color:var(--ink3);font-weight:400;margin-left:8px;}
  .tl-item ul{margin:6px 0 0;padding-left:20px;}
  .tl-item li{margin:3px 0;font-size:13.5px;color:var(--ink2);}
  .box{border-radius:4px;padding:14px 18px;margin:16px 0;font-size:14px;}
  .box .bt{font-weight:800;margin-bottom:6px;display:block;}
  .box ul{margin:6px 0 0;padding-left:20px;}
  .box li{margin:4px 0;}
  .box-warn{background:#fff6e8;border:1px solid #f0d49d;border-left:5px solid var(--gold);}
  .box-warn .bt{color:var(--gold);}
  .box-danger{background:var(--red-l);border:1px solid #e9bcc1;border-left:5px solid var(--red);}
  .box-danger .bt{color:var(--red-d);}
  .box-tip{background:#eaf4ee;border:1px solid #bcdcc9;border-left:5px solid var(--green);}
  .box-tip .bt{color:var(--green);}
  .flow{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:12px;margin:16px 0;}
  .flow .f{background:#fff;border:1px solid var(--line);border-left:4px solid var(--red);
           padding:12px 14px;border-radius:3px;font-size:13.5px;}
  .flow .f b{display:block;color:var(--red-d);font-size:13px;margin-bottom:3px;}
  .flow .f span{color:var(--ink2);}
  .qa{background:#fff;border:1px solid var(--line);border-radius:4px;margin:12px 0;overflow:hidden;}
  .qa .q{background:var(--red-ll);padding:12px 16px;font-weight:700;color:var(--red-d);
         font-size:14.5px;border-bottom:1px solid var(--line);}
  .qa .a{padding:12px 16px;font-size:13.8px;color:var(--ink2);}
  footer{margin-top:56px;padding-top:20px;border-top:2px solid var(--line);
         font-size:12.5px;color:var(--ink3);line-height:1.8;}
  @media print{
    body{background:#fff;font-size:12px;}
    header.hero{background:var(--red) !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    thead th{background:var(--red) !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
    tbody tr:nth-child(even){background:#faf7f5 !important;}
    h2{page-break-after:avoid;} tr{page-break-inside:avoid;}
  }
  @media (max-width:720px){
    .hero h1{font-size:23px;} table{font-size:12.5px;}
    tbody td,thead th{padding:8px 9px;} .wrap{padding:0 14px 50px;}
  }
"""


# ---------------- 主流程 ----------------
def build(data, title_override=None):
    meta = data.get("meta") or {}
    warn = []

    parts = ['<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">'
             '<meta name="viewport" content="width=device-width, initial-scale=1.0">']
    page_title = title_override or raw(meta.get("title")) or \
        ("%s推免政策 · 家长学生解读版" % raw(meta.get("school")) if meta.get("school") else "推免政策解读")
    parts.append("<title>%s</title><style>%s</style></head><body>" % (page_title, CSS))
    parts.append(render_hero(meta))
    parts.append('<div class="wrap">')

    idx = 0
    rendered = []

    def emit(title_key, default_title, body_html, note="", lead=""):
        nonlocal idx
        if lead:
            body_html = '<p class="lead">%s</p>' % lead + body_html
        rendered.append((title_key, h2(idx, default_title, note) + body_html))
        idx += 1

    # 1 速览
    if data.get("highlights"):
        body = render_highlights(data["highlights"])
        if data.get("summary"):
            s = data["summary"]
            body += box(s.get("type", "tip"), raw(s.get("title")), raw(s.get("html")), s.get("items"))
        emit("highlights",
             data.get("highlights_title") or DEFAULT_TITLES["highlights"] + "：四个数字看懂保研",
             body, data.get("highlights_note", "如果只有 3 分钟，看完这一屏就够了"))
    else:
        warn.append("缺 highlights（顶部数字卡）——建议写 4 个关键数字")

    # 2 门槛
    sec = data.get("thresholds")
    if sec and sec.get("rows"):
        emit("thresholds", sec.get("title", DEFAULT_TITLES["thresholds"]),
             render_thresholds(sec), sec.get("note", ""), sec.get("lead", ""))
    else:
        warn.append("缺 thresholds（资格门槛表）——政策解读的核心，不能没有")

    # 3 一票否决
    sec = data.get("vetoes")
    if sec and sec.get("rows"):
        body = ""
        if sec.get("h3"):
            body += "<h3>%s</h3>" % raw(sec["h3"])
        body += render_vetoes(sec)
        emit("vetoes", sec.get("title", DEFAULT_TITLES["vetoes"]), body, sec.get("note", ""))
    else:
        warn.append("缺 vetoes（一票否决）——政策里的'首次修读不合格'等高杀伤条款必须单列")

    # 4 综合成绩
    sec = data.get("formula")
    if sec:
        emit("formula", sec.get("title", DEFAULT_TITLES["formula"] + "：主战场在哪里？"),
             render_formula(sec), sec.get("note", ""))
    else:
        warn.append("缺 formula（综合成绩算法）——建议写清 A+B 口径")

    # 5 奖励分
    sec = data.get("bonus")
    if sec and sec.get("rows"):
        emit("bonus", sec.get("title", DEFAULT_TITLES["bonus"]),
             render_bonus(sec), sec.get("note", ""), sec.get("lead", ""))
    else:
        warn.append("缺 bonus（奖励分全景表）——家长最关心的一张表")

    # 6 路线图
    sec = data.get("roadmap")
    if sec and sec.get("items"):
        emit("roadmap", sec.get("title", DEFAULT_TITLES["roadmap"] + "：大学三年该怎么排？"),
             render_roadmap(sec), sec.get("note", "基于政策倒推的建议时间表（非政策原文）"))
    else:
        warn.append("缺 roadmap（行动路线图）")

    # 7 名额
    sec = data.get("quota")
    if sec and sec.get("rows"):
        emit("quota", sec.get("title", DEFAULT_TITLES["quota"]), render_quota(sec), sec.get("note", ""))

    # 8 流程
    sec = data.get("flow")
    if sec and sec.get("steps"):
        emit("flow", sec.get("title", DEFAULT_TITLES["flow"]), render_flow(sec), sec.get("note", ""))

    # 9 红线
    sec = data.get("risks")
    if sec and sec.get("rows"):
        emit("risks", sec.get("title", DEFAULT_TITLES["risks"]), render_risks(sec), sec.get("note", ""))
    else:
        warn.append("缺 risks（红线与风险清单）")

    # 10 FAQ
    if data.get("faq"):
        n = len(data["faq"])
        emit("faq", data.get("faq_title") or ("家长最常问的 %d 个问题" % n), render_faq(data["faq"]))

    # 11 待确认
    sec = data.get("todos")
    if sec and sec.get("rows"):
        n = len(sec["rows"])
        emit("todos", sec.get("title") or ("需要向学院确认的 %d 个问题" % n),
             render_todos(sec), sec.get("note", "本表无法回答的部分 · 建议家长陪同孩子一次性问清"))
    else:
        warn.append("缺 todos（待确认清单）——学院自定规则必须落到这里")

    parts.extend(h for _, h in rendered)
    parts.append(render_footer(data, meta))
    parts.append("</div></body></html>")
    return "".join(parts), [k for k, _ in rendered], warn


def main():
    ap = argparse.ArgumentParser(description="推免政策解读 HTML 生成器")
    ap.add_argument("json")
    ap.add_argument("output", nargs="?", default=None)
    ap.add_argument("--title", default=None, help="覆盖 <title> 与 H1")
    args = ap.parse_args()

    if not os.path.isfile(args.json):
        sys.exit("找不到数据文件：%s" % args.json)
    with open(args.json, "r", encoding="utf-8") as f:
        data = json.load(f)

    html, keys, warn = build(data, args.title)

    if args.output:
        out = args.output
    else:
        school = raw((data.get("meta") or {}).get("school")) or "政策"
        out = "%s推免政策_家长学生解读分析表.html" % school
    if not out.lower().endswith(".html"):
        out += ".html"

    with open(out, "w", encoding="utf-8") as f:
        f.write(html)

    print("模块（%d）：%s" % (len(keys), " / ".join(keys)))
    print("输出：%s" % os.path.abspath(out))
    print("字符数：%d ｜ 表格数：%d" % (len(html), html.count("<table>")))
    if warn:
        print("⚠ 提示：")
        for w in warn:
            print("   - " + w)
    else:
        print("✅ 11 个模块齐全")
    print("交付前自查：政策原文类栏目是否逐条对应条款？页脚是否声明了'解读意见非官方口径'？")


if __name__ == "__main__":
    main()

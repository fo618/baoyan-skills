# -*- coding: utf-8 -*-
"""
从**全国推荐免试攻读研究生（免初试、转段）信息公开管理服务系统**「推荐办法」栏目
（https://yz.chsi.com.cn/tm/tjbf/index）抓取指定高校的**校级 + 学院级**推免方案，
**默认生成自包含 HTML 交付件**落到桌面。

这是学员政策的**一手源头**：校级「XX大学2027届本科毕业生推荐免试攻读研究生工作方案」
+ 学员所在学院的「XX学院2027届…工作方案」（含名额、绩点线、英语线、加分细则表）。
配合 fetch_yanzhaowang_policy.py（全国规定）一起用，政策链条才完整。

交付形态（2026-09-28 用户定版）：**只交付 HTML**。院校政策原文是富文本（含加分细则表格），
原样嵌进 HTML 保真度最高；自研重排 PDF 会破坏表格与缩进，用户明确反馈"排版太乱"。
加 `--pdf` 才会额外生成重排 PDF（仅作备选，不作为交付物）。用户自己浏览器 Ctrl+P 转存即可。

用法：
    python fetch_school_tuimian_policy.py --school "中国地质大学" --college "艺术与传媒学院"
    python fetch_school_tuimian_policy.py --school "中国地质大学" --list          # 只看有哪些院系发布了
    python fetch_school_tuimian_policy.py --school "中国地质大学" --college "艺术与传媒" --dir D:\\交付
    python fetch_school_tuimian_policy.py --school "中国地质大学" --dry-run
    python fetch_school_tuimian_policy.py --school "中国地质大学" --pdf           # 额外出重排 PDF（不推荐）

实现要点（踩坑沉淀，改脚本前先读）：
1. **必须用 curl 发请求**：yz.chsi.com.cn/tm 下有 TLS/UA 指纹校验，Python urllib 拿不到
   detailAll 的首页 taskId（实测 4 次全空，换完整请求头也没用），而 curl 一次就成。
   所以本脚本所有 HTTP 都走 subprocess 调用 curl（-sL -k + cookie jar），Python 只做解析。
2. **detailAll 是异步任务**：GET detailAll?schCode=X → 页面里拿到 taskId →
   GET /tm/asynProgress.do?taskId=X 轮询（返回 {"state":"SUCCESS"}）→
   POST /tm/tjbf/detailAll（taskId + schCode）→ 拿到含 schList / yxsList 的结果页。
3. 校级清单在 `schList:`，**院系清单在 `yxsList:`**（JSON 内联在 Vue data 里，形如 `xxx: [...] || []`）。
   正文在 detail 页的 `content: "..."`（富文本 HTML，含表格）。
4. 官方**只给 HTML 富文本，没有 PDF 附件**，所以交付件原样嵌入该富文本（表格完整保留）。
5. **多数高校的学院方案就是转发校级办法 + 同一张加分表**，两份正文常只差标题行——这是真实情况不是 bug。
"""

import os
import re
import json
import html as htmlmod
import argparse
import subprocess
import tempfile
import datetime
import sys

BASE = "https://yz.chsi.com.cn"
TM = BASE + "/tm"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"

# 复用全国政策脚本里的 PDF 排版器（同目录）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from fetch_yanzhaowang_policy import build_pdf, desk_dir, safe_name, write_html
except Exception:  # 单独运行时兜底
    build_pdf = None
    write_html = None
    def desk_dir():
        for k in ("USERPROFILE", "HOME"):
            b = os.environ.get(k)
            if not b:
                continue
            for c in (os.path.join(b, "Desktop"), os.path.join(b, "OneDrive", "Desktop")):
                if os.path.isdir(c):
                    return c
        return os.getcwd()
    def safe_name(s):
        return re.sub(r'[\\/:*?"<>|]', "_", s).strip()


# --------------------------------------------------------------------------- HTTP（走 curl）
class Curl:
    def __init__(self):
        self.jar = os.path.join(tempfile.gettempdir(), "yz_tm_cookies.txt")
        self.exe = self._which()

    @staticmethod
    def _which():
        for c in ("curl", "curl.exe"):
            try:
                subprocess.run([c, "--version"], stdout=subprocess.DEVNULL,
                               stderr=subprocess.DEVNULL, timeout=15)
                return c
            except Exception:
                continue
        return None

    def run(self, args, timeout=60):
        if not self.exe:
            raise RuntimeError("未找到 curl，无法访问研招网推免系统（该站点对 urllib 有指纹拦截）")
        cmd = [self.exe, "-sL", "-k", "-c", self.jar, "-b", self.jar, "-A", UA]
        cmd += args
        p = subprocess.run(cmd, capture_output=True, timeout=timeout)
        out = p.stdout
        for enc in ("utf-8", "gb18030"):
            try:
                return out.decode(enc)
            except UnicodeDecodeError:
                continue
        return out.decode("utf-8", "ignore")

    def get(self, url, timeout=60):
        return self.run([url], timeout=timeout)

    def post(self, url, fields, timeout=60):
        args = [url]
        for k, v in fields.items():
            args += ["--data-urlencode", f"{k}={v}"]
        return self.run(args, timeout=timeout)


# --------------------------------------------------------------------------- 步骤
def search_school(curl, name):
    """POST /tm/tjbf/queryAll 按校名检索，返回候选 [(dwdm, dwmc, existZsgs)]"""
    r = curl.post(TM + "/tjbf/queryAll",
                  {"dwmc": name, "ssdm": "", "curPage": 1, "pageCount": 20,
                   "start": 0, "totalCount": 0})
    try:
        j = json.loads(r)
    except Exception:
        m = re.search(r'"list"\s*:\s*(\[.*?\])', r, re.S)
        if not m:
            return []
        j = {"msg": {"list": json.loads(m.group(1))}}
    lst = (j.get("msg") or {}).get("list") or []
    return [(it.get("dwdm"), it.get("dwmc"), it.get("existZsgs"))
            for it in lst if it.get("dwdm")]


def fetch_policies(curl, sch_code):
    """三步异步：拿 taskId → 轮询 → POST 取结果页；返回 (schList, yxsList, schName)"""
    page = curl.get(f"{TM}/tjbf/detailAll?schCode={sch_code}")
    m = re.search(r'taskId:\s*"([^"]+)"', page)
    if not m:
        raise RuntimeError("未能从 detailAll 页面取到 taskId（站点可能改版或限流，稍后重试）")
    task = m.group(1)

    ok = False
    for _ in range(12):
        r = curl.get(f"{TM}/asynProgress.do?taskId={task}", timeout=30)
        if '"SUCCESS"' in r or '"flag":true' in r.replace(" ", "") or '"flag": true' in r:
            ok = True
            break
        try:
            import time
            time.sleep(1.5)
        except Exception:
            pass
    if not ok:
        raise RuntimeError(f"异步任务未完成：taskId={task}")

    res = curl.post(f"{TM}/tjbf/detailAll", {"taskId": task, "schCode": sch_code}, timeout=90)

    def arr(key):
        m = re.search(re.escape(key) + r'\s*:\s*(\[.*?\])\s*\|\|', res, re.S)
        if not m:
            m = re.search(re.escape(key) + r'\s*:\s*(\[.*?\])', res, re.S)
        if not m:
            return []
        try:
            return json.loads(m.group(1))
        except Exception:
            return []

    nm = re.search(r'schName:\s*"([^"]*)"', res)
    sch_name = ""
    if nm:
        sch_name = nm.group(1).encode().decode("unicode_escape")
    return arr("schList"), arr("yxsList"), sch_name


def nice_name(unit, title):
    """把冗长的公示标题收成一个干净的文件名。
    例：《XX大学学校办公室 关于印发《2027届本科毕业生推荐免试 攻读研究生工作方案》的通知》
      → '中国地质大学(武汉)2027届本科毕业生推荐免试攻读研究生工作方案'"""
    m = re.search(r'《([^》]+)》', title)
    core = m.group(1) if m else title
    core = re.sub(r'^\s*关于印发\s*', '', core)
    core = re.sub(r'\s*的通知\s*$', '', core)
    core = re.sub(r'\s+', '', core)          # 标题里的空格/换行都是排版残留，去掉
    core = core.strip('《》。 ')
    if not core:
        core = re.sub(r'\s+', '', title)[:40]
    unit_c = re.sub(r'\s+', '', unit or '')
    # 只有单位名「完全没出现在标题里」才补前缀，否则会出现
    # 「仪器与电子学院中北大学仪器与电子学院…」这种叠名（2026-09-29 中北大学实测）
    if unit_c and unit_c not in core:
        core = unit_c + core
    return core[:70]


def html_to_lines(content_html):
    """富文本 → 纯文本行；表格按 '列 | 列' 铺开，保持可读。"""
    s = content_html
    s = re.sub(r"(?is)<br\s*/?>", "\n", s)
    s = re.sub(r"(?is)</t[dh]>", " | ", s)
    s = re.sub(r"(?is)</tr>", "\n", s)
    s = re.sub(r"(?is)</p>", "\n", s)
    s = re.sub(r"(?is)</(div|h[1-6]|li)>", "\n", s)
    s = re.sub(r"(?s)<[^>]+>", "", s)
    s = htmlmod.unescape(s)
    s = s.replace("\u00a0", " ").replace("\u3000", " ")
    out = []
    for line in s.split("\n"):
        line = re.sub(r"[ \t]+", " ", line).strip()
        line = re.sub(r"(\s*\|\s*)+", " | ", line).strip(" |").strip()
        if line:
            out.append(line)
    return out


def fetch_detail(curl, zsgs_id, sch_code):
    """取单份办法的正文，返回 (title, lines, deployTime, content_html)
    content_html 是原文富文本（含表格/加粗/缩进），用于生成高保真 HTML 交付件。"""
    page = curl.get(f"{TM}/tjbf/detail?zsgsId={zsgs_id}&schCode={sch_code}", timeout=60)
    title, time_ = "", ""
    mt = re.search(r'title\s*:\s*"((?:[^"\\]|\\.)*)"', page)
    if mt:
        try:
            title = json.loads('"' + mt.group(1) + '"')
        except Exception:
            title = mt.group(1)
    md = re.search(r'deployTime\s*:\s*"([^"]*)"', page)
    if md:
        time_ = md.group(1)

    body = ""
    mc = re.search(r'\bcontent\s*:\s*"((?:[^"\\]|\\.)*)"', page, re.S)
    if mc:
        try:
            body = json.loads('"' + mc.group(1) + '"')
        except Exception:
            body = mc.group(1).replace('\\"', '"').replace("\\/", "/")
    # 正文里的图片若用相对路径，补成绝对地址，保证单机打开也能看到（签章/扫描件常见）
    if body:
        body = re.sub(r'(<img[^>]+src=")(/[^"]+")', lambda m: m.group(1) + BASE + m.group(2), body)
    return title, html_to_lines(body), time_, body


# --------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="从研招网推免系统抓取指定高校（校级+学院级）推免方案 PDF")
    ap.add_argument("--school", required=True, help="本科院校名称关键字，如 中国地质大学")
    ap.add_argument("--college", default="", help="学院名称关键字，如 艺术与传媒（不给则只下校级并列院系）")
    ap.add_argument("--pick", default="", help="校名消歧关键字，如 武汉 / 北京")
    ap.add_argument("--dir", help="输出目录（默认桌面）")
    ap.add_argument("--list", action="store_true", help="只列出已发布的院系清单，不下载")
    ap.add_argument("--pdf", action="store_true",
                    help="额外再生成一份重排 PDF（默认只交付 HTML，重排 PDF 排版不如原文保真）")
    ap.add_argument("--dry-run", action="store_true", help="只预览")
    a = ap.parse_args()

    curl = Curl()
    cands = search_school(curl, a.school)
    if not cands:
        print(f"❌ 推免系统里没检索到「{a.school}」，换个关键字试试（如去掉「大学」二字）")
        sys.exit(1)

    pick = cands
    if a.pick:
        pick = [c for c in cands if a.pick in (c[1] or "")] or cands
    if len(pick) > 1:
        print("⚠ 检索到多所同名院校：")
        for dwdm, dwmc, ex in pick:
            print(f"   {dwdm}  {dwmc}  {'已发布' if ex else '未发布'}")
        print(f"   默认取第一所：{pick[0][1]}（用 --pick 指定）")
    dwdm, dwmc, exist = pick[0]
    if not exist:
        print(f"❌ {dwmc} 在推免系统里尚未发布推荐办法")
        sys.exit(2)
    print(f"目标院校  : {dwmc}（代码 {dwdm}）")

    sch_list, yxs_list, sch_name = fetch_policies(curl, dwdm)
    sch_name = sch_name or dwmc
    print(f"校级办法  : {len(sch_list)} 条 | 院系办法: {len(yxs_list)} 条（其中已发布 "
          f"{len([x for x in yxs_list if x.get('status') == 'PUBLISH'])} 条）")

    if a.list or not a.college:
        print("\n已发布院系：")
        for x in yxs_list:
            st = x.get("status")
            print(f"   [{'√' if st == 'PUBLISH' else '×'}] {x.get('dicName')}"
                  f"  {(x.get('title') or '（参照校级办法）')[:40]}  {x.get('deployTime') or ''}")
        if not a.college:
            print("\n（加 --college <学院关键字> 下载对应学院的方案 PDF）")

    targets = []
    for x in sch_list:
        if x.get("id"):
            targets.append(("校级", x.get("dicName") or sch_name, x["id"], x.get("title") or "", x.get("deployTime") or ""))
    if a.college:
        hit = [x for x in yxs_list
               if x.get("id") and a.college in (x.get("dicName") or "")]
        if not hit:
            print(f"⚠ 未匹配到学院「{a.college}」，可先用 --list 看准确院名")
        for x in hit:
            targets.append(("院系", x.get("dicName"), x["id"], x.get("title") or "", x.get("deployTime") or ""))

    if not targets:
        print("❌ 没有可下载的目标（未发布或无匹配）")
        sys.exit(3)

    outdir = a.dir or desk_dir()
    os.makedirs(outdir, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    for scope, unit, zid, title, dtime in targets:
        t, lines, dt, body_html = fetch_detail(curl, zid, dwdm)
        t = t or title or f"{unit}推免工作方案"
        n = len(lines)
        print(f"\n【{scope}】{t}")
        print(f"   正文 {n} 行，发布日 {dt or dtime or '—'}")
        if a.dry_run:
            print("   ---- 前 300 字 ----")
            print("   " + " ".join(lines)[:300].replace("\n", " "))
            continue
        fname = nice_name(unit, t)
        meta = "　|　".join(x for x in (unit, f"发布日 {dt or dtime}" if (dt or dtime) else "",
                                       "研招网推免信息公开系统公示") if x)
        src = f"{TM}/tjbf/detail?zsgsId={zid}&schCode={dwdm}"
        if write_html is not None:
            dst = os.path.join(outdir, safe_name(f"《{fname}》.html"))
            size = write_html(dst, fname, meta, src, body_html or "<p>（未取到正文）</p>")
            print(f"   ✅ 已生成 HTML：{dst}（{size/1024:.0f} KB）")
        if a.pdf and build_pdf is not None:
            dst = os.path.join(outdir, safe_name(f"《{fname}》.pdf"))
            build_pdf({"title": fname, "docno": meta, "year": str(datetime.datetime.now().year),
                       "paras": [f"【{scope}】{unit}"]
                       + ([f"发布日期：{dt or dtime}"] if (dt or dtime) else []) + lines},
                      src, dst)
            print(f"   ✅ 已生成重排 PDF：{dst}（排版不如 HTML 保真，仅作备选）")

    if not a.dry_run:
        print(f"\n完成。输出目录：{outdir}")
        print("口径说明：交付的是依研招网原文生成的 HTML（保留原表格与原排版），"
              "浏览器打开 Ctrl+P 另存为 PDF 即为最接近原件的版本。")


if __name__ == "__main__":
    main()

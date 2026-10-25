# -*- coding: utf-8 -*-
"""
保研规划表 → PPT 投放器（原生表格版）

把两张规划表（学员信息表 / 冲稳保定档表）
以 **原生 PowerPoint 表格** 写进用户的 PPT，而不是贴 Excel 截图。
（2026-09-17 起由"四表"改为"三表"；**2026-09-19 二次定版：PPT 不再投放
「主线任务表」**，底层模板第 4 页已删，内置模板 26 页。main_timeline 数据
仍留在 JSON / Excel 里，SPECS 里该表的 spec 以注释形式保留，需要时恢复。）

三个设计目标：
  1. 通用：任何 PPT 都能跑。自动识别页面尺寸（16:9 / 4:3 / 任意自定义比例），
     表格尺寸、字号、行高全部按版心精算，交付即成型，不需要手动拖。
  2. 就地：优先找到 PPT 里对应的三个小节页（按标题关键词匹配），
     把该页原来的表格截图（图片）替换成可编辑表格，页眉、左侧导航、背景全部原样保留。
     找不到对应页时，自动克隆一页同版式页面再放表。
     表格位置用 2026-10-12 定稿的「标准版心」（左避导航、下留 0.56in，不跟随占位截图位置）。
  3. 精简：PPT 用短文案（JSON 里给 ppt_tables 覆盖，或由脚本从完整版自动压缩），
     字号自动下调到刚好装下并竖向拉满版心。

用法:
  python build_ppt_slides.py <data.json> [source.pptx] [output.pptx] [选项]

模板 PPT（第 2 个位置参数）可以省略：省略时会在 <skill>/assets/templates/ 下
自动找唯一的 .pptx。把自己的模板丢进那个目录即可，不用改命令。

输出位置（2026-09-21 用户定版：交付物直接落在桌面）:
  没给 [output.pptx]      → 桌面/保研规划_<院校>.pptx（院校名取自 cover.school → meta.school
                            → student_data_<后缀>.json 的后缀）
  给了个文件名            → 桌面/<该文件名>
  给了绝对/带目录的路径    → 原样写到你指定的位置（同事自定义场景）
  --dir <路径>            → 强制指定输出目录（优先级最高）
  --no-desktop            → 关掉桌面默认，退回"当前工作目录"

选项:
  --slides 2,3          手动指定两张表落在第几页（顺序：学员信息/院校定档）
  --mode append         追加新页到末尾（默认 replace=就地替换原有图片）
  --max-font 12         字号搜索上限（pt，以 13.33x7.5in 为基准；PPT 演示最佳区间 10~14）
  --min-font 7          字号搜索下限
  --theme auto|红|HEX,HEX,HEX   主色。默认沿用 skill 的深红主题；
                                 auto=从 PPT 提取强调色；也可给 主色,浅色,描边 三个十六进制
  --no-fill-height      不让表格竖向拉满版心（默认拉满，视觉更整齐）
  --no-cover            不改写第 1 页封面（默认会把封面院校/专业换成本次学员信息）
  --no-tier-check       跳过院校层次铁律自检（默认开启，不通过则中止）
  --force               自检不通过也继续生成
  --keep-empty-pages    保留"没投表格的空内容页"（默认会删掉；2.0 模板三个内容页刚好用满，
                         不会删页；用旧版模板时多余的内容页会被删掉）
  --ppt-verbose         回到"全量"字数上限（默认是精简版，字数上限见下）
  --dry-run             只打印方案，不写文件

data.json:
  数据源（student_info / school_tiers / main_timeline），字段同 generate_tables.py。
  PPT 只渲染 student_info（自动过滤「院校定位」行）与 school_tiers 两张表；
  main_timeline 已下线（2026-09-19），结构保留兼容但不渲染。
  PPT 精简文案优先取 JSON 中的 "ppt_tables" **覆盖块**——只需写与完整版不同的字段，
  脚本按"逐字段覆盖"合并（例如只覆盖 school_tiers[i].school）。

  可选 "cover" 块 —— 用于改写 PPT 第 1 页封面（每回自动换成当前学员信息）：
    "cover": { "school": "中国石油大学（北京）", "major": "油气储运工程" }
  没写 "cover" 时，脚本会从 meta.school / meta.major 自动推断。
  想让院校名用简称（如"中国石油大学"），直接在 cover.school 里写你要显示的字即可。

PPT 文案字数上限（单位：该单元格总字数，默认精简版；超过只会告警不会报错）：
  学员信息表（4 行）表头摘要 40 / 内容 160
  冲稳保定档表（4 列）适配院校 12 / 方向 20 / 优势 24
    （2026-09-19 起删掉「录取大致所需条件」列，requirement 只进 Excel 不进 PPT）
  主线任务表 **已全面下线**（2026-09-19 三次定版：PPT / Excel 均不生成）。
用 --ppt-verbose 可切回放宽版（60/320、20/26/40）。
"""
import copy
import glob
import json
import os
import re
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

# ---------------- 主题常量（与 Excel 版保持一致） ----------------
HEADER_BG = "9E2F2F"
HEADER_FG = "FFFFFF"
FIRSTCOL_BG = "F7E4E1"
FIRSTCOL_FG = "7A2222"
BODY_BG = "FFFFFF"
BODY_FG = "333333"
BORDER_COLOR = "C97B7B"
FONT_NAME = "微软雅黑"

EMU_PER_IN = 914400
PX_PER_IN = 96.0

# 排版估算参数
LINE_FACTOR = 1.24        # 行距系数
CELL_PAD = 2.4            # 单元格上下留白（px，按 12px 字号标定）
HEADER_BUMP = 1.1         # 表头行高系数
WIDTH_SAFETY = 1.06       # 宽度估算安全系数
HEIGHT_SAFETY = 1.03
CJK_RATIO = 1.0
ASCII_RATIO = 0.53
MAX_LINES_PER_CELL = 12
CELL_H_PAD = 8.0          # 单元格左右可用宽度扣减（px）

# ================ 标准表格版心（2026-10-12 用户定版：以后出 PPT 一律按此排版）================
# 山东大学实例定稿：两张表统一从 (2.95, 1.22)in 起、占 10.05 x 5.72in（16:9 页面
# 13.33 x 7.5in 下）——左缘避开左侧导航（2.52in）留 0.43in，顶部避开页标题，
# 右侧距页边 0.33in、底部距页边 0.56in。改排版只改下面几个常量。
# 注意：不再跟随源 PPT 里占位截图的位置——截图位置不可控（山东大学的占位图在
# 2.27in，表格曾被直接压上左侧导航栏），而各院校模板内容区结构一致，标准版心通用。
STD_GAP_AFTER_NAV = 0.43  # 表格左缘距导航右缘（in）
STD_TOP_RATIO = 0.1627    # 表格顶部 = 页高 x 0.1627（16:9 即 1.22in，避开页标题）
STD_RIGHT_MARGIN = 0.33   # 表格右缘距页右（in）
STD_BOTTOM_MARGIN = 0.56  # 表格底缘距页底（in）
STD_LEFT_RATIO = 0.2213   # 无导航时的左缘 = 页宽 x 0.2213（16:9 即 2.95in）

SPECS = [
    {
        "key": "student_info",
        "title": "学员信息",
        "sub": "学员信息表",
        "match": ["学员信息", "学员情况", "学员背景", "个人信息"],
        # 表头第 2 格在运行期替换为 student_info.headline
        "columns": ["学员情况", "内容"],
        "header_from": "headline",
        "weights": [0.11, 0.89],
        "center_cols": [0],
        "min_px": [150, 0],
        "max_px": [240, 0],
        "caps": {"headline": 40, "content": 160},
        "caps_full": {"headline": 60, "content": 320},
    },
    {
        "key": "school_tiers",
        "title": "院校定档",
        "sub": "冲稳保定档院校适配表",
        "match": ["院校定档", "院校分析", "定档", "冲稳保", "目标院校"],
        # 2026-09-19 用户定版：PPT 定档表删掉「录取大致所需条件」列，固定 4 列。
        # requirement 数据仍保留在 JSON / Excel 版里，只是 PPT 不渲染。
        "columns": ["梯度", "适配院校", "核心适配方向", "匹配你的核心优势"],
        "weights": [0.085, 0.30, 0.30, 0.315],
        "center_cols": [0],
        "merge_first": True,
        "min_px": [80, 250, 230, 250],
        "max_px": [115, 360, 340, 420],
        "caps": {"school": 12, "direction": 20, "fit": 24},
        "caps_full": {"school": 20, "direction": 26, "fit": 40},
    },
    # 2026-09-19 二次定版：**PPT 不再出「主线任务表」**（底层模板第 4 页已删，26 页）。
    # main_timeline 数据仍保留在 JSON / Excel 版里，只是 PPT 不渲染；
    # 若哪天要放回去，把下面这段 spec 恢复即可（FIELD_KEYS 里也还留着字段顺序）。
    # {
    #     "key": "main_timeline",
    #     "title": "主线任务",
    #     "sub": "保研主线任务表",
    #     "match": ["主线任务", "主线", "任务清单"],
    #     "columns": ["时间", "成绩", "英语", "竞赛", "科研 & 项目", "其他"],
    #     "weights": [0.125, 0.17, 0.17, 0.165, 0.18, 0.19],
    #     "center_cols": [0],
    #     "vmerge": {"col": 1, "text": "排名越靠前越好"},
    #     "min_px": [88, 150, 165, 155, 165, 165],
    #     "max_px": [130, 330, 260, 250, 270, 270],
    #     "caps": {"time": 12, "score": 45, "english": 32,
    #              "competition": 30, "research": 30, "other": 30},
    #     "caps_full": {"time": 20, "score": 70, "english": 50,
    #                   "competition": 46, "research": 46, "other": 46},
    # },
]

# 各表在 JSON 里的字段顺序（与 SPECS[*]["columns"] 一一对应）
FIELD_KEYS = {
    "student_info": ["label", "content"],
    "school_tiers": ["tier", "school", "direction", "fit"],
    "main_timeline": ["time", "score", "english", "competition", "research", "other"],
}

# 是否使用全量版字数上限（默认 False＝用精简后的 PPT 标准）
USE_VERBOSE_CAPS = False


def _caps(spec):
    return spec["caps_full"] if USE_VERBOSE_CAPS else spec["caps"]


# ---------------- 基础工具 ----------------
def emu2px(v):
    return (v or 0) / EMU_PER_IN * PX_PER_IN


def px2emu(v):
    return int(round(v / PX_PER_IN * EMU_PER_IN))


def text_lines(text):
    return [ln for ln in str(text or "").split("\n")]


def line_units(line):
    return sum(CJK_RATIO if ord(ch) > 127 else ASCII_RATIO for ch in line)


def max_line_units(text):
    return max([line_units(ln) for ln in text_lines(text)] or [0.0])


def wrap_count(line, avail_units):
    if avail_units <= 0:
        return 1
    total = line_units(line) * WIDTH_SAFETY
    n = int(total / avail_units)
    if total % avail_units > 1e-3:
        n += 1
    return max(1, n)


def mix(hex_a, hex_b, t=0.5):
    a = RGBColor.from_string(hex_a.lstrip("#"))
    b = RGBColor.from_string(hex_b.lstrip("#"))
    out = []
    for x, y in zip(a, b):
        out.append(int(round(x + (y - x) * t)))
    return "%02X%02X%02X" % tuple(out)


def set_font(run, name, size_pt, bold=False, color_hex=None):
    f = run.font
    if size_pt:
        f.size = Pt(size_pt)
    f.bold = bold
    if color_hex:
        f.color.rgb = RGBColor.from_string(color_hex)
    f.name = name
    try:
        rPr = run._r.get_or_add_rPr()
        latin = rPr.find(qn("a:latin"))
        ea = rPr.find(qn("a:ea"))
        if ea is None:
            ea = etree.Element(qn("a:ea"))
            if latin is not None:
                latin.addnext(ea)
            else:
                rPr.append(ea)
        ea.set("typeface", name)
    except Exception:
        pass


# ---------------- 主题识别 ----------------
def _sat_lum(rgb_str):
    r, g, b = (int(rgb_str[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    mx, mn = max(r, g, b), min(r, g, b)
    lum = (mx + mn) / 2
    sat = 0 if mx == mn else (mx - mn) / (1 - abs(2 * lum - 1)) if abs(2 * lum - 1) != 1 else 0
    return sat, lum


def extract_accent(prs, limit=8):
    """从 PPT 抽一个最常用的强调色（跳过白/黑/灰）。"""
    from collections import Counter
    cnt = Counter()
    for slide in list(prs.slides)[:limit]:
        for shp in slide.shapes:
            try:
                if shp.shape_type == 13 or str(shp.shape_type).startswith("PICTURE"):
                    continue
                if shp.fill.type is not None and str(shp.fill.type) == "MSO_FILL_TYPE.SOLID (1)":
                    cnt[str(shp.fill.fore_color.rgb)] += 1
            except Exception:
                pass
            if shp.has_text_frame:
                for p in shp.text_frame.paragraphs:
                    for r in p.runs:
                        try:
                            if r.font.color and r.font.color.type is not None:
                                cnt["F" + str(r.font.color.rgb)] += 1
                        except Exception:
                            pass
    best, best_n = None, 0
    for rgb, n in cnt.most_common():
        key = rgb[1:] if rgb.startswith("F") else rgb
        sat, lum = _sat_lum(key)
        if sat < 0.25 or lum < 0.18 or lum > 0.72:
            continue
        if n > best_n:
            best, best_n = key, n
    return best


# ---------------- 页面/形状工具 ----------------
def slide_text(slide):
    return "\n".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)


def content_hole_slides(prs):
    """识别"空内容页"——只剩左侧导航 + 一个页标题、内容区什么都没有的页。

    判定从严（宁可少认，也不要误删有内容的页）：必须同时满足
      ① 没有任何表格；② 没有任何图片；③ 内容区（left >= 15% 页宽）里
      只有恰好 1 个短文本框（就是那个页标题，模板里在 21% 处）。
    返回按页码升序的页码列表（1-based）。
    """
    sw = prs.slide_width
    holes = []
    for i, slide in enumerate(prs.slides, 1):
        if any(s.has_table for s in slide.shapes):
            continue
        if largest_picture(slide) is not None:
            continue
        content_texts = []
        for s in slide.shapes:
            if not s.has_text_frame:
                continue
            t = s.text_frame.text.strip()
            if not t:
                continue
            if (s.left or 0) < sw * 0.15:      # 左侧导航/页眉，不算内容
                continue
            content_texts.append(t)
        if len(content_texts) == 1 and len(content_texts[0]) <= 20:
            holes.append(i)
    return holes


def largest_picture(slide):
    best = None
    for shp in slide.shapes:
        is_pic = shp.shape_type == 13 or (shp.shape_type is not None
                                          and str(shp.shape_type).startswith("PICTURE"))
        if not is_pic and not shp.element.tag.endswith("}pic"):
            continue
        area = (shp.width or 0) * (shp.height or 0)
        if area <= 0:
            continue
        if best is None or area > (best.width or 0) * (best.height or 0):
            best = shp
    return best


def remove_shape(slide, shp):
    try:
        rids = [b.get(qn("r:embed")) for b in shp.element.iter(qn("a:blip"))]
        shp.element.getparent().remove(shp.element)
        for rid in rids:
            if rid:
                try:
                    slide.part.drop_rel(rid)
                except Exception:
                    pass
    except Exception:
        pass


def duplicate_slide(prs, source):
    """克隆一页并搬运图片资源（背景/导航/页眉一并保留）。"""
    new = prs.slides.add_slide(source.slide_layout)
    for ph in list(new.placeholders):
        ph._element.getparent().remove(ph._element)
    rmap = {}
    for rid, rel in list(source.part.rels.items()):
        if "notesSlide" in rel.reltype or "slideLayout" in rel.reltype:
            continue
        try:
            if rel.is_external:
                rmap[rid] = new.part.rels.get_or_add_ext_rel(rel.reltype, rel._target)
            else:
                rmap[rid] = new.part.rels._add_relationship(rel.reltype, rel._target, None, False)
        except Exception:
            rmap[rid] = None
    for shp in source.shapes:
        el = copy.deepcopy(shp.element)
        for node in el.iter():
            for attr in (qn("r:embed"), qn("r:link"), qn("r:id")):
                v = node.get(attr)
                if v and v in rmap and rmap[v]:
                    node.set(attr, rmap[v])
        new.shapes._spTree.append(el)
    return new


def move_slide_to_end(prs, slide):
    lst = prs.slides._sldIdLst
    target = None
    for entry in lst:
        try:
            if entry.rId == slide.part.rId:
                target = entry
        except Exception:
            pass
    if target is not None:
        lst.remove(target)
        lst.append(target)


def delete_slides(prs, one_based_indexes):
    """按下标（1 基）删除幻灯片，同时清掉关系，避免残留引用。"""
    lst = prs.slides._sldIdLst
    entries = list(lst)
    for i in sorted(set(one_based_indexes), reverse=True):
        if not (1 <= i <= len(entries)):
            continue
        entry = entries[i - 1]
        rid = entry.get(qn("r:id"))
        lst.remove(entry)
        if rid:
            try:
                prs.part.drop_rel(rid)
            except Exception:
                pass


def slide_index(prs, slide):
    for i, s in enumerate(prs.slides):
        if s is slide:
            return i
    return -1


def pick_template(prs):
    best, best_score = None, -1
    for slide in prs.slides:
        score = len(slide.shapes)
        if largest_picture(slide) is not None:
            score += 20
        if any(s.has_table for s in slide.shapes):
            score += 25
        if slide_text(slide).strip():
            score += 5
        if score > best_score:
            best_score, best = score, slide
    return best


def set_slide_title(slide, text):
    """改写看起来是页标题的那个文本框，保留原有字体/颜色。"""
    cands = []
    for shp in slide.shapes:
        if not shp.has_text_frame:
            continue
        t = shp.text_frame.text.strip()
        if not t or len(t) > 20:
            continue
        if (shp.top or 0) < slide.part.package.presentation.slide_height * 0.3:
            cands.append((shp.width or 0, shp))
    if not cands:
        return None
    cands.sort(key=lambda x: -x[0])
    shp = cands[0][1]
    tf = shp.text_frame
    paras = list(tf.paragraphs)
    p = paras[0]
    for extra in list(tf.paragraphs)[1:]:
        extra._p.getparent().remove(extra._p)
    runs = list(p.runs)
    if runs:
        runs[0].text = text
        for extra in runs[1:]:
            extra._r.getparent().remove(extra._r)
    else:
        p.text = text
    return shp


def clear_content(slide):
    """清掉模板页上的图片/表格，保留背景、导航与页眉。"""
    for shp in list(slide.shapes):
        if shp.has_table:
            remove_shape(slide, shp)
        elif shp.shape_type == 13 or shp.element.tag.endswith("}pic"):
            remove_shape(slide, shp)


# ---------------- 封面（第 1 页）改写 ----------------
def _runs_of(shp):
    return [r for para in shp.text_frame.paragraphs for r in para.runs]


def _set_text_keep_format(shp, text):
    """把整个文本框改写成 text，保留第一个 run 的全部格式（字号/粗体/字体/颜色）。"""
    tf = shp.text_frame
    paras = list(tf.paragraphs)
    keep = paras[0]
    for extra in paras[1:]:
        extra._p.getparent().remove(extra._p)
    runs = list(keep.runs)
    if not runs:
        keep.text = text
        return
    runs[0].text = text
    for extra in runs[1:]:
        extra._r.getparent().remove(extra._r)


def _biggest_text_box(slide, exclude_top_ratio=0.0, max_len=40):
    """按面积挑最大的文本形状（用于识别封面院校名/专业名）。"""
    cands = []
    for shp in slide.shapes:
        if not shp.has_text_frame:
            continue
        t = shp.text_frame.text.strip()
        if not t or len(t) > max_len:
            continue
        if exclude_top_ratio and (shp.top or 0) < slide.part.package.presentation.slide_height * exclude_top_ratio:
            continue
        area = (shp.width or 0) * (shp.height or 0)
        if area > 0:
            cands.append((area, shp))
    cands.sort(key=lambda x: -x[0])
    return [c[1] for c in cands]


def set_cover(prs, cover, warn=None, dry=False):
    """改写 PPT 第 1 页封面：把占位院校名/专业名换成本次学员信息。

    cover = {"school": "中国石油大学（北京）", "major": "油气储运工程"}
    识别策略（按优先级）：
      1) JSON 给了 cover.school_box / cover.major_box（形状名），精确匹配；
      2) 文本里含 "XX大学"/"XX专业" 占位符 → 直接替换；
      3) 兜底：封面最大的两个文本框，字号大的当院校、小的当专业。
    """
    if not cover:
        return []
    school = (cover.get("school") or "").strip()
    major = (cover.get("major") or "").strip()
    if not school and not major:
        return []

    slide = prs.slides[0]
    sh = prs.slide_height or 1
    changed = []

    def _apply(shp, new_text, what):
        if not new_text:
            return
        old = shp.text_frame.text.strip()
        if old == new_text:
            return
        if not dry:
            _set_text_keep_format(shp, new_text)
        changed.append((what, old, new_text))

    # ---- 1) 形状名精确指定 ----
    by_name = {}
    if cover.get("school_box") or cover.get("major_box"):
        for shp in slide.shapes:
            by_name[shp.name] = shp
        if cover.get("school_box") in by_name:
            _apply(by_name[cover["school_box"]], school, "院校(按名)")
        if cover.get("major_box") in by_name:
            _apply(by_name[cover["major_box"]], major, "专业(按名)")
        if changed:
            return changed

    # ---- 2) 占位符文本匹配（XX大学 / XX专业 / 某某大学 ...） ----
    import re as _re
    ph_school = _re.compile(r"^(X{1,4}|某{1,4}|待填|【.*?】|\(.*?\)|（.*?）)?(大学|学院|大學)$")
    ph_major = _re.compile(r"^(X{1,4}|某{1,4}|待填|【.*?】)(专业|类)?$")
    for shp in slide.shapes:
        if not shp.has_text_frame:
            continue
        t = shp.text_frame.text.strip()
        if not t or len(t) > 40:
            continue
        if ph_school.match(t) and school:
            _apply(shp, school, "院校(占位符)")
            continue
        if ph_major.match(t) and major:
            _apply(shp, major, "专业(占位符)")
            continue
    if changed:
        return changed

    # ---- 3) 兜底：按字号 + 面积挑两个最大的文本框 ----
    boxes = _biggest_text_box(slide, max_len=40)
    if not boxes:
        if warn:
            warn("封面：没找到可改写的文本框，已跳过（可用 cover.school_box 指定形状名）")
        return []
    ranked = []
    for shp in boxes:
        rs = _runs_of(shp)
        sz = max([r.font.size.pt for r in rs if r.font.size] or [0])
        ranked.append((sz, (shp.width or 0) * (shp.height or 0), shp))
    ranked.sort(key=lambda x: (-x[0], -x[1]))
    if school and len(ranked) >= 1:
        _apply(ranked[0][2], school, "院校(按字号)")
    if major:
        rest = [x for x in ranked if x[2] is not ranked[0][2]]
        if rest:
            _apply(rest[0][2], major, "专业(按字号)")
    if not changed and warn:
        warn("封面：兜底策略未命中，封面保持原样")
    return changed


# ---------------- 铁律自检：院校层次 ----------------
TIER_ORDER = {"冲刺档": 0, "稳妥档": 1, "保底档": 2}
# 非 985/211 但符合铁律的"中科院系"特例
SPECIAL_OK = ("中国科学院大学", "中国科学院", "中科院", "国科大")
# 认可的国内层次标签
LEVEL_OK = ("985", "211", "双一流")
# 层次高低分档（用于"稳妥/保底不得低于学员本科层次"校验）
# 注意：铁律规定「中国科学院大学及中科院各研究所视同顶尖 985，可用于任何档位」，
# 因此 985 出身的学员把院所放进稳妥 / 保底档时必须判为"不低于本科层次"，
# 否则会把合规行误判为不合格（2026-10-05 北大计算机学员案例实测）。
LEVEL_SCORE = {"985": 3, "211": 2, "双一流": 1, "双非": 0, "四非": 0, "中科院特例": 3}
# 冲刺档允许的层次（中上 985；中科院系视同顶尖 985）
SPRINT_OK_LEVELS = ("985", "中科院特例")
# 海外兜底关键词（只允许出现在保底档末行，且学员须有中外合作/双学位背景）
OVERSEAS_OK = ("海外", "境外")
OVERSEAS_STUDENT_HINTS = ("中外合作", "双学位", "留学", "2+2", "3+1", "4+0", "4+1",
                          "国际班", "合作办学")


def school_level(row):
    """返回 (层次标签, 软科位次或 None, 是否合规)。

    层次来源优先级：row["level"] 字段 > school 文本里的 985/211/双一流。
    软科位次来源：row["rank"] 字段 > school 文本里的"软科第N"。
    """
    if not isinstance(row, dict):
        row = {}
    name = str(row.get("school", ""))
    level = str(row.get("level", "")).strip()

    rank = row.get("rank")
    try:
        rank = int(rank) if rank not in (None, "") else None
    except (TypeError, ValueError):
        rank = None
    if rank is None:
        m = re.search(r"软科\s*第?\s*(\d+)", name)
        if m:
            rank = int(m.group(1))

    if row.get("overseas") or any(k in level for k in OVERSEAS_OK):
        return ("海外兜底", rank, True)
    if any(k in name for k in SPECIAL_OK) or any(k in level for k in SPECIAL_OK):
        return ("中科院特例", rank, True)
    if any(k in level for k in LEVEL_OK):
        return (level, rank, True)
    if any(k in name for k in LEVEL_OK):
        return ("（按校名判定）", rank, True)
    return ("未标注", rank, False)


def validate_tiers(data, warn=None):
    """铁律自检：定档院校的层次、每档数量与软科排名梯度。返回 (ok, 报告行列表)。

    检查项：
      1) 每所国内院校必须能判定为 985 / 211 / 双一流 / 中科院特例；
      2) 唯一例外是"海外兜底"行（overseas=true）：只能放保底档、且必须是最后一行，
         学员还必须有中外合作办学 / 双学位 / 留学背景；
      3) 冲、稳、保各至少 2 所（推荐各 4 所，总数 12 行）；
      4) 档位顺序必须是 冲刺→稳妥→保底；
      5) 软科排名梯度：冲刺档均位次应优于保底档均位次。若因学科实力（如本校 A+ 学科）
         而有逆转，只提示、不判不合格。
    """
    tiers = data.get("school_tiers") or []
    lines, ok = [], True

    if len(tiers) < 6:
        ok = False
        lines.append("定档院校共 %d 行（至少 6 行，推荐 12 行）" % len(tiers))

    counts, seq = {}, []
    by_tier = {"冲刺档": [], "稳妥档": [], "保底档": []}
    ranks_by_tier = {"冲刺档": [], "稳妥档": [], "保底档": []}
    for i, r in enumerate(tiers):
        tag, rank, good = school_level(r)
        tier = r.get("tier", "?")
        counts[tier] = counts.get(tier, 0) + 1
        seq.append(TIER_ORDER.get(tier, 9))
        if tier in by_tier:
            by_tier[tier].append((tag, r.get("school", "")))
        if rank is not None and tier in ranks_by_tier:
            ranks_by_tier[tier].append(rank)

        overseas = tag == "海外兜底"
        if not good:
            ok = False
        if overseas:
            if tier != "保底档":
                ok = False
                lines.append("!! 海外兜底行「%s」只能放保底档（现为 %s）"
                             % (r.get("school", ""), tier))
            elif i != len(tiers) - 1:
                ok = False
                lines.append("!! 海外兜底行「%s」应排在保底档最后一行" % r.get("school", ""))

        lines.append("%s%-4s %-26s -> %s%s" % (
            "OK " if good else "!! ", tier, r.get("school", ""), tag,
            "" if rank is None else "（软科第 %d）" % rank))
        if rank is None and not overseas:
            lines.append("     ^ 提示：没给软科位次（rank 字段或校名里的「软科第N」），无法校验梯度")

    for t in ("冲刺档", "稳妥档", "保底档"):
        n = counts.get(t, 0)
        if n < 2:
            ok = False
            lines.append("!! %s 只有 %d 所（每档至少 2 所）" % (t, n))
        elif n != 4:
            lines.append("提示：%s 有 %d 所（新标准推荐每档 4 所）" % (t, n))

    if seq != sorted(seq):
        ok = False
        lines.append("!! 档位顺序不是 冲刺→稳妥→保底，请检查数据顺序")

    # 海外兜底与学员背景是否对得上（只提示）
    has_overseas = any(r.get("overseas") for r in tiers if isinstance(r, dict))
    if has_overseas:
        blob = json.dumps(data.get("student_info") or {}, ensure_ascii=False)
        if not any(h in blob for h in OVERSEAS_STUDENT_HINTS):
            lines.append("提示：定档表里有海外兜底行，但学员信息里没提到中外合作 / 双学位 / "
                         "留学项目，请确认该兜底路径真实存在。")

    # ---- 档位层次铁律（2026-09-19 用户定版）----
    # 1) 冲刺档必须全部为中上 985（优先 C9 / 华五），中科院系视同顶尖 985
    for tag, name in by_tier["冲刺档"]:
        if tag not in SPRINT_OK_LEVELS:
            ok = False
            lines.append("!! 冲刺档「%s」层次为 %s —— 冲刺档必须是中上 985（优先 C9 / 华五），"
                         "非 985 一律移到稳妥档。" % (name, tag))

    # 2) 稳妥档必须"大部分是 985"：985 数量过半（4 所档位即 ≥3 所）
    wen = by_tier["稳妥档"]
    if wen:
        n985 = sum(1 for tag, _ in wen if tag == "985")
        if n985 * 2 <= len(wen):
            ok = False
            lines.append("!! 稳妥档 985 只有 %d/%d 所 —— 稳妥档必须大部分是 985，"
                         "请替换掉非 985 院校。" % (n985, len(wen)))

    # 3) 稳妥 / 保底档不得低于学员本科院校层次（以 meta.undergrad_level 为基准）
    ug = str((data.get("meta") or {}).get("undergrad_level", "")).strip()
    ug_score = LEVEL_SCORE.get(ug)
    if ug_score is None:
        lines.append("提示：meta 里没写 undergrad_level（985 / 211 / 双一流 / 双非），"
                     "无法校验「稳妥与保底档不低于学员本科层次」。")
    else:
        for tier in ("稳妥档", "保底档"):
            for tag, name in by_tier[tier]:
                if tag == "海外兜底":
                    continue
                sc = LEVEL_SCORE.get(tag, LEVEL_SCORE.get(tag[:3], None))
                if sc is None:
                    sc = 3 if "985" in tag else (2 if "211" in tag else 1)
                if sc < ug_score:
                    ok = False
                    lines.append("!! %s「%s」层次 %s 低于学员本科层次 %s —— "
                                 "稳妥档与保底档一律不得低于学员本科院校层次。"
                                 % (tier, name, tag, ug))

    # 软科排名梯度（仅提示）
    def _avg(v):
        return sum(v) / len(v) if v else None
    a_s, a_w, a_b = (_avg(ranks_by_tier["冲刺档"]),
                     _avg(ranks_by_tier["稳妥档"]),
                     _avg(ranks_by_tier["保底档"]))
    if a_s is not None and a_b is not None:
        if a_s > a_b:
            lines.append("提示：冲刺档均位次 %.0f 差于保底档 %.0f —— 若因学科实力（如本校 A+ 学科）"
                         "而有意为之可忽略，否则请按软科排名调整档位。" % (a_s, a_b))
        else:
            lines.append("软科排名梯度：冲刺均 %.0f < 稳妥均 %s < 保底均 %.0f（位次越好数字越小）"
                         % (a_s, "%.0f" % a_w if a_w else "—", a_b))
    return ok, lines


# ---------------- 表格构建 ----------------
def allocate_widths(total_px, rows, n_cols, weights, min_px=None, max_px=None):
    """按"内容长度 + 权重"分配列宽，再在 [min_px, max_px] 约束下归一化到 total_px。

    两条经验（都是踩过坑的）：
      - 标签列（学员情况 / 梯度 / 时间）内容很短，纯按内容分配会被挤成竖条，
        必须给 **min_px 硬下限**；
      - 长文本列不能被允许吃掉整行，否则两列表会变成 1:1 五五开，
        所以要给 **max_px 硬上限**。
    """
    content = []
    for j in range(n_cols):
        units = 2.0
        for r in rows:
            if j < len(r):
                units = max(units, max_line_units(r[j]))
        content.append(units)
    raw = [content[j] * 0.8 + max(weights[j], 0.01) * 70 for j in range(n_cols)]
    s = sum(raw) or 1.0
    w = [total_px * x / s for x in raw]

    lo, hi = [], []
    for j in range(n_cols):
        a = float(min_px[j]) if (min_px and j < len(min_px) and min_px[j]) else total_px * 0.05
        b = float(max_px[j]) if (max_px and j < len(max_px) and max_px[j]) else total_px
        lo.append(a)
        hi.append(max(a, b))
    if sum(lo) > total_px:                    # 下限之和超总宽：整体按比例缩下限
        k = total_px / sum(lo)
        lo = [x * k for x in lo]
        hi = [max(lo[j], hi[j]) for j in range(n_cols)]
    w = [min(max(w[j], lo[j]), hi[j]) for j in range(n_cols)]

    for _ in range(80):                       # 迭代把差额摊给"还没顶到边界"的列
        diff = total_px - sum(w)
        if abs(diff) < 0.5:
            break
        if diff > 0:
            free = [j for j in range(n_cols) if w[j] < hi[j] - 0.5]
        else:
            free = [j for j in range(n_cols) if w[j] > lo[j] + 0.5]
        if not free:
            break
        base = sum(w[j] for j in free) or 1.0
        for j in free:
            w[j] = min(max(w[j] + diff * w[j] / base, lo[j]), hi[j])
    if abs(total_px - sum(w)) > 0.5:          # 还剩零头就甩给最宽的一列
        j = max(range(n_cols), key=lambda k: w[k])
        w[j] = max(w[j] + total_px - sum(w), lo[j])
    return w


def estimate_row(row, widths_px, font_px, is_header=False):
    per_unit = 1.0 / font_px if font_px else 1.0
    lines = 1
    for j, cell in enumerate(row):
        avail = max(widths_px[j] - CELL_H_PAD, 10.0) * per_unit
        n = sum(wrap_count(ln, avail) for ln in text_lines(cell))
        lines = max(lines, n)
    line_h = font_px * LINE_FACTOR * (HEADER_BUMP if is_header else 1.0)
    pad = 2 * CELL_PAD * font_px / 12.0
    return lines * line_h + pad, lines


def total_height(all_rows, widths_px, font_pt):
    font_px = font_pt / 72.0 * PX_PER_IN
    h, details = 0.0, []
    for i, r in enumerate(all_rows):
        px, lines = estimate_row(r, widths_px, font_px, is_header=(i == 0))
        px *= HEIGHT_SAFETY
        h += px
        details.append((px, lines))
    return h, details


def apply_border(cell, color_hex, width_emu):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
    prev = None
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        el = etree.Element(qn(tag), {"w": str(int(width_emu)), "cap": "flat",
                                     "cmpd": "sng", "algn": "ctr"})
        fill = etree.SubElement(el, qn("a:solidFill"))
        etree.SubElement(fill, qn("a:srgbClr")).set("val", color_hex)
        etree.SubElement(el, qn("a:prstDash")).set("val", "solid")
        if prev is None:
            tcPr.insert(0, el)
        else:
            prev.addnext(el)
        prev = el


def write_cell(cell, text, font_pt, bold, color_hex, align_center, font_scale):
    tf = cell.text_frame
    tf.word_wrap = True
    cell.margin_left = Emu(int(0.05 * EMU_PER_IN * font_scale))
    cell.margin_right = Emu(int(0.05 * EMU_PER_IN * font_scale))
    cell.margin_top = Emu(int(0.012 * EMU_PER_IN * font_scale))
    cell.margin_bottom = Emu(int(0.012 * EMU_PER_IN * font_scale))
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    bodyPr = tf._txBody.find(qn("a:bodyPr"))
    if bodyPr is not None:
        for child in list(bodyPr):
            if child.tag in (qn("a:spAutoFit"), qn("a:autofit"), qn("a:normAutofit")):
                bodyPr.remove(child)
        if bodyPr.find(qn("a:noAutofit")) is None:
            bodyPr.append(etree.SubElement(bodyPr, qn("a:noAutofit")))
    lines = text_lines(text)
    paras = list(tf.paragraphs)
    paras[0].text = lines[0] if lines else ""
    for extra in paras[1:]:
        extra._p.getparent().remove(extra._p)
    for ln in lines[1:]:
        tf.add_paragraph().text = ln
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.CENTER if align_center else PP_ALIGN.LEFT
        p.line_spacing = LINE_FACTOR
        p.space_after = Pt(0)
        p.space_before = Pt(0)
        for r in p.runs:
            set_font(r, FONT_NAME, font_pt, bold, color_hex)


def build_table(slide, spec, rows, box, theme, max_font, min_font, font_scale,
                fill_height=True, header=None):
    left, top, width, height = box
    n_cols = len(spec["columns"])
    head = list(header) if header else list(spec["columns"])
    all_rows = [head] + [list(r) for r in rows]
    # 「成绩」列整列纵向合并：先把各行该列文案统一替换成固定短文案，
    # 这样行高估算、列宽分配都按合并后的真实内容算（合并格跨全部数据行，不约束行高）。
    vmerge = spec.get("vmerge") or None
    vcol = int(vmerge["col"]) if vmerge else -1
    if vmerge:
        vtext = str(vmerge.get("text", ""))
        for r in all_rows[1:]:
            if 0 <= vcol < len(r):
                r[vcol] = vtext
    widths_px = allocate_widths(emu2px(width), all_rows, n_cols, spec["weights"],
                                min_px=spec.get("min_px"), max_px=spec.get("max_px"))
    avail_px = emu2px(height)

    chosen, details = min_font, None
    p = int(round(max_font * 2))
    low = int(round(min_font * 2))
    while p >= low:
        fp = p / 2.0
        h, d = total_height(all_rows, widths_px, fp)
        if h <= avail_px:
            chosen, details = fp, d
            break
        p -= 1
    if details is None:
        _, details = total_height(all_rows, widths_px, chosen)

    total_px = sum(d[0] for d in details)
    font_pt = max(chosen * font_scale, 5.0)
    graphic = slide.shapes.add_table(len(all_rows), n_cols, left, top, width,
                                     Emu(px2emu(min(total_px, avail_px))))
    tbl = graphic.table
    tblPr = tbl._tbl.find(qn("a:tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(tbl._tbl, qn("a:tblPr"))
    for c in list(tblPr):
        if c.tag == qn("a:tableStyleId"):
            tblPr.remove(c)
    etree.SubElement(tblPr, qn("a:tableStyleId")).text = \
        "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    tblPr.set("firstRow", "1")
    tblPr.set("bandRow", "0")

    # 列宽：直接写 XML，绕开 python-pptx 的 notify_width_changed 累加逻辑
    total_emu = sum(px2emu(w) for w in widths_px)
    grid = tbl._tbl.find(qn("a:tblGrid"))
    if grid is not None:
        cols = grid.findall(qn("a:gridCol"))
        for j, w in enumerate(widths_px):
            if j < len(cols):
                cols[j].set("w", str(int(px2emu(w))))
    graphic.width = Emu(int(total_emu))

    if fill_height and total_px > 0:
        extra = avail_px - total_px
        row_px = [d[0] + extra * (d[0] / total_px) for d in details]
    else:
        row_px = [d[0] for d in details]
    min_row_px = font_pt / 72.0 * PX_PER_IN * LINE_FACTOR + 2
    for i, px in enumerate(row_px):
        tbl.rows[i].height = px2emu(max(px, min_row_px))

    # 首列同值纵向合并（定档表的"梯度"列：冲刺档一次占 4 行）
    merged_skip = set()
    if spec.get("merge_first") and rows:
        a = 0
        prev = rows[0][0] if rows[0] else None
        for idx in range(1, len(rows) + 1):
            cur = rows[idx][0] if idx < len(rows) and rows[idx] else None
            if cur != prev:
                if idx - a >= 2:
                    try:
                        tbl.cell(a + 1, 0).merge(tbl.cell(idx, 0))
                        for i2 in range(a + 2, idx + 1):
                            merged_skip.add((i2, 0))
                    except Exception:
                        pass  # 合并失败就退化成"不合并"，不影响出表
                a, prev = idx, cur

    # 指定列整列纵向合并（主线任务表"成绩"列 -> 固定文案，见 SPECS vmerge）
    if vmerge and len(all_rows) >= 3:
        try:
            tbl.cell(1, vcol).merge(tbl.cell(len(all_rows) - 1, vcol))
            for i2 in range(2, len(all_rows)):
                merged_skip.add((i2, vcol))
        except Exception:
            pass  # 合并失败就退化成每行独立写文案，不影响出表

    center_cols = set(spec.get("center_cols") or [])
    warnings = []
    for i, r in enumerate(all_rows):
        for j, val in enumerate(r):
            if (i, j) in merged_skip:
                continue
            cell = tbl.cell(i, j)
            if i == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(theme["header_bg"])
                write_cell(cell, val, round(font_pt * 0.96, 1), True,
                           theme["header_fg"], True, font_scale)
            elif j == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(theme["firstcol_bg"])
                write_cell(cell, val, round(font_pt, 1), True,
                           theme["firstcol_fg"], (0 in center_cols), font_scale)
            else:
                # 纵向合并格（成绩列）水平居中，视觉与样图一致
                align_c = (j in center_cols) or (vmerge and i == 1 and j == vcol)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(theme["body_bg"])
                write_cell(cell, val, round(font_pt, 1), False, BODY_FG,
                           align_c, font_scale)
            apply_border(cell, theme["border"], max(int(9525 * font_scale), 7620))
            if i and details[i][1] > MAX_LINES_PER_CELL:
                warnings.append("第 %d 行第 %d 列约 %d 行，建议在 JSON 里再精简"
                                % (i, j + 1, details[i][1]))
    return {
        "font_pt": round(font_pt, 1),
        "table_px": round(sum(row_px), 0),
        "avail_px": round(avail_px, 0),
        "widths": [round(w) for w in widths_px],
        "warnings": warnings,
        "overflow": total_px > avail_px + 1,
    }


# ---------------- 文案精简（无 PPT 版数据时的兜底） ----------------
_SENT = re.compile(r"[。；;！？]")


def cut(text, cap):
    text = str(text or "").strip()
    if len(text) <= cap:
        return text
    pieces = _SENT.split(text)
    cur = ""
    for i in range(0, len(pieces), 2):
        seg = pieces[i]
        tail = pieces[i + 1] if i + 1 < len(pieces) else ""
        if cur and len(cur) + len(seg) + len(tail) > cap:
            break
        cur += seg + tail
    if not cur:
        cur = text[:cap]
    cur = cur.rstrip("。；;，,、 ")
    return cur + "…" if len(text) > len(cur) else cur


def listify(text, cap):
    text = str(text or "").strip().replace(";", "；")
    items = [x.strip() for x in text.split("；") if x.strip()]
    out, total = [], 0
    for it in items:
        it = re.sub(r"^\s*\d{1,2}[.、]\s*", "", it)
        it = cut(it, max(8, cap // 6))
        if total + len(it) > cap:
            break
        out.append(it)
        total += len(it)
    return "\n".join(out) if out else cut(text, cap)


def base_rows(data, spec):
    """取该表的完整版数据行（dict 副本列表）。"""
    k = spec["key"]
    if k == "student_info":
        si = data.get("student_info")
        si = si if isinstance(si, dict) else {}
        # 2026-09-19 用户定版：学员信息表删掉「院校定位」行（三档院校以定档表为准），
        # JSON 里就算写了也会被过滤，新旧数据源都能跑。
        return [dict(r) for r in (si.get("rows") or [])
                if isinstance(r, dict) and "院校定位" not in str(r.get("label", ""))]
    return [dict(r) for r in (data.get(k) or []) if isinstance(r, dict)]


def overlay(rows, ov):
    """把 ppt_tables 的短文案**逐字段**盖到完整版上——只覆盖真的给了值的字段。

    这样同事只需要在 ppt_tables 里写"和 Excel 不一样的那几格"，
    不用把整张表抄一遍（例如只写 school_tiers[i].requirement）。
    """
    if not isinstance(ov, list):
        return rows
    for i, r in enumerate(rows):
        if i >= len(ov) or not isinstance(ov[i], dict):
            continue
        for k, v in ov[i].items():
            if str(k).startswith("_") or v in (None, ""):
                continue
            r[k] = v
    return rows


def auto_condense(data, spec, rows):
    """--ppt-trim 时启用的兜底压缩：把超长文案按句号/分号截断。默认不启用。"""
    caps = _caps(spec)
    k = spec["key"]
    out = []
    for r in rows:
        if k == "student_info":
            out.append([r.get("label", ""), cut(r.get("content", ""), caps["content"])])
        elif k == "school_tiers":
            out.append([r.get("tier", ""), r.get("school", ""),
                        cut(r.get("direction", ""), caps["direction"]),
                        cut(r.get("fit", ""), caps["fit"])])
        else:
            out.append([r.get("time", ""), cut(r.get("score", ""), caps["score"]),
                        cut(r.get("english", ""), caps["english"]),
                        cut(r.get("competition", ""), caps["competition"]),
                        cut(r.get("research", ""), caps["research"]),
                        cut(r.get("other", ""), caps["other"])])
    return out


def get_rows(data, spec, has_ppt_version, trim=False):
    """取该表最终要渲染的文案：完整版 -> ppt_tables 逐字段覆盖 ->（可选）截断。"""
    rows = base_rows(data, spec)
    if has_ppt_version:
        ppt = data.get("ppt_tables")
        ppt = ppt if isinstance(ppt, dict) else {}
        rows = overlay(rows, ppt.get(spec["key"]))
    if trim:
        return auto_condense(data, spec, rows)
    return [[str(r.get(f, "")) for f in FIELD_KEYS[spec["key"]]] for r in rows]


def header_row(data, spec):
    """表头行：学员信息表第 2 格是学员摘要（headline），其余表用固定列名。"""
    head = list(spec["columns"])
    key = spec.get("header_from")
    if key:
        si = data.get("student_info")
        si = si if isinstance(si, dict) else {}
        head[1] = str(si.get(key, "") or head[1])
    return head


def audit_ppt_text(data, warn):
    """体检最终渲染文案是否超出字数上限（超了只告警——不影响出表，但字号会被压小）。"""
    issues = []
    for spec in SPECS:
        caps = spec["caps"]
        fields = FIELD_KEYS[spec["key"]]
        for idx, r in enumerate(get_rows(data, spec, True), 1):
            for j, val in enumerate(r):
                f = fields[j]
                cap = caps.get(f)
                if cap is None:
                    continue
                n = len(str(val).replace("\n", ""))
                if n > cap:
                    issues.append("[%s] 第%d行 %s：%d 字 > 上限 %d 字"
                                  % (spec["title"], idx, f, n, cap))
    hl = str(((data.get("student_info") or {}).get("headline") or ""))
    cap_hl = SPECS[0]["caps"]["headline"]
    if hl and len(hl) > cap_hl:
        issues.append("[学员信息] 表头摘要：%d 字 > 上限 %d 字" % (len(hl), cap_hl))
    if issues:
        warn("有 %d 处 PPT 文案超出字数上限（可精简，或在 ppt_tables 里写短版覆盖）："
             % len(issues))
        for it in issues[:12]:
            warn("   " + it)
        if len(issues) > 12:
            warn("   …另有 %d 处" % (len(issues) - 12))
    return issues


# ---------------- 模板解析（同事分发版：支持省略模板参数、自动发现） ----------------
def templates_dir():
    """<skill 根目录>/assets/templates/"""
    return os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "assets", "templates")


def resolve_template(src_path):
    """解析模板 PPT 路径。

    给了且文件存在  → 直接用；
    没给 / 给的找不到 → 在 assets/templates/ 下自动找唯一的 .pptx。
    这样使用者只要把自己的模板丢进那个目录，命令行不用改。
    """
    if src_path and os.path.isfile(src_path):
        return src_path
    tdir = templates_dir()
    found = []
    if os.path.isdir(tdir):
        found = sorted(f for f in glob.glob(os.path.join(tdir, "*.pptx"))
                       if not os.path.basename(f).startswith("~$"))
    if not found:
        print("× 找不到模板 PPT。做法二选一：")
        print("  ① 命令行第 2 个参数显式给模板路径；")
        print("  ② 把自己的模板 pptx 放进 %s（放一个即可）" % tdir)
        if src_path:
            print("  你给的路径不存在：%s" % src_path)
        sys.exit(1)
    if src_path:
        print("  提示：模板 %s 不存在，改用 %s" % (src_path, os.path.basename(found[0])))
    if len(found) > 1:
        print("  提示：assets/templates/ 下有 %d 个 pptx，取第一个：%s"
              % (len(found), os.path.basename(found[0])))
    print("  模板：%s" % found[0])
    return found[0]


# ---------------- 主流程 ----------------
def compute_box(prs, slide, fallback):
    """2026-10-12 用户定版：表格一律放标准版心，不再跟随源 PPT 占位截图的位置。

    旧逻辑按占位图定位并向下方延展 1.35 倍——截图位置不可控，山东大学的占位图
    在 2.27in，表格被直接压上左侧导航栏且整体超高。各院校模板的内容区结构一致
    （左导航 + 页标题 + 空白内容区），标准版心通用；占位图仍由调用方删除，
    只是不再影响表格位置。"""
    pic = largest_picture(slide)
    return fallback, pic


# ---------------- 输出路径解析（2026-09-21 用户定版：PPT 默认交付到桌面） ----------------
def desktop_dir():
    """返回本机桌面目录；找不到就回退到当前工作目录。

    覆盖 Windows（含 OneDrive 重定向）、macOS、Linux(XDG) 三种常见情形。
    """
    home = os.path.expanduser("~")
    cands = []
    if os.name == "nt":
        up = os.environ.get("USERPROFILE") or home
        cands += [os.path.join(up, "Desktop"),
                  os.path.join(up, "OneDrive", "Desktop"),
                  os.path.join(home, "Desktop")]
        # OneDrive 中文/企业目录名不固定，扫一遍家目录下含 Desktop 的一级目录
        try:
            for nm in os.listdir(up):
                p = os.path.join(up, nm, "Desktop")
                if os.path.isdir(p):
                    cands.append(p)
        except OSError:
            pass
    else:
        cands += [os.path.join(home, "Desktop"), os.path.join(home, "桌面")]
        xdg = os.path.join(home, ".config", "user-dirs.dirs")
        if os.path.exists(xdg):
            try:
                with open(xdg, "r", encoding="utf-8", errors="ignore") as f:
                    for line in f:
                        m = re.match(r'XDG_DESKTOP_DIR\s*=\s*"(.+)"', line.strip())
                        if m:
                            cands.insert(0, m.group(1).replace("$HOME", home))
            except OSError:
                pass
    for c in cands:
        if c and os.path.isdir(c):
            return c
    return os.getcwd()


def default_out_name(data_path, data):
    """没给输出文件名时按学员院校拼一个：保研规划_<院校>.pptx"""
    meta = data.get("meta") or {}
    school = (data.get("cover") or {}).get("school") or meta.get("school") or ""
    school = re.sub(r'[\\/:*?"<>|]', "", str(school)).strip()
    base = os.path.splitext(os.path.basename(data_path))[0]
    m = re.search(r"student_data_(.+)$", base)
    if not school and m:
        school = m.group(1)
    if not school:
        school = base or "学员"
    return "保研规划_%s.pptx" % school


def resolve_out_path(raw, data_path, data, opts):
    """解析输出路径（2026-09-21 用户定版：交付物直接落在桌面）。

    优先级：
      1. --dir <路径>           强制指定输出目录（最高优先级）
      2. 显式给了绝对/带目录路径  原样尊重（同事自定义场景）
      3. 只给了文件名            → 桌面/该文件名
      4. 什么都没给              → 桌面/保研规划_<院校>.pptx
      --no-desktop 可关掉桌面默认，退回"当前目录"的老行为。
    """
    if raw:
        if os.path.isabs(raw) or os.path.dirname(raw):
            return os.path.abspath(raw)
        name = raw
    else:
        name = default_out_name(data_path, data)
    if not name.lower().endswith(".pptx"):
        name += ".pptx"
    d = opts.get("dir") if isinstance(opts.get("dir"), str) else None
    if d:
        return os.path.abspath(os.path.join(d, name))
    if opts.get("no-desktop"):
        return os.path.abspath(name)
    return os.path.join(desktop_dir(), name)


def main():
    argv = sys.argv[1:]
    opts, pos = {}, []
    i = 0
    while i < len(argv):
        a = argv[i]
        if a.startswith("--"):
            if i + 1 < len(argv) and not argv[i + 1].startswith("--"):
                opts[a[2:]] = argv[i + 1]
                i += 2
            else:
                opts[a[2:]] = True
                i += 1
        else:
            pos.append(a)
            i += 1
    if len(pos) < 1:
        print(__doc__)
        sys.exit(1)
    data_path = pos[0]
    src_path = pos[1] if len(pos) > 1 else None  # 模板 PPT，可省略（自动发现）
    out_raw = pos[2] if len(pos) > 2 else None   # 真实输出路径待 data 读入后解析（默认桌面）

    dry_run = bool(opts.get("dry-run"))
    _warns = []

    def warn_wrap(msg):
        _warns.append(msg)
        print("  ⚠ %s" % msg)

    # 文案密度：默认＝精简版（字数上限已较原标准减半）；--ppt-verbose 回到原标准
    global USE_VERBOSE_CAPS
    USE_VERBOSE_CAPS = bool(opts.get("ppt-verbose"))

    with open(data_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # 输出路径：默认落在桌面（2026-09-21 用户定版），详见 resolve_out_path()
    out_path = resolve_out_path(out_raw, data_path, data, opts)
    print("输出目录: %s" % os.path.dirname(os.path.abspath(out_path)))

    legacy = [k for k in ("student_profile", "main_tasks", "stage_tasks") if k in data]
    if legacy:
        print("× 这份 JSON 是旧版四表结构（发现字段：%s）。" % "、".join(legacy))
        print("  新版结构：student_info / school_tiers / main_timeline（main_timeline 已下线不渲染）。")
        print("  请对照 assets/sample_data.json 重新生成数据源。")
        sys.exit(1)
    # 2026-09-19 起 main_timeline 已下线（PPT 不渲染、模板第 4 页已删），
    # 因此它不再是必需字段——缺了照样出表，只是不渲染。若哪天恢复主线任务表，
    # 需要同时把 "main_timeline" 加回这个元组、并恢复 SPECS 里注释掉的那段 spec。
    missing = [k for k in ("student_info", "school_tiers") if k not in data]
    if missing:
        print("× JSON 缺少必需字段：%s（对照 assets/sample_data.json）" % "、".join(missing))
        sys.exit(1)

    src_path = resolve_template(src_path)
    prs = Presentation(src_path)
    sw, sh = prs.slide_width, prs.slide_height
    w_in, h_in = sw / EMU_PER_IN, sh / EMU_PER_IN
    font_scale = max(0.72, min(1.25, ((w_in / 13.333) * (h_in / 7.5)) ** 0.5))
    max_font = float(opts.get("max-font", 12))
    min_font = float(opts.get("min-font", 7))

    theme = {"header_bg": HEADER_BG, "header_fg": HEADER_FG,
             "firstcol_bg": FIRSTCOL_BG, "firstcol_fg": FIRSTCOL_FG,
             "body_bg": BODY_BG, "border": BORDER_COLOR}
    tv = str(opts.get("theme", "")).strip()
    if tv.lower() == "auto":
        acc = extract_accent(prs)
        if acc:
            theme["header_bg"] = acc
            theme["firstcol_bg"] = mix(acc, "FFFFFF", 0.86)
            theme["border"] = mix(acc, "FFFFFF", 0.55)
            theme["firstcol_fg"] = mix(acc, "000000", 0.25)
            print("已提取 PPT 强调色: #%s" % acc)
    elif "," in tv:
        parts = [x.strip().lstrip("#").upper() for x in tv.split(",")]
        theme["header_bg"] = parts[0]
        if len(parts) > 1:
            theme["firstcol_bg"] = parts[1]
        if len(parts) > 2:
            theme["border"] = parts[2]
    elif tv:
        theme["header_bg"] = tv.lstrip("#").upper()

    explicit = None
    if opts.get("slides"):
        explicit = [int(x) for x in str(opts["slides"]).split(",") if x.strip()]
    append_mode = bool(opts.get("append")) or str(opts.get("mode")) == "append"
    want = len(SPECS)
    # 原始"空内容页"（只剩页标题、内容区全空）——出表后未被用到的会被删掉
    orig_holes = content_hole_slides(prs)
    # 自动页序：先按关键词匹配；若匹配不足 want 张，用"空内容页"补齐（内置模板/已清空的 PPT 走这条）
    if explicit is None and not append_mode and not opts.get("no-auto-strings"):
        matched = {}
        for spec in SPECS:
            best_i, best_score = None, 0
            for i2, s2 in enumerate(prs.slides):
                t = slide_text(s2)
                sc = max([len(kw) for kw in spec["match"] if kw in t] or [0])
                if sc > best_score:
                    best_score, best_i = sc, i2 + 1
            if best_score >= 2:
                matched[spec["key"]] = best_i
        if len(set(matched.values())) < want:
            holes = [h for h in orig_holes if h not in matched.values()]
            # 关键词命中的页若与空页重叠，以空页为准（空页更可能是待填位置）
            need = [sp["key"] for sp in SPECS if sp["key"] not in matched
                    or matched[sp["key"]] not in orig_holes]
            for key, hole in zip(need, holes):
                matched[key] = hole
        if len(set(matched.values())) == want:
            explicit = [matched[sp["key"]] for sp in SPECS]
            print("自动页序  : %s（%s）"
                  % (explicit, " / ".join(sp["title"] for sp in SPECS)))

    after_n = int(opts.get("after", 0)) if opts.get("after") else None

    print("源文件  : %s" % src_path)
    print("页面尺寸: %.2f x %.2f 英寸（%.0f x %.0f px，比例 %.2f），字号基准缩放 %.2f"
          % (w_in, h_in, emu2px(sw), emu2px(sh), w_in / h_in, font_scale))
    print("模式    : %s" % ("追加新页" if append_mode else "就地替换原表/图片"))
    print("文案密度: %s" % ("精简（默认，超限会在下方告警）" if not USE_VERBOSE_CAPS else "放宽（--ppt-verbose）"))
    print("-" * 74)

    # ---- PPT 文案字数体检 ----
    if not USE_VERBOSE_CAPS:
        audit_ppt_text(data, warn_wrap)

    # ---- 封面（第 1 页）改写：把院校/专业换成本次学员信息 ----
    cover = data.get("cover") if isinstance(data.get("cover"), dict) else None
    if cover is None and not opts.get("no-cover"):
        # 没写 cover 块时，从 meta 里兜底推断
        meta = data.get("meta") if isinstance(data.get("meta"), dict) else {}
        guess_school = meta.get("school")
        guess_major = meta.get("major")
        if guess_school or guess_major:
            cover = {"school": guess_school or "", "major": guess_major or ""}
            print("封面信息  : 未提供 cover 块，已从 meta 推断 -> %s / %s"
                  % (cover["school"], cover["major"]))
    if cover and not opts.get("no-cover"):
        hits = set_cover(prs, cover, warn=warn_wrap, dry=dry_run)
        if hits:
            for what, old, new in hits:
                print("封面改写  : %s  「%s」-> 「%s」" % (what, old, new))
        else:
            print("封面改写  : 未命中可改写文本（封面保持原样）")
    else:
        print("封面改写  : 已跳过（--no-cover）")

    # ---- 铁律自检：院校层次 ----
    if not opts.get("no-tier-check"):
        tok, tlines = validate_tiers(data, warn=warn_wrap)
        print("-" * 74)
        print("院校层次自检（铁律：985 / 211 / 双一流 / 中科院特例，每档至少 2 所）")
        for ln in tlines:
            print("  " + ln)
        if not tok:
            print("  ==> 自检未通过！请回去调整 school_tiers 后再生成。")
            if not opts.get("force"):
                print("      （如需强行生成，加 --force）")
                sys.exit(2)
        else:
            print("  ==> 自检通过")
    print("-" * 74)

    # 版心兜底：避开左侧导航
    # 2026-10-12 修：只认「完全位于左侧导航区内」的形状。旧逻辑只要 left < 25% 页宽
    # 就算导航，而页面标题文本框左缘虽在 25% 内、右缘却远超（如 2.83in 起 6.67in 宽），
    # 会被误判成导航，把 nav_right 抬到 9.5in，进而触发下方「识别不准」兜底分支，
    # 表格被放到 0.17×页宽（2.27in）压住导航条（山东大学实例）。同时把无文字的
    # 导航高亮条（矩形）也计入，让表格与导航区之间按 STD_GAP_AFTER_NAV 留距。
    nav_right = 0
    for slide in list(prs.slides)[:8]:
        for s in slide.shapes:
            l = s.left or 0
            r = l + (s.width or 0)
            if l >= sw * 0.25 or r > sw * 0.25:
                continue
            if s.has_text_frame and s.text_frame.text.strip():
                nav_right = max(nav_right, r)
            elif s.shape_type in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.PICTURE):
                nav_right = max(nav_right, r)
    # 标准版心（2026-10-12 用户定版）：左缘=导航右缘+0.43in，无导航按页宽比例取；
    # 顶部=页高x0.1627（避开页标题），右侧距 0.33in、底部距 0.56in
    fb_left = (int(nav_right + STD_GAP_AFTER_NAV * EMU_PER_IN) if nav_right
               else int(sw * STD_LEFT_RATIO))
    fb_top = int(sh * STD_TOP_RATIO)
    fb_width = int(sw - fb_left - STD_RIGHT_MARGIN * EMU_PER_IN)
    fb_height = int(sh - fb_top - STD_BOTTOM_MARGIN * EMU_PER_IN)
    fallback = (fb_left, fb_top, fb_width, fb_height)
    # 版心宽度不能小于内容区的一半，否则说明导航识别异常，退回按页宽比例的标准左缘
    if fallback[2] < sw * 0.5:
        fb_left = int(sw * STD_LEFT_RATIO)
        fallback = (fb_left, fb_top,
                    int(sw - fb_left - STD_RIGHT_MARGIN * EMU_PER_IN), fb_height)

    template = None
    used = set()
    table_pages = set()
    done = 0
    trim = bool(opts.get("ppt-trim"))
    for k, spec in enumerate(SPECS):
        rows = get_rows(data, spec, True, trim=trim)
        if not rows:
            print("  [%s] 数据为空，跳过" % spec["title"])
            continue

        slide, created = None, False
        if explicit and k < len(explicit):
            idx = explicit[k] - 1
            if 0 <= idx < len(prs.slides):
                slide = prs.slides[idx]
        elif append_mode:
            if template is None:
                template = pick_template(prs)
            slide = duplicate_slide(prs, template)
            clear_content(slide)
            set_slide_title(slide, spec["sub"] if opts.get("long-title") else spec["title"])
            move_slide_to_end(prs, slide)
            if after_n:  # 插到第 N 页之后
                lst = prs.slides._sldIdLst
                lst.remove(slide._element if False else None)
            created = True
        else:
            best_i, best_score = None, 0
            for i2, s2 in enumerate(prs.slides):
                if i2 in used:
                    continue
                t = slide_text(s2)
                score = max([len(kw) for kw in spec["match"] if kw in t] or [0])
                if score and (largest_picture(s2) is not None or any(x.has_table for x in s2.shapes)):
                    score += 1
                if score > best_score:
                    best_score, best_i = score, i2
            if best_i is not None:
                slide = prs.slides[best_i]
                used.add(best_i)
            else:
                if template is None:
                    template = pick_template(prs)
                slide = duplicate_slide(prs, template)
                clear_content(slide)
                set_slide_title(slide, spec["sub"] if opts.get("long-title") else spec["title"])
                move_slide_to_end(prs, slide)
                created = True

        if slide is None:
            print("  [%s] 无可用页面，跳过" % spec["title"])
            continue

        box, pic = compute_box(prs, slide, fallback)
        if pic is not None:
            remove_shape(slide, pic)
        info = build_table(slide, spec, rows, box, theme, max_font, min_font,
                           font_scale, fill_height=not opts.get("no-fill-height"),
                           header=header_row(data, spec))
        page_no = slide_index(prs, slide) + 1
        table_pages.add(page_no)
        flag = " ⚠溢出" if info["overflow"] else ""
        print("  [%s] 第 %d 页%s  版心 %.0f x %.0f px  字号 %.1f pt  表高 %.0f/%.0f px%s"
              % (spec["title"], page_no, "（新建）" if created else "",
                 emu2px(box[2]), emu2px(box[3]), info["font_pt"],
                 info["table_px"], info["avail_px"], flag))
        print("        行数列宽(px): %s" % info["widths"])
        for wmsg in info["warnings"]:
            print("        ⚠ %s" % wmsg)
        done += 1

    # ---- 没投到表格的空内容页：默认删掉（三表制下模板会多出「各阶段重点任务分析」页）----
    if orig_holes and not opts.get("keep-empty-pages"):
        drop = [h for h in orig_holes if h not in table_pages]
        if drop:
            titles = []
            for h in drop:
                t = ""
                try:
                    for shp in prs.slides[h - 1].shapes:
                        if shp.has_text_frame and (shp.left or 0) >= prs.slide_width * 0.15:
                            tt = shp.text_frame.text.strip().replace("\n", "/")
                            if tt and len(tt) <= 20:
                                t = tt
                                break
                except Exception:
                    pass
                titles.append("第%d页%s" % (h, "（%s）" % t if t else ""))
            print("-" * 74)
            print("已移除未使用的空内容页: %s" % "、".join(titles))
            if not dry_run:
                delete_slides(prs, drop)
    elif orig_holes and opts.get("keep-empty-pages"):
        left = [h for h in orig_holes if h not in table_pages]
        if left:
            print("-" * 74)
            print("提示：第 %s 页是空内容页，本次未投放表格（--keep-empty-pages 已保留）"
                  % "、".join(str(x) for x in left))

    if dry_run:
        print("\n(dry-run，未写出文件)  → 将输出到: %s" % os.path.abspath(out_path))
        return
    out_dir = os.path.dirname(os.path.abspath(out_path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    prs.save(out_path)
    print("-" * 74)
    print("已生成  : %s" % os.path.abspath(out_path))
    print("完成 %d/%d 张表，输出共 %d 页" % (done, want, len(prs.slides)))


if __name__ == "__main__":
    main()

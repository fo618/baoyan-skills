# -*- coding: utf-8 -*-
"""
保研规划两表生成器（2026-09-19 定版）

根据学员数据 JSON 生成一份包含 2 个工作表的 Excel：
  1. 学员信息表      —— 2 列（学员情况标签 | 内容），表头行右格写学员摘要，
                        固定 4 行：基本情况 / 保研数据 / 保研方向 / 注意事项
                        （「院校定位」行已于 2026-09-19 删除，JSON 里写了也会被过滤）
  2. 冲稳保定档表    —— 5 列，每档 4 所，首列"梯度"同档位纵向合并

变更记录：主线任务表已于 2026-09-19 下线（PPT / Excel 均不再生成）；
JSON 里的 main_timeline 字段保留但不再渲染。

主题：深红表头 + 浅红首列 + 细红边框，两张表风格统一。
排版：列宽固化、行高按实际换行行数精算、打印区域与缩放已设置，生成即可截图/导出，无需手动调行高列宽。

用法:
  python generate_tables.py <data.json> [output.xlsx] [--no-split]

选项:
  --no-split   关掉"1. …；2. …"自动转换行编号清单

data.json 结构见 assets/sample_data.json。
"""
import json
import math
import re
import sys
import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.page import PageMargins

# ---------- 主题常量（所有输出保持一致，勿随意改动） ----------
HEADER_BG = "9E2F2F"      # 深红表头
HEADER_FONT = "FFFFFF"    # 白色字
FIRSTCOL_BG = "F7E4E1"    # 浅红首列
BODY_BG = "FFFFFF"
BORDER_COLOR = "C97B7B"
FIRSTCOL_FONT_COLOR = "7A2222"

FONT_NAME = "微软雅黑"
BODY_SIZE = 11            # 正文字号
HEADER_SIZE = 11          # 表头字号

# ---------- 排版常量（行高精算用） ----------
# Excel 列宽 1 单位 ≈ 7px（Calibri 11 的 '0' 宽度）；1pt = 4/3 px
PX_PER_UNIT = 7.0
PT_TO_PX = 4.0 / 3.0
LATIN_WIDTH_RATIO = 0.58            # 半角字符宽度 ≈ 字号的 0.58 倍
LINE_HEIGHT = BODY_SIZE * 1.34 + 2  # 单行行高（pt）
ROW_PADDING = 9                     # 单元格上下留白（pt）
SAFETY = 1.05                       # 行高余量，防止不同机器字体度量差异导致截断
COL_PADDING = 1.2                   # 列宽两侧留白（字符单位）

CJK_UNITS = (BODY_SIZE * PT_TO_PX) / PX_PER_UNIT          # 全角字符占用的列宽单位
LATIN_UNITS = (BODY_SIZE * LATIN_WIDTH_RATIO * PT_TO_PX) / PX_PER_UNIT

HEADER_FILL = PatternFill("solid", fgColor=HEADER_BG)
FIRSTCOL_FILL = PatternFill("solid", fgColor=FIRSTCOL_BG)
BODY_FILL = PatternFill("solid", fgColor=BODY_BG)

THIN = Side(style="thin", color=BORDER_COLOR)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

HEADER_FONT_OBJ = Font(name=FONT_NAME, size=HEADER_SIZE, bold=True, color=HEADER_FONT)
BODY_FONT_OBJ = Font(name=FONT_NAME, size=BODY_SIZE, color="333333")
FIRSTCOL_FONT_OBJ = Font(name=FONT_NAME, size=BODY_SIZE, bold=True, color=FIRSTCOL_FONT_COLOR)

HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
# 正文默认：横向靠左 + 纵向居中（长文本单元格）
BODY_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)
# 短标签列（梯度 / 时间）：横向居中 + 纵向居中
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

# "1. …；2. …；3. …" -> 换行编号清单（截图可读性显著提升）
_NUM_SPLIT = re.compile(r"([；;])\s*(?=(?:\d{1,2})[.、])")


def pretty_numbered(text):
    """把分号分隔的编号条目转成换行列表，长文本截图时更易读。"""
    if not isinstance(text, str) or len(text) < 30:
        return text
    return _NUM_SPLIT.sub(r"\1\n", text)


def est_lines(text, width_chars):
    """按列宽与字体度量估算换行后的行数。"""
    if text is None:
        return 1
    usable = max(4.0, width_chars - COL_PADDING)
    lines = 0
    for seg in str(text).split("\n"):
        w = 0.0
        for c in seg:
            w += CJK_UNITS if ord(c) > 127 else LATIN_UNITS
        lines += max(1, math.ceil(w / usable))
    return lines


def setup_print(ws, ncols, last_row):
    """横向 A4、缩放至一页宽、锁定打印区域、逐页重复表头——便于直接导出 PDF / 整表截图。"""
    last_col = get_column_letter(ncols)
    ws.print_area = "A1:%s%d" % (last_col, last_row)
    ws.print_title_rows = "1:1"          # 表头只有第 1 行（不再有独立标题行）
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9  # A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.5, bottom=0.5,
                                  header=0.2, footer=0.2)


def build_sheet(ws, headers, widths, rows, center_cols=(), auto_merge_first=False,
                split_numbered=True):
    """通用建表：深红表头、浅红首列、统一边框、自适应行高、首列同值纵向合并、打印设置。

    headers          表头文本列表（第 1 行）
    widths           各列列宽（字符单位）
    rows             数据行（list of tuple）
    center_cols      需要"横向居中"的列序号（0 基），一般给短标签列（梯度 / 时间）
    auto_merge_first 首列相邻同值单元格纵向合并（定档表的"梯度"列用）
    """
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    ncols = len(headers)

    # ---- 表头（第 1 行）----
    hdr_lines = 1
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=j, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT_OBJ
        cell.alignment = HEADER_ALIGN
        cell.border = BORDER
        hdr_lines = max(hdr_lines, est_lines(h, widths[j - 1]))
    ws.row_dimensions[1].height = round((hdr_lines * LINE_HEIGHT + ROW_PADDING) * SAFETY, 1)

    # ---- 列宽 ----
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

    # ---- 数据行 ----
    # 先逐格设好样式再合并：openpyxl 中被合并掉的 MergedCell 不能再写样式
    last_row = 1
    for i, row in enumerate(rows, start=2):
        last_row = i
        max_lines = 1
        for j, val in enumerate(row, start=1):
            if j > 1 and split_numbered:
                val = pretty_numbered(val)
            cell = ws.cell(row=i, column=j, value=val)
            cell.border = BORDER
            max_lines = max(max_lines, est_lines(val, widths[j - 1]))
            if j == 1:
                cell.fill = FIRSTCOL_FILL
                cell.font = FIRSTCOL_FONT_OBJ
            else:
                cell.fill = BODY_FILL
                cell.font = BODY_FONT_OBJ
            cell.alignment = CENTER_ALIGN if (j - 1) in center_cols else BODY_ALIGN
        ws.row_dimensions[i].height = round((max_lines * LINE_HEIGHT + ROW_PADDING) * SAFETY, 1)

    # ---- 首列同值纵向合并（如"冲刺档"占 4 行）----
    if auto_merge_first and rows:
        # 数据第 i 行（0 基）= Excel 第 i+2 行；连续同值的块 [a, b] -> Excel 行 [a+2, b+2]
        start = 0
        prev = rows[0][0]
        for idx in range(1, len(rows) + 1):
            cur = rows[idx][0] if idx < len(rows) else None
            if cur != prev:
                r1, r2 = start + 2, idx + 1
                if r2 - r1 >= 1:           # 至少两行才值得合并
                    ws.merge_cells(start_row=r1, start_column=1, end_row=r2, end_column=1)
                start = idx
                prev = cur

    ws.freeze_panes = "A2"
    setup_print(ws, ncols, last_row)


# ---------- 三张表的列定义（列名 / 列宽 / 居中列）----------
TIER_HEADERS = ["梯度", "适配院校", "核心适配方向", "匹配你的核心优势", "录取大致所需条件"]
TIER_WIDTHS = [10, 22, 26, 33, 49]

TIMELINE_HEADERS = ["时间", "成绩", "英语", "竞赛", "科研 & 项目", "其他"]
TIMELINE_WIDTHS = [19, 32, 27, 23, 24, 20]


def build_all(data, split_numbered=True):
    wb = Workbook()
    s = data

    # ---- 1. 学员信息表（2 列；表头行第二格是学员摘要）----
    # 2026-09-19 用户定版：删掉「院校定位」行（三档院校以定档表为准），固定 4 行：
    # 基本情况 / 保研数据 / 保研方向 / 注意事项。JSON 里就算写了也会被过滤。
    si = s["student_info"]
    ws1 = wb.active
    ws1.title = "学员信息表"
    rows1 = [(r["label"], r["content"]) for r in si["rows"]
             if "院校定位" not in str(r.get("label", ""))]
    build_sheet(ws1, ["学员情况", si["headline"]], [12, 118], rows1,
                split_numbered=split_numbered)

    # ---- 2. 冲稳保定档表（5 列，每档 4 所，梯度列纵向合并）----
    ws2 = wb.create_sheet("冲稳保定档表")
    rows2 = [(r["tier"], r["school"], r["direction"], r["fit"], r["requirement"])
             for r in s["school_tiers"]]
    build_sheet(ws2, TIER_HEADERS, TIER_WIDTHS, rows2,
                center_cols=(0,), auto_merge_first=True, split_numbered=split_numbered)

    # ---- 主线任务表已于 2026-09-19 下线（PPT / Excel 均不再生成）----
    # rows3 = [(r["time"], r["score"], r["english"], r["competition"], r["research"], r["other"])
    #          for r in s["main_timeline"]]
    # build_sheet(ws3, TIMELINE_HEADERS, TIMELINE_WIDTHS, rows3,
    #             center_cols=(0,), split_numbered=split_numbered)

    return wb


# 旧版（四表）字段 —— 命中就给明确提示，避免拿旧 JSON 静默出错
_LEGACY_KEYS = ("student_profile", "main_tasks", "stage_tasks")


def _check_schema(data):
    legacy = [k for k in _LEGACY_KEYS if k in data]
    if legacy:
        print("× 这份 JSON 是旧版四表结构（发现字段：%s）。" % "、".join(legacy))
        print("  新版只有三张表：student_info / school_tiers / main_timeline。")
        print("  请对照 assets/sample_data.json 重新生成数据源。")
        sys.exit(1)
    need = ("student_info", "school_tiers", "main_timeline")
    missing = [k for k in need if k not in data]
    if missing:
        print("× JSON 缺少必需字段：%s（对照 assets/sample_data.json）" % "、".join(missing))
        sys.exit(1)


def main():
    args = sys.argv[1:]
    split_numbered = "--no-split" not in args
    args = [a for a in args if not a.startswith("--")]
    if len(args) < 1:
        print("用法: python generate_tables.py <data.json> [output.xlsx] [--no-split]")
        sys.exit(1)
    data_path = args[0]
    out_path = args[1] if len(args) > 1 else "保研规划三表.xlsx"

    with open(data_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    _check_schema(data)

    wb = build_all(data, split_numbered)
    out_dir = os.path.dirname(os.path.abspath(out_path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    wb.save(out_path)
    print("已生成:", out_path, "（2 个工作表：学员信息表 / 冲稳保定档表；主线任务表已于 2026-09-19 下线）")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
微信聊天记录捕获工具（WeChat 4.x / Weixin.exe 适配版）
=======================================================

背景：微信 4.x（Weixin.exe）主窗口是 Qt 自绘渲染（MMUIRenderSubWindow），
UI Automation 完全读不到聊天文本（实测只有 4 个空壳节点）。
因此采用「剪贴板捕获」方案：

    你在微信里选中消息 → Ctrl+C → 本工具读取剪贴板并解析入库

同时本工具会从微信窗口标题读取「当前聊天对象」，自动给记录归到对应联系人。

用法：
    python wechat_capture.py                 # 读取一次剪贴板并入库
    python wechat_capture.py --watch 120      # 持续监听 120 秒（可多次复制）
    python wechat_capture.py --list           # 列出已保存的会话
    python wechat_capture.py --show 张同学     # 打印某个会话
    python wechat_capture.py --raw            # 只打印剪贴板原文（排查格式用）

数据目录默认 ./微信聊天记录/
"""

import argparse
import ctypes
import ctypes.wintypes as wt
import hashlib
import json
import os
import re
import sys
import time
from datetime import datetime

# ----------------------------------------------------------------------------
# 路径
# ----------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(os.path.dirname(BASE_DIR), "微信聊天记录")
JSONL_PATH = os.path.join(DATA_DIR, "capture.jsonl")

# 自己的昵称（微信复制时自己的消息会显示昵称；"我" 是常见占位）
SELF_NAMES = {"我", "self", "me", "我自己"}

# ----------------------------------------------------------------------------
# Win32 剪贴板 / 窗口
# ----------------------------------------------------------------------------
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
CF_UNICODETEXT = 13

user32.IsClipboardFormatAvailable.argtypes = [wt.UINT]
user32.IsClipboardFormatAvailable.restype = wt.BOOL
user32.OpenClipboard.argtypes = [wt.HWND]
user32.OpenClipboard.restype = wt.BOOL
user32.CloseClipboard.restype = wt.BOOL
user32.GetClipboardData.argtypes = [wt.UINT]
user32.GetClipboardData.restype = wt.HANDLE
kernel32.GlobalLock.argtypes = [wt.HGLOBAL]
kernel32.GlobalLock.restype = wt.LPVOID
kernel32.GlobalUnlock.argtypes = [wt.HGLOBAL]
kernel32.GlobalUnlock.restype = wt.BOOL

WNDENUMPROC = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)
user32.EnumWindows.argtypes = [WNDENUMPROC, wt.LPARAM]
user32.GetWindowTextW.argtypes = [wt.HWND, ctypes.c_wchar_p, ctypes.c_int]
user32.GetClassNameW.argtypes = [wt.HWND, ctypes.c_wchar_p, ctypes.c_int]
user32.IsWindowVisible.argtypes = [wt.HWND]
user32.GetWindowThreadProcessId.argtypes = [wt.HWND, ctypes.POINTER(wt.DWORD)]

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
kernel32.OpenProcess.argtypes = [wt.DWORD, wt.BOOL, wt.DWORD]
kernel32.OpenProcess.restype = wt.HANDLE
kernel32.QueryFullProcessImageNameW.argtypes = [wt.HANDLE, wt.DWORD, ctypes.c_wchar_p, ctypes.POINTER(wt.DWORD)]
kernel32.QueryFullProcessImageNameW.restype = wt.BOOL
kernel32.CloseHandle.argtypes = [wt.HANDLE]


def read_clipboard_text():
    """读取剪贴板 Unicode 文本。"""
    try:
        if not user32.IsClipboardFormatAvailable(CF_UNICODETEXT):
            return None
        if not user32.OpenClipboard(0):
            return None
        try:
            h = user32.GetClipboardData(CF_UNICODETEXT)
            if not h:
                return None
            ptr = kernel32.GlobalLock(h)
            if not ptr:
                return None
            try:
                return ctypes.wstring_at(ptr)
            finally:
                kernel32.GlobalUnlock(h)
        finally:
            user32.CloseClipboard()
    except Exception:
        return None


def _proc_name(pid):
    name = "?"
    h = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if h:
        buf = ctypes.create_unicode_buffer(1024)
        size = wt.DWORD(1024)
        if kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
            name = buf.value.split("\\")[-1]
        kernel32.CloseHandle(h)
    return name


def get_wechat_chat_title():
    """读取微信当前聊天窗口标题（即当前聊天对象名）。读不到返回 None。"""
    found = []

    def cb(hwnd, _):
        if not user32.IsWindowVisible(hwnd):
            return True
        pid = wt.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if _proc_name(pid.value).lower() not in ("weixin.exe", "wechat.exe"):
            return True
        cbuf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, cbuf, 256)
        if "QWindowIcon" not in cbuf.value:
            return True
        tbuf = ctypes.create_unicode_buffer(512)
        user32.GetWindowTextW(hwnd, tbuf, 512)
        if tbuf.value and tbuf.value.strip() not in ("微信", "Weixin", "WeChat"):
            found.append((hwnd, tbuf.value.strip()))
        return True

    try:
        user32.EnumWindows(WNDENUMPROC(cb), 0)
    except Exception:
        return None
    if not found:
        return None
    # 取面积/存在感最大的：简单取最后一个（通常为主窗口）
    return found[-1][1]


# ----------------------------------------------------------------------------
# 解析器：把微信复制出来的文本解析为结构化消息
# ----------------------------------------------------------------------------
# 支持的头部格式：
#   1) 2026-10-10 14:32:15 张三     （时间在前）
#   2) 张三 2026-10-10 14:32        （姓名在前）
#   3) 张三  14:32                  （仅时间）
TS = r"\d{4}[-/]\d{1,2}[-/]\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?"
TS_SHORT = r"\d{1,2}:\d{2}(?::\d{2})?"
P_TS_FIRST = re.compile(r"^(%s)\s+(.+)$" % TS)          # 时间 姓名
P_NAME_FIRST = re.compile(r"^(.+?)\s+(%s)\s*$" % TS)     # 姓名 时间
P_SHORT_TS_FIRST = re.compile(r"^(%s)\s+(.+)$" % TS_SHORT)
P_NAME_SHORT = re.compile(r"^(.+?)\s+(%s)\s*$" % TS_SHORT)

SYSTEM_HINTS = ("撤回了一条消息", "你已添加了", "以上是打招呼的内容",
                "以下是新消息", "查看更多消息", "拍了拍", "群公告",
                "你领取了", "对方正在输入", "以下为新消息")


def _clean_sender(s):
    return s.strip().strip("：:").strip()


def _looks_like_time_only(line):
    return re.fullmatch(r"\d{1,2}:\d{2}(?::\d{2})?", line.strip()) is not None


def parse_messages(text):
    """解析微信复制文本 -> [{'sender','ts','content'}]"""
    if not text or len(text.strip()) < 2:
        return []
    lines = [l.rstrip("\n") for l in text.splitlines()]
    msgs = []
    cur = None

    def flush():
        nonlocal cur
        if cur and cur["content"].strip():
            msgs.append(cur)
        cur = None

    for raw in lines:
        line = raw.strip()
        if not line:
            continue
        if any(h in line for h in SYSTEM_HINTS):
            continue

        m = P_TS_FIRST.match(line)
        if m:
            flush()
            cur = {"ts": m.group(1).strip(), "sender": _clean_sender(m.group(2)),
                   "content": ""}
            continue
        m = P_NAME_FIRST.match(line)
        if m and not m.group(1).strip().isdigit():
            flush()
            cur = {"ts": m.group(2).strip(), "sender": _clean_sender(m.group(1)),
                   "content": ""}
            continue
        m = P_SHORT_TS_FIRST.match(line)
        if m and _looks_like_time_only(m.group(1)):
            flush()
            cur = {"ts": m.group(1).strip(), "sender": _clean_sender(m.group(2)),
                   "content": ""}
            continue
        m = P_NAME_SHORT.match(line)
        if m and _looks_like_time_only(m.group(2)):
            flush()
            cur = {"ts": m.group(2).strip(), "sender": _clean_sender(m.group(1)),
                   "content": ""}
            continue

        # 不是头部 -> 作为内容（或孤立消息）
        if cur is None:
            cur = {"ts": "", "sender": "", "content": ""}
        cur["content"] = (cur["content"] + "\n" + line) if cur["content"] else line

    flush()

    # 兜底：整段没有任何头部 -> 视作一条纯文本消息
    if not msgs and text.strip():
        msgs = [{"ts": "", "sender": "", "content": text.strip()}]

    return msgs


def detect_role(sender, chat_title):
    """判断消息是自己发的还是对方发的。"""
    if not sender:
        return "unknown"
    if sender in SELF_NAMES:
        return "self"
    if chat_title:
        # 一对一：发送者==聊天对象 -> 对方；否则视为自己
        if _norm(sender) == _norm(chat_title):
            return "other"
        return "self"
    return "other"


def _norm(s):
    return re.sub(r"[\s（）()【】\[\]]", "", s or "")


# ----------------------------------------------------------------------------
# 存储
# ----------------------------------------------------------------------------
def _key(chat, sender, ts, content):
    return hashlib.md5(f"{chat}|{sender}|{ts}|{content}".encode("utf-8")).hexdigest()


def load_keys():
    keys = set()
    if os.path.exists(JSONL_PATH):
        with open(JSONL_PATH, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                    keys.add(rec.get("key", ""))
                except Exception:
                    continue
    return keys


def save_records(records):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(JSONL_PATH, "a", encoding="utf-8") as f:
        for r in records:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")

    # 同步生成可读 txt（按会话分文件）
    by_chat = {}
    for r in records:
        by_chat.setdefault(r["chat"] or "未命名会话", []).append(r)
    for chat, recs in by_chat.items():
        safe = re.sub(r'[\\/:*?"<>|]', "_", chat)
        path = os.path.join(DATA_DIR, f"{safe}.txt")
        with open(path, "a", encoding="utf-8") as f:
            for r in recs:
                ts = f"[{r['ts']}] " if r["ts"] else ""
                who = r["sender"] or "?"
                f.write(f"{ts}{who}: {r['content']}\n")


def capture_once(chat_title=None, verbose=True):
    text = read_clipboard_text()
    if not text or not text.strip():
        if verbose:
            print("× 剪贴板为空或无文本。请在微信里选中消息后 Ctrl+C，再运行本工具。")
        return 0

    chat = chat_title or get_wechat_chat_title() or "未命名会话"
    parsed = parse_messages(text)

    # 若解析出的消息都没有 sender，把这些内容当作"当前聊天对象"发来的？
    # 保守起见：无 sender 时标为 unknown，不臆测。
    seen = load_keys()
    new_records = []
    now = datetime.now().isoformat(timespec="seconds")
    for m in parsed:
        sender = m["sender"]
        role = detect_role(sender, chat)
        k = _key(chat, sender, m["ts"], m["content"])
        if k in seen:
            continue
        seen.add(k)
        new_records.append({
            "chat": chat, "sender": sender, "role": role,
            "ts": m["ts"], "content": m["content"],
            "captured_at": now, "key": k,
        })

    if new_records:
        save_records(new_records)

    if verbose:
        print(f"当前聊天对象：{chat}")
        print(f"剪贴板解析出 {len(parsed)} 条，新增入库 {len(new_records)} 条"
              f"（重复 {len(parsed) - len(new_records)} 条已跳过）")
        for r in new_records[:8]:
            tag = {"self": "我", "other": "对方", "unknown": "?"}[r["role"]]
            print(f"  · [{r['ts'] or '-'}] {r['sender'] or tag}: {r['content'][:40]}")
        if len(new_records) > 8:
            print(f"  ... 其余 {len(new_records) - 8} 条已入库")
        if new_records:
            print(f"→ 已保存到：{DATA_DIR}")
    return len(new_records)


def watch(seconds):
    print(f"▶ 开始监听剪贴板 {seconds} 秒。请到微信里选中消息并 Ctrl+C（可多次复制）。")
    last = None
    total = 0
    end = time.time() + seconds
    try:
        while time.time() < end:
            t = read_clipboard_text()
            if t and t != last:
                last = t
                n = capture_once(verbose=True)
                total += n
                print("-" * 50)
            time.sleep(0.5)
    except KeyboardInterrupt:
        pass
    print(f"■ 监听结束，本次共新增入库 {total} 条。")


def list_chats():
    if not os.path.exists(JSONL_PATH):
        print("还没有任何记录。")
        return
    stat = {}
    with open(JSONL_PATH, "r", encoding="utf-8") as f:
        for line in f:
            try:
                r = json.loads(line)
            except Exception:
                continue
            c = r["chat"]
            s = stat.setdefault(c, {"n": 0, "self": 0, "other": 0, "last": ""})
            s["n"] += 1
            if r["role"] == "self":
                s["self"] += 1
            elif r["role"] == "other":
                s["other"] += 1
            s["last"] = max(s["last"], r["captured_at"])
    print(f"{'会话':<20}{'条数':>6}{'我':>6}{'对方':>6}  最后记录")
    for c, s in sorted(stat.items(), key=lambda x: -x[1]["n"]):
        print(f"{c:<20}{s['n']:>6}{s['self']:>6}{s['other']:>6}  {s['last']}")


def show_chat(name):
    if not os.path.exists(JSONL_PATH):
        print("还没有任何记录。")
        return
    n = 0
    with open(JSONL_PATH, "r", encoding="utf-8") as f:
        for line in f:
            try:
                r = json.loads(line)
            except Exception:
                continue
            if name in r["chat"]:
                ts = f"[{r['ts']}] " if r["ts"] else ""
                print(f"{ts}{r['sender'] or '?'}: {r['content']}")
                n += 1
    print(f"--- 共 {n} 条 ---")


def main():
    ap = argparse.ArgumentParser(description="微信聊天记录捕获工具（4.x 剪贴板方案）")
    ap.add_argument("--watch", type=int, metavar="SEC", help="持续监听 N 秒")
    ap.add_argument("--raw", action="store_true", help="只打印剪贴板原文")
    ap.add_argument("--list", action="store_true", help="列出已保存会话")
    ap.add_argument("--show", metavar="NAME", help="打印某会话内容")
    ap.add_argument("--chat-name", help="手动指定聊天对象名")
    args = ap.parse_args()

    if args.raw:
        t = read_clipboard_text()
        print(f"剪贴板长度: {len(t) if t else 0}")
        print("当前微信聊天窗口:", get_wechat_chat_title())
        print("-" * 60)
        print(t or "(空)")
        return
    if args.list:
        list_chats()
        return
    if args.show:
        show_chat(args.show)
        return
    if args.watch:
        watch(args.watch)
        return
    capture_once(chat_title=args.chat_name, verbose=True)


if __name__ == "__main__":
    main()

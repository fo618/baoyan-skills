# -*- coding: utf-8 -*-
r"""核对文章要引用的关键数字 —— 修正版。
关键修正：成交周期 = 首次接触 -> 首次「成交动作」（不是首条->末条，那含售后期）。"""
import collections
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import wxlib as W  # noqa: E402

sys.stdout.reconfigure(encoding='utf-8')
cfg = W.load_config()
contacts, self_wxid = W.read_contacts(cfg)

# --- 1) 「已成交」标签命中（直接按 label_id=42 数，避免名字误差）---
c42 = [c for c in contacts.values() if 42 in (c.get('label_ids') or [])]
cname = [c for c in contacts.values() if '已成交' in (c.get('labels') or [])]
all_lab = collections.Counter()
for c in contacts.values():
    for t in c['labels']:
        all_lab[t] += 1
print('[1] label_id=42 命中 = %d ；标签名「已成交」命中 = %d' % (len(c42), len(cname)))
print('    被真正用上的标签种类 = %d' % len(all_lab))
print('    Top8:', all_lab.most_common(8))

# --- 2) 成交动作 & 成交周期 ---
DEAL_ACT = ['合同', '协议', 'e签宝', 'esign', '甲方信息', '对公', '收款账户', '定金', '签署']
chats = W.load_chats(cfg, dirs=[cfg['live_dir'], cfg['deal_dir']])
spans, first_contact, act_lag, detail = [], [], [], []
done = set()
for wxid, ch in chats.items():
    ct = contacts.get(wxid) or {}
    if '已成交' not in (ct.get('labels') or []) or wxid in done:
        continue
    msgs = [m for m in ch['msgs'] if m.get('ts')]
    if not msgs:
        continue
    done.add(wxid)
    t0 = W.ts_to_dt(msgs[0]['ts'])
    t1 = W.ts_to_dt(msgs[-1]['ts'])
    if t0 and t1:
        spans.append((t1 - t0).days)
    # 我方首次发出成交动作的时间
    hit = None
    for m in msgs:
        if not m.get('is_self'):
            continue
        c = m.get('content') or ''
        if any(k in c for k in DEAL_ACT):
            hit = W.ts_to_dt(m['ts'])
            break
    if t0 and hit:
        act_lag.append((hit - t0).days)
        detail.append((ct.get('remark') or '', (hit - t0).days))

def pct(arr, lo, hi=None):
    n = len(arr)
    if not n:
        return '—'
    k = sum(1 for x in arr if (x <= lo if hi is None else lo <= x <= hi))
    return '%d (%.0f%%)' % (k, k / n * 100)

for nm, arr in [('关系总时长(首->末)', spans), ('成交周期(首->首次成交动作)', act_lag)]:
    a = sorted(arr)
    if not a:
        continue
    print('\n[2] %s  样本 %d' % (nm, len(a)))
    print('    中位数 %d 天 · 均值 %.0f 天 · 最短 %d · 最长 %d' % (a[len(a)//2], sum(a)/len(a), a[0], a[-1]))
    print('    ≤7 天: %s   |  8-30 天: %s   |  >30 天: %s'
          % (pct(arr, 7), pct(arr, 8, 30), pct(arr, 31, 99999)))

# --- 3) 成交客户里在聊天中提到退费诉求的 ---
REF_KW = ['退费', '退款', '退一下', '退钱', '要退', '退还']
ref_hit = []
for wxid, ch in chats.items():
    ct = contacts.get(wxid) or {}
    if '已成交' not in (ct.get('labels') or []):
        continue
    other = '\n'.join((m.get('content') or '') for m in ch['msgs'] if not m.get('is_self'))
    if sum(1 for k in REF_KW if k in other) >= 1:
        ref_hit.append(ct.get('remark') or wxid)
print('\n[3] 成交客户里对方提过退费/退款诉求的 = %d 个' % len(ref_hit))
for r in ref_hit[:14]:
    print('    ·', r)

print('\n[4] 成交客户总数(去重) = %d' % len(done))

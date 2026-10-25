# -*- coding: utf-8 -*-
r"""
从成交会话里抽取：客户异议 → 我的应答 配对、逼单话术、风险逆转话术。
用于定期给话术库「补料」。

用法: python objection_mine.py [--dir <导出目录>] [--limit 30]
"""
import argparse
import collections
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import wxlib as W  # noqa: E402

OBJ = ['太贵', '有点贵', '贵', '优惠', '便宜', '打折', '预算', '再考虑', '考虑一下', '对比',
       '商量', '商量一下', '看看再说', '再看看', '价格', '费用高', '考虑考虑', '不急', '过几天',
       '回头', '再说吧', '犹豫']
CLOSE = ['名额', '最后', '截止', '尽快', '这几天', '月底', '抓紧', '锁定', '占个', '保留',
         '定下来', '敲定', '今天', '明天', '优惠', '申请', '确定', '回复我', '给我答复']
GUAR = ['退费', '退还', '不满意', '无违约金', '随时退出', '保障', '全额']


def load(cfg, d):
    src = os.path.join(d or cfg['deal_dir'], 'all_messages.jsonl')
    if not os.path.isfile(src):
        raise SystemExit('找不到 %s，请先跑 refresh.py --with-deals' % src)
    chats = collections.OrderedDict()
    for line in open(src, encoding='utf-8-sig'):
        line = line.strip()
        if not line:
            continue
        r = json.loads(line)
        chats.setdefault(r['chat'], []).append(r)
    return chats


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--dir')
    ap.add_argument('--limit', type=int, default=30)
    args = ap.parse_args()
    cfg = W.load_config()
    chats = load(cfg, args.dir)
    print('成交会话:', len(chats))

    print('\n' + '=' * 90)
    print('【A】客户异议 → 我的应答（成交会话）')
    print('=' * 90)
    seen = 0
    for name, msgs in chats.items():
        for i, m in enumerate(msgs):
            if m.get('is_self'):
                continue
            c = m.get('content') or ''
            if not any(w in c for w in OBJ) or not (4 <= len(c) <= 90):
                continue
            nxt = None
            for j in range(i + 1, min(i + 6, len(msgs))):
                if msgs[j].get('is_self') and len(msgs[j].get('content') or '') > 8:
                    nxt = msgs[j]
                    break
            if not nxt:
                continue
            seen += 1
            if seen > args.limit:
                break
            print('\n·「%s」   <%s>' % (c.strip()[:80], name[:16]))
            print('   → 我：%s' % (nxt.get('content') or '').replace('\n', ' / ')[:150])
        if seen > args.limit:
            break

    print('\n' + '=' * 90)
    print('【B】逼单/促单话术（我发出、含推进词）')
    print('=' * 90)
    samples = collections.defaultdict(list)
    for name, msgs in chats.items():
        for m in msgs:
            if not m.get('is_self'):
                continue
            c = (m.get('content') or '').strip()
            if not c or c.startswith('[') or not (8 <= len(c) <= 120):
                continue
            for w in CLOSE:
                if w in c:
                    samples[w].append((name, c))
                    break
    for w in CLOSE:
        lst = samples.get(w) or []
        if len(lst) < 2:
            continue
        print('\n— 含「%s」(%d 次) —' % (w, len(lst)))
        for name, c in lst[:3]:
            print('   · [%s] %s' % (name[:12], c.replace('\n', ' / ')[:110]))

    print('\n' + '=' * 90)
    print('【C】风险逆转/保障话术（我发出）')
    print('=' * 90)
    shown = set()
    for name, msgs in chats.items():
        for m in msgs:
            if not m.get('is_self'):
                continue
            c = m.get('content') or ''
            if any(w in c for w in GUAR) and 10 < len(c) < 260:
                k = c[:28]
                if k in shown:
                    continue
                shown.add(k)
                print('   · [%s] %s' % (name[:14], c.replace('\n', ' / ')[:190]))
                if len(shown) > 24:
                    break
        if len(shown) > 24:
            break


if __name__ == '__main__':
    main()

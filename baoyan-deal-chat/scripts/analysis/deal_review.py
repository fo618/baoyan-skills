# -*- coding: utf-8 -*-
r"""
成交客户周期性复盘（画像 / 节奏 / 证据覆盖 / 跨会话复用话术）。

数据源默认取 config 里的 deal_dir（成交客户完整历史）。
用法: python deal_review.py [--dir <导出目录>]
"""
import argparse
import collections
import datetime
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import wxlib as W  # noqa: E402

LEVEL = [
    ('985', ['985', '北京', '清华', '北大', '复旦', '上海交通', '浙江', '南京', '中国科学技术',
             '哈工', '西安交通', '中山', '华中科技', '武汉大学', '同济', '东南', '天津', '南开',
             '四川大学', '山东大学', '中南大学', '湖南大学', '电子科技', '北理', '北航', '中国农业',
             '大连理工', '东北大学', '吉林大学', '重庆大学', '兰州大学', '西北工业', '西北农林',
             '中海洋', '中国海洋', '中国人民', '北师大', '华东师范', '华南理工', '厦门大学']),
    ('211/双一流', ['211', '双一流', '北京交通', '北京邮电', '中央财经', '对外经贸', '上海财经',
                    '南京航空航天', '南京理工', '西安电子', '合肥工业', '西南交通', '华中师范',
                    '东北师范', '郑州大学', '苏州大学', '江南大学', '南昌大学', '华北电力',
                    '哈尔滨工程', '中国地质', '中国石油', '中南财经', '上海大学', '暨南',
                    '北京工业', '北京科技', '武汉理工', '中国矿业', '河海大学', '长安大学', '扬州大学']),
]

EVID = {
    '甲方信息表': ['甲方信息'],
    'e签宝/签合同': ['esign', 'e签宝', '合同签署链接', '签合同'],
    '对公收款账户': ['支付宝账户', '银行账户', '户名：', '收款码'],
    '转账凭证': ['[转账]', '收到转账'],
}


def level_of(name):
    for tag, kws in LEVEL:
        if any(k in name for k in kws):
            return tag
    if any(k in name for k in ['大学', '学院']):
        return '其他本科'
    return '未知'


_PUNCT = set('，。！？、；：（）【】[].,!?:;()~～"“”‘’·—')


def norm(t):
    t = re.sub(r'\s+', '', t)
    t = re.sub(r'\d+', '#', t)
    return ''.join(c for c in t if c not in _PUNCT)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--dir', help='导出目录（默认 config.deal_dir）')
    ap.add_argument('--top', type=int, default=25)
    args = ap.parse_args()

    cfg = W.load_config()
    d = args.dir or cfg['deal_dir']
    src = os.path.join(d, 'all_messages.jsonl')
    if not os.path.isfile(src):
        raise SystemExit('找不到 %s，请先跑 refresh.py --with-deals' % src)

    chats = collections.OrderedDict()
    for line in open(src, encoding='utf-8-sig'):
        line = line.strip()
        if not line:
            continue
        r = json.loads(line)
        chats.setdefault(r['chat'], []).append(r)
    print('成交会话数:', len(chats), ' 总消息:', sum(len(v) for v in chats.values()))

    def dt(s):
        try:
            return datetime.datetime.strptime(s[:10], '%Y-%m-%d')
        except Exception:
            return None

    lev = collections.Counter(level_of(k) for k in chats)
    print('\n=== 成交客户院校层次 ===')
    for k, v in lev.most_common():
        print('  %s: %d' % (k, v))

    tot = self_n = voice = 0
    for msgs in chats.values():
        tot += len(msgs)
        self_n += sum(1 for m in msgs if m.get('is_self'))
        voice += sum(1 for m in msgs if (m.get('content') or '').startswith('[语音'))
    print('\n=== 发言结构 ===')
    print('  我发出 %d (%.0f%%) | 对方 %d | 语音 %d | 平均每会话 %.0f 条'
          % (self_n, self_n / tot * 100, tot - self_n, voice, tot / len(chats)))

    print('\n=== 成交链证据覆盖 ===')
    for label, kws in EVID.items():
        hit = sum(1 for msgs in chats.values()
                  if any(any(w in (m.get('content') or '') for w in kws) for m in msgs))
        print('  %-14s %3d/%d 会话 (%.0f%%)' % (label, hit, len(chats), hit / len(chats) * 100))

    print('\n=== 加好友 → 首次签约动作 周期 ===')
    cyc = []
    for name, msgs in chats.items():
        first = dt(msgs[0]['ts'])
        deal_ts = None
        for m in msgs:
            if any(w in (m.get('content') or '') for w in ['esign', '甲方信息', '签合同需要']):
                deal_ts = dt(m['ts'])
                break
        if first and deal_ts:
            cyc.append((name, (deal_ts - first).days, len(msgs)))
    cyc.sort(key=lambda x: x[1])
    for name, days, n in cyc[:15]:
        print('  %-30s %4d 天   (%d 条)' % (name[:28], days, n))
    if cyc:
        ds = sorted(c[1] for c in cyc)
        print('  -- 样本 %d；中位数 %d 天；平均 %.0f 天' % (len(ds), ds[len(ds) // 2], sum(ds) / len(ds)))
        print('  -- ≤7 天: %d；≤30 天: %d；>30 天: %d'
              % (sum(1 for x in ds if x <= 7), sum(1 for x in ds if x <= 30),
                 sum(1 for x in ds if x > 30)))

    bucket = collections.defaultdict(lambda: {'chats': set(), 'raw': '', 'n': 0})
    for name, msgs in chats.items():
        for m in msgs:
            if not m.get('is_self'):
                continue
            c = m.get('content') or ''
            if len(c) < 25 or c.startswith('['):
                continue
            key = norm(c)[:60]
            if len(key) < 15:
                continue
            b = bucket[key]
            b['chats'].add(name)
            b['n'] += 1
            if len(c) > len(b['raw']):
                b['raw'] = c
    multi = sorted(((k, v) for k, v in bucket.items() if len(v['chats']) >= 5),
                   key=lambda x: -len(x[1]['chats']))
    print('\n=== 跨 ≥5 个成交会话复用的话术（共 %d 条）Top %d ===' % (len(multi), args.top))
    for _k, v in multi[:args.top]:
        print('[×%d] %s' % (len(v['chats']), v['raw'].replace('\n', ' / ')[:170]))


if __name__ == '__main__':
    main()

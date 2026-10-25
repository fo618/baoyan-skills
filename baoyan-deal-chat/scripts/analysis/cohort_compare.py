# -*- coding: utf-8 -*-
r"""
成交组 vs 未成交组 —— 同窗口对比（避免观测窗口不等长的伪相关）

为什么单独写这个脚本：
  成交客户的完整聊天历史天然比未成交客户长（前者聊成功了、后者可能早早流失）。
  如果拿「成交客户全历史」去比「未成交客户近 30 天」，动作率必然虚高，
  得出的「发合同率 100% vs 29%」有一半是观测窗口的功劳，不是行为差异。
  本脚本强制两组来自**同一批数据集、同一时间窗口**。

用法：
  python cohort_compare.py                 # 用 config.json 的 datasets
  python cohort_compare.py --days 180      # 只保留最近 180 天的消息（两组同裁）
  python cohort_compare.py --dir <单一数据集目录>
"""
import argparse
import collections
import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import wxlib as W  # noqa: E402

# ---- 成交动作词表（对手方消息里出现即算「我方做过这个动作」）----
ACTIONS = collections.OrderedDict([
    ('发合同/协议', ['合同', '协议', '甲方信息', 'e签宝', 'esign.cn', '签署', '电子签']),
    ('要甲方信息', ['甲方信息', '身份证', '实名认证', '开户信息']),
    ('发对公账户', ['对公', '收款账户', '账户名', '开户行', '收款码', '收款账号', '打款']),
    ('收定金/转账', ['定金', '定金已交', '已转账', '转账凭证', '缴费成功', '已付款']),
    ('给退费保障', ['退费', '不成功退', '没上岸', '保底', '风险保障', '承诺退']),
])


def act_flags(msgs):
    """只统计**我方发出**的消息里是否有该动作。"""
    mine = '\n'.join((m.get('content') or '') for m in msgs if m.get('is_self'))
    return {name: any(k in mine for k in kws) for name, kws in ACTIONS.items()}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--days', type=int, default=0, help='只保留最近 N 天（两组同裁）；0=不限')
    ap.add_argument('--dir', default='', help='只用这一个数据集目录（留空=config.json 的 datasets）')
    ap.add_argument('--min-msgs', type=int, default=20, help='「深度接触」门槛：双方消息数下限')
    ap.add_argument('--json', action='store_true')
    args = ap.parse_args()

    cfg = W.load_config()
    dirs = [args.dir] if args.dir else cfg['datasets']
    contacts, self_wxid = W.read_contacts(cfg)

    cutoff = ''
    if args.days > 0:
        cutoff = (datetime.datetime.now(W.TZ) - datetime.timedelta(days=args.days)).strftime('%Y-%m-%d %H:%M:%S')

    chats = W.load_chats(cfg, dirs=dirs)

    deal, lead, peer, internal, bot = [], [], 0, 0, 0
    for wxid, ch in chats.items():
        msgs = [m for m in ch['msgs'] if m.get('ts') and (not cutoff or m['ts'] >= cutoff)]
        if not msgs:
            continue
        ct = contacts.get(wxid) or {}
        name = ct.get('remark') or ct.get('nick') or ch['chat'] or wxid
        seg = W.segment(name, ct.get('labels'), msgs)
        if seg == '渠道·同行':
            peer += 1
            continue
        if seg == '内部同事':
            internal += 1
            continue
        if seg == '机器人':
            bot += 1
            continue
        kind = W.client_kind(msgs, contacts, wxid)
        row = (name, msgs, kind)
        (deal if W.is_deal(ct) else lead).append(row)

    def agg(rows):
        n = len(rows)
        if not n:
            return None
        cnt = collections.Counter()
        voice = me = tot = 0
        msgcnt = []
        for _name, msgs, _k in rows:
            for k, v in act_flags(msgs).items():
                if v:
                    cnt[k] += 1
            voice += sum(1 for m in msgs if (m.get('content') or '').startswith('[语音'))
            me += sum(1 for m in msgs if m.get('is_self'))
            tot += len(msgs)
            msgcnt.append(len(msgs))
        msgcnt.sort()
        return {
            'n': n,
            'actions': {k: round(cnt[k] / n * 100) for k in ACTIONS},
            'voice_avg': round(voice / n, 1),
            'me_pct': round(me / tot * 100) if tot else 0,
            'msgs_median': msgcnt[len(msgcnt) // 2],
            'msgs_total': tot,
        }

    deep = [r for r in lead if len(r[1]) >= args.min_msgs]
    out = {
        'window_days': args.days or 'ALL',
        'self_wxid': self_wxid,
        'deal': agg(deal),
        'lead_deep': agg(deep),
        'lead_all': agg(lead),
        'excluded': {'渠道·同行': peer, '内部同事': internal, '机器人': bot},
    }

    if args.json:
        import json
        print(json.dumps(out, ensure_ascii=False, indent=1))
        return

    print('=' * 74)
    print('成交组 vs 未成交组 · 同窗口对比     窗口=%s' % (('近 %d 天' % args.days) if args.days else '全历史'))
    print('数据集:', ', '.join(os.path.basename(d.rstrip('/\\')) for d in dirs))
    print('已剔除: 渠道·同行 %d · 内部同事 %d · 机器人 %d' % (peer, internal, bot))
    print('=' * 74)

    def col(r):
        return '—' if not r else '%d 个' % r['n']

    print('%-16s %14s %16s %14s' % ('指标', '成交组', '未成交·深度', '未成交·全部'))
    print('-' * 74)
    print('%-16s %14s %16s %14s' % ('会话数', col(out['deal']), col(out['lead_deep']), col(out['lead_all'])))
    print('%-16s %14s %16s %14s' % ('中位消息数',
          out['deal']['msgs_median'] if out['deal'] else '—',
          out['lead_deep']['msgs_median'] if out['lead_deep'] else '—',
          out['lead_all']['msgs_median'] if out['lead_all'] else '—'))
    print('%-16s %13s%% %15s%% %13s%%' % ('我发言占比',
          out['deal']['me_pct'] if out['deal'] else '—',
          out['lead_deep']['me_pct'] if out['lead_deep'] else '—',
          out['lead_all']['me_pct'] if out['lead_all'] else '—'))
    print('%-16s %14s %16s %14s' % ('平均语音条数',
          out['deal']['voice_avg'] if out['deal'] else '—',
          out['lead_deep']['voice_avg'] if out['lead_deep'] else '—',
          out['lead_all']['voice_avg'] if out['lead_all'] else '—'))
    print('-' * 74)
    print('%-16s %14s %16s %14s' % ('【成交动作率】', '', '', ''))
    for k in ACTIONS:
        a = out['deal']['actions'][k] if out['deal'] else '—'
        b = out['lead_deep']['actions'][k] if out['lead_deep'] else '—'
        c = out['lead_all']['actions'][k] if out['lead_all'] else '—'
        fmt = lambda v: ('%d%%' % v) if isinstance(v, int) else v
        print('%-16s %14s %16s %14s' % (k, fmt(a), fmt(b), fmt(c)))
    print('-' * 74)


if __name__ == '__main__':
    main()

# -*- coding: utf-8 -*-
r"""
聊单工作台：把所有在跟客户一次排出来，按「今天该先跟谁」排序。

用法:
  python pipeline.py                     # 控制台输出
  python pipeline.py --md out.md         # 另存 Markdown（供日报引用）
  python pipeline.py --min-msgs 6 --top 25
  python pipeline.py --json

排序依据 = 成交信号强度 + 风险等级 + 是否超跟进节奏 + 阶段推进度。
"""
import argparse
import datetime
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import wxlib as W  # noqa: E402
from copilot import STAGE_ACTION, POST_ACTION  # noqa: E402

STAGE_RANK = {'S0': 0, 'S1': 1, 'S2': 2, 'S3': 3, 'S4': 4, 'S5': 5, 'S6': 6}


def build(cfg, min_msgs=6):
    contacts, self_wxid = W.read_contacts(cfg)
    chats = W.load_chats(cfg)
    rows = []
    for wxid, ch in chats.items():
        msgs = ch['msgs']
        if len(msgs) < min_msgs:
            continue
        if wxid.endswith('@chatroom') or wxid.startswith('gh_'):
            continue
        ct = contacts.get(wxid) or {}
        name = ct.get('remark') or ct.get('nick') or ch['chat'] or wxid
        seg = W.segment(name, ct.get('labels'), msgs)
        if seg == '机器人':
            continue
        deal = W.is_deal(ct)
        refunded = W.is_refunded(ct)
        sil = W.silence_days(msgs)
        if sil is None:
            continue
        if sil > cfg.get('cold_silence_days', 21):
            continue                       # 超冷线，不进日常池
        kind = W.client_kind(msgs, contacts, wxid)
        if deal:
            stage, ev = W.post_sale_stage(msgs)
            if refunded:
                stage, ev = 'P3 风险处置', '已退费'
            act = POST_ACTION.get(stage, '')
        else:
            stage, ev = W.infer_stage(msgs)
            act = STAGE_ACTION.get(stage, '')
        sig = W.scan_signals(msgs)
        risk = W.scan_risks(msgs)
        cadence = cfg['cadence_days'].get(kind, 4)
        overdue = sil >= cadence
        srank = STAGE_RANK.get(stage[:2], 0)

        score = sig['score'] + risk['level'] * 4 + (5 if overdue else 0) \
            + min(sil, 30) * 0.2 + srank * 1.5
        if refunded:
            score += 10

        todos = []
        if refunded:
            todos.append('⚠ 已退费：走交付复盘 + 关系修复，别硬推新单')
        if risk['level'] >= 4:
            todos.append('风险优先：' + '、'.join(r[0] for r in risk['hits']))
        if deal:
            if sil >= 7:
                todos.append('老客户关怀：' + act)
        else:
            if overdue:
                todos.append('已静默 %d 天（阈值 %d）→ 今天触达' % (sil, cadence))
            if sig['score'] >= 8:
                todos.append('高意向 → ' + act)
            elif not overdue:
                todos.append(act or '按阶段正常推进')

        rows.append({
            'name': name, 'wxid': wxid, 'msgs': len(msgs), 'silence': sil, 'segment': seg,
            'kind': kind, 'stage': stage, 'stage_evidence': ev,
            'signals': [n for n, _w, _c in sig['hits']], 'signal_score': sig['score'],
            'risks': [n for n, _w, _c in risk['hits']], 'risk_level': risk['level'],
            'is_deal': deal, 'refunded': refunded, 'overdue': overdue,
            'cadence': cadence, 'score': round(score, 1), 'todos': todos,
        })
    rows.sort(key=lambda r: -r['score'])
    return rows, self_wxid


def to_md(rows, now):
    L = ['# 聊单工作台 · %s' % now.strftime('%Y-%m-%d %H:%M'), '']
    clients = [r for r in rows if r['segment'] == '客户']
    peers = [r for r in rows if r['segment'] == '渠道·同行']
    internals = [r for r in rows if r['segment'] == '内部同事']
    lead = [r for r in clients if not r['is_deal']]
    keep = [r for r in clients if r['is_deal']]
    L.append('客户池：在跟线索 **%d** 个 · 存量老客户 **%d** 个 · 其中超跟进节奏 **%d** 个 ｜ 已剔除：渠道/同行 %d、内部同事 %d'
             % (len(lead), len(keep), sum(1 for r in clients if r['overdue']),
                len(peers), len(internals)))
    L.append('')
    L.append('## 一、今日优先跟进（在跟线索，按优先级）')
    L.append('')
    L.append('| # | 客户 | 阶段 | 信号 | 风险 | 静默 | 类型 | 建议动作 |')
    L.append('| --- | --- | --- | --- | --- | --- | --- | --- |')
    for i, r in enumerate(lead[:25], 1):
        L.append('| %d | %s | %s | %s | %s | %d天 | %s | %s |' % (
            i, r['name'], r['stage'], ('、'.join(r['signals']) or '-'),
            ('★%d %s' % (r['risk_level'], '、'.join(r['risks']))) if r['risks'] else '-',
            r['silence'], r['kind'], '；'.join(r['todos']) or '正常推进'))
    L.append('')
    L.append('## 二、存量老客户（续费 / 转介绍 / 交付关怀）')
    L.append('')
    L.append('| 客户 | 静默 | 阶段 | 建议动作 |')
    L.append('| --- | --- | --- | --- |')
    for r in sorted(keep, key=lambda x: -x['silence'])[:30]:
        L.append('| %s | %d天 | %s | %s |' % (
            r['name'], r['silence'], r['stage'], '；'.join(r['todos']) or '常规维护'))
    L.append('')
    return '\n'.join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--min-msgs', type=int, default=6)
    ap.add_argument('--top', type=int, default=25)
    ap.add_argument('--md', help='写入 Markdown 文件')
    ap.add_argument('--json', action='store_true')
    ap.add_argument('--include-peers', action='store_true',
                    help='把渠道/同行也并进客户池（默认剔除）')
    args = ap.parse_args()

    cfg = W.load_config()
    rows, self_wxid = build(cfg, args.min_msgs)
    now = datetime.datetime.now(W.TZ)

    if args.json:
        print(json.dumps({'self_wxid': self_wxid, 'rows': rows}, ensure_ascii=False, indent=1))
        return

    clients = [r for r in rows if r['segment'] == '客户']
    peers = [r for r in rows if r['segment'] == '渠道·同行']
    internals = [r for r in rows if r['segment'] == '内部同事']
    if args.include_peers:
        clients = clients + peers
    lead = [r for r in clients if not r['is_deal']]
    keep = [r for r in clients if r['is_deal']]

    print('=' * 96)
    print('聊单工作台  %s' % now.strftime('%Y-%m-%d %H:%M'))
    print('客户池：在跟线索 %d · 老客户 %d · 超跟进节奏 %d   ｜   已剔除：渠道·同行 %d · 内部同事 %d'
          % (len(lead), len(keep), sum(1 for r in clients if r['overdue']),
             len(peers), len(internals)))
    print('=' * 96)
    print('%-3s %-30s %-14s %-26s %-9s %-6s %s'
          % ('#', '客户', '阶段', '信号', '风险', '静默', '类型'))
    print('-' * 96)
    for i, r in enumerate(lead[:args.top], 1):
        print('%-3d %-30s %-14s %-26s %-9s %-6s %s'
              % (i, r['name'][:28], r['stage'], ('、'.join(r['signals']) or '-')[:24],
                 ('★%d %s' % (r['risk_level'], '、'.join(r['risks']))) if r['risks'] else '-',
                 '%d天' % r['silence'], r['kind']))
    print('-' * 96)
    print('\n【今日必做 Top 8】')
    todo_rows = [x for x in clients if x['todos']][:8]
    if not todo_rows:
        print('   （无待办）')
    for i, r in enumerate(todo_rows, 1):
        print('%d. %s（%s，静默 %d 天）' % (i, r['name'], r['stage'], r['silence']))
        for t in r['todos']:
            print('     → ' + t)
    if keep:
        print('\n【存量老客户 · 静默最久 Top 10】')
        for r in sorted(keep, key=lambda x: -x['silence'])[:10]:
            print('   %-30s 静默 %d 天  %s' % (r['name'][:28], r['silence'], r['stage']))
    if peers and not args.include_peers:
        print('\n【渠道·同行（走转介绍/对账，不属于客户聊单池）Top 10】')
        for r in peers[:10]:
            print('   %-30s 静默 %d 天  %s' % (r['name'][:28], r['silence'], r['stage']))

    if args.md:
        os.makedirs(os.path.dirname(os.path.abspath(args.md)), exist_ok=True)
        open(args.md, 'w', encoding='utf-8').write(to_md(rows, now))
        print('\n已写入 Markdown:', args.md)


if __name__ == '__main__':
    main()

# -*- coding: utf-8 -*-
r"""
单客户聊单档案（换人只改一个名字）。

用法:
  python copilot.py 张同学
  python copilot.py 24届某985会计学家长 --tail 40
  python copilot.py --list                 # 列出可匹配的客户名
  python copilot.py 张三 --json            # 结构化输出（供程序消费）

输出：身份 / 你的聊单笔记 / 关系概览 / 当前阶段(推断) / 成交信号 / 风险预警 /
      跟进建议 / 话术库该翻哪一页 / 最近对话原文

说明：本脚本只做「事实检索 + 规则判断」，最终回复话术由 AI 依据
      references/01-成交话术库.md 结合上下文撰写。
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import wxlib as W  # noqa: E402

PLAYBOOK_MAP = [
    (['合同', '甲方', 'e签', '签署', '身份证'],
     '第一章·第6步「签约交付」：甲方信息表 → e签宝 → 对公账户 → 建群承诺'),
    (['退费', '退款', '保障', '不成功', '没上岸'],
     '第二章 2.3 效果顾虑 + 第四章 风险逆转库（把风险挪到自己身上）'),
    (['优惠', '贵', '价格', '便宜', '预算', '费用', '分期'],
     '第二章 2.1 价格异议：守价 → 帮您向领导申请 → 给「争取来的」优惠'),
    (['考虑一下', '考虑考虑', '再看看', '商量', '不急', '过几天', '回头'],
     '第二章 2.2 拖延异议：把拖延换算成孩子的代价（挂期末/绩点节点）'),
    (['对比', '别家', '其他机构', '保研岛', '去保研'],
     '第二章 2.4 对比异议：不比价格比确定性（品牌年限 + 案例量 + 奖项）'),
    (['名额', '开课', '什么时候', '截止', '开班'],
     '第三章 逼单三件套：名额稀缺 + 时间节点 + 领导权限'),
    (['师资', '课时', '导师', '助教', '教授'],
     '第一章·第3步 项目卡 + 第五章 标准件清单（流程长图 / 项目卡）'),
    (['排名', '绩点', '目标院校', '保研政策', '竞赛', '论文'],
     '第一章·第2步 诊断定档：学校保研政策 + 84 项竞赛名录 + 把焦虑具体化'),
]

STAGE_ACTION = {
    'S0 新线索/未分类': '先破冰建档：发规划师自我介绍 + 信息采集表（7 项）。别急着报价。',
    'S1 破冰建档': '我方已建档 → 发保研政策 + 竞赛名录，把「不确定性」讲透，约腾讯会议。',
    'S2 诊断定档': '发方案 + 约会议收口；会议后当天出书面规划，制造「专业且上心」感。',
    'S3 方案报价': '先给成果（项目卡）再给价格；同时给两档（分期/全款）+ 锚定高价档。',
    'S4 破异议': '逐条拆异议：价格走领导权限、效果走退费保障、对比走品牌确定性。',
    'S5 逼单': '顺势收口：名额 + 时间节点 + 领导权限三件套，直接推甲方信息表 + e签宝。',
    'S6 签约交付': '24h 内闭环：甲方信息表 → e签宝 → 对公账户 → 当晚建服务群。',
}

POST_ACTION = {
    'P3 风险处置': '当轮响应：明确时间点 + 专人对接 + 主动同步进度，绝不拖。拖一次沟通就变对抗。',
    'P2 交付期': '交付可视化：每节课后给一条反馈，每周一条进度小结，让家长「看得见服务」。',
    'P1 维护期': '经营复购：保研节点关怀（夏令营/预推免/结果季）+ 转介绍开口。',
}

POST_PLAYBOOK = {
    'P3 风险处置': '第八章 风控提醒：退费当轮响应 + 书面时间表 + 专人对接（防对抗升级）',
    'P2 交付期': '第八章 风控提醒：签后 48h 交付清单 + 每节课反馈（防反悔）',
    'P1 维护期': '第六章 按客户类型分打法：转介绍直接给力度，快速成单',
}


def match(candidates, key):
    key = key.strip()
    exact = [c for c in candidates if c['wxid'] == key]
    if exact:
        return exact
    return [c for c in candidates if key in (c['name'] or '') or key in (c['wxid'] or '')]


def build(cfg, name, tail=30):
    contacts, self_wxid = W.read_contacts(cfg)
    chats = W.load_chats(cfg)

    cands = []
    for wxid, ch in chats.items():
        ct = contacts.get(wxid) or {}
        nm = ct.get('remark') or ct.get('nick') or ch.get('chat') or wxid
        cands.append({'wxid': wxid, 'name': nm, 'chat': ch, 'contact': ct})

    hit = match(cands, name)
    if not hit:
        return {'error': '未找到客户: %s' % name,
                'candidates': [c['name'] for c in cands[:60]]}
    if len(hit) > 1:
        return {'error': '匹配到多个，请用更完整的名字或直接给 wxid',
                'candidates': ['%s (%s)' % (c['name'], c['wxid']) for c in hit[:20]]}

    h = hit[0]
    ct, ch = h['contact'], h['chat']
    msgs = ch['msgs']
    stats = W.side_stats(msgs)
    sil = W.silence_days(msgs)

    deal = W.is_deal(ct)
    refunded = W.is_refunded(ct)
    kind = W.client_kind(msgs, contacts, h['wxid'])

    if deal:
        stage, evidence = W.post_sale_stage(msgs)
        stage_action = POST_ACTION.get(stage, '')
        if refunded:
            stage, evidence = 'P3 风险处置', '已退费'
            stage_action = POST_ACTION['P3 风险处置']
    else:
        stage, evidence = W.infer_stage(msgs)
        stage_action = STAGE_ACTION.get(stage, '')

    sig = W.scan_signals(msgs)
    risk = W.scan_risks(msgs)

    cadence = cfg['cadence_days'].get(kind, 4)
    overdue = (sil is not None and sil >= cadence)

    playbook = []
    if deal:
        if POST_PLAYBOOK.get(stage):
            playbook.append(POST_PLAYBOOK[stage])
    else:
        text = '\n'.join((m.get('content') or '') for m in msgs if not m.get('is_self'))[-4000:]
        for kws, ref in PLAYBOOK_MAP:
            if any(k in text for k in kws):
                playbook.append(ref)

    first = msgs[0]['ts'][:16] if msgs else '-'
    last = msgs[-1]['ts'][:16] if msgs else '-'

    return {
        'name': h['name'], 'wxid': h['wxid'], 'self_wxid': self_wxid,
        'segment': W.segment(h['name'], ct.get('labels'), msgs),
        'labels': ct.get('labels') or [], 'remark': ct.get('remark') or '',
        'note': ct.get('description') or '', 'alias': ct.get('alias') or '',
        'is_deal': deal, 'refunded': refunded, 'kind': kind,
        'total': stats['total'], 'me_pct': round(stats['me_pct']),
        'voice': stats['voice'], 'attach': stats['attach'],
        'first': first, 'last': last, 'silence_days': sil,
        'cadence_days': cadence, 'overdue': overdue,
        'stage': stage, 'stage_evidence': evidence, 'stage_action': stage_action,
        'signals': [{'name': n, 'weight': w, 'count': c} for n, w, c in sig['hits']],
        'signal_score': sig['score'],
        'risks': [{'name': n, 'weight': w, 'count': c} for n, w, c in risk['hits']],
        'risk_level': risk['level'],
        'playbook': playbook,
        'recent': [{'ts': m.get('ts'), 'who': '我' if m.get('is_self') else '客',
                    'content': W.clean_content(m.get('content'))} for m in msgs[-tail:]],
    }


def render(d):
    if d.get('error'):
        print('!!', d['error'])
        for c in d.get('candidates') or []:
            print('  -', c)
        return
    L = []
    L.append('=' * 78)
    L.append('客户档案：%s   [%s]' % (d['name'], d['wxid']))
    L.append('=' * 78)
    tags = '、'.join(d['labels']) or '（无）'
    role = '老客户/复购' if d['is_deal'] else '新线索'
    L.append('【身份】%s | 分层: %s | 标签: %s | 备注: %s'
             % (role, d.get('segment', '客户'), tags, d['remark'] or '-'))
    if d['refunded']:
        L.append('         ⚠ 疑似已退费 —— 走「交付复盘 + 关系修复」，不要硬推新单')
    if d['note']:
        L.append('【你的聊单笔记】%s' % d['note'])
    L.append('【关系概览】%d 条（我 %d%% / 客 %d%%）| 语音 %d | 附件 %d | 首触 %s | 最近 %s | 静默 %s 天'
             % (d['total'], d['me_pct'], 100 - d['me_pct'], d['voice'], d['attach'],
                d['first'], d['last'], d['silence_days']))
    L.append('【当前阶段·推断】%s%s'
             % (d['stage'], '（命中“%s”）' % d['stage_evidence'] if d['stage_evidence'] else ''))
    if d['stage_action']:
        L.append('          → %s' % d['stage_action'])
    if d['signals']:
        L.append('【成交意向信号】'
                 + '  '.join('%s(★%d×%d)' % (s['name'], s['weight'], s['count']) for s in d['signals'])
                 + '   总强度 %d' % d['signal_score'])
    if d['risks']:
        L.append('【风险预警】' + '  '.join('%s(★%d)' % (r['name'], r['weight']) for r in d['risks'])
                 + '   最高 ★%d' % d['risk_level'])
    L.append('【跟进节奏】类型=%s → 建议每 %d 天触达一次；已静默 %s 天 → %s'
             % (d['kind'], d['cadence_days'], d['silence_days'],
                '⚠ 已超期，今天必须触达' if d['overdue'] else '节奏正常'))
    if d['playbook']:
        L.append('【翻话术库这几节】')
        for p in d['playbook']:
            L.append('   · ' + p)
    L.append('-' * 78)
    L.append('【最近 %d 条对话】' % len(d['recent']))
    for m in d['recent']:
        L.append('  [%s] %s: %s' % (m['ts'][:16], m['who'], m['content'].replace('\n', ' / ')[:160]))
    print('\n'.join(L))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('client', nargs='?', help='客户名（备注/昵称的一部分）或 wxid')
    ap.add_argument('--tail', type=int, default=30, help='打印最近多少条对话')
    ap.add_argument('--json', action='store_true')
    ap.add_argument('--list', action='store_true', help='列出所有可匹配的客户')
    args = ap.parse_args()

    cfg = W.load_config()
    if args.list or not args.client:
        contacts, _ = W.read_contacts(cfg)
        chats = W.load_chats(cfg)
        rows = []
        for wxid, ch in chats.items():
            ct = contacts.get(wxid) or {}
            nm = ct.get('remark') or ct.get('nick') or ch['chat']
            rows.append((nm, len(ch['msgs']), W.silence_days(ch['msgs']),
                         W.segment(nm, ct.get('labels'), ch['msgs']), '已成交' if W.is_deal(ct) else ''))
        rows.sort(key=lambda x: -x[1])
        for n, c, s, seg, t in rows:
            print('%-36s %6d 条   静默 %s 天  %-10s %s' % (n[:34], c, s, seg, t))
        return

    d = build(cfg, args.client, args.tail)
    if args.json:
        print(json.dumps(d, ensure_ascii=False, indent=1))
    else:
        render(d)


if __name__ == '__main__':
    main()

# -*- coding: utf-8 -*-
r"""
解析微信通讯录的「已成交」标签，输出准确成交客户名单（wxid 文件）。

- 标签表: contact.db -> contact_label（label_id=42 -> '已成交'）
- 联系人标签: contact.extra_buffer (protobuf) 的 field 30 = 逗号分隔的标签 ID

也可直接命令行跑：
  python parse_labels.py            # 打印名单
  python parse_labels.py --wxids    # 只输出 wxid 清单文件路径
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import wxlib as W  # noqa: E402


def collect(cfg):
    contacts, self_wxid = W.read_contacts(cfg)
    out = []
    for wxid, c in contacts.items():
        if W.is_deal(c):
            out.append({
                'username': wxid,
                'remark': c['remark'] or c['nick'],
                'nick': c['nick'],
                'labels': c['labels'],
                'description': c['description'],
                'by_label': '已成交' in c['labels'],
                'refunded': W.is_refunded(c),
            })
    out.sort(key=lambda x: (not x['by_label'], x['remark']))
    return out, self_wxid


def export_deal_wxids(cfg=None, out_dir=None):
    cfg = cfg or W.load_config()
    deals, _self = collect(cfg)
    out_dir = out_dir or cfg['out_dir']
    os.makedirs(out_dir, exist_ok=True)
    wxid_file = os.path.join(out_dir, 'deal_wxids.txt')
    with open(wxid_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(d['username'] for d in deals))
    json.dump(deals, open(os.path.join(out_dir, 'deals_labeled.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)
    return len(deals), wxid_file


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--wxids', action='store_true', help='只输出 wxid 清单')
    args = ap.parse_args()
    cfg = W.load_config()
    deals, self_wxid = collect(cfg)
    by_label = [d for d in deals if d['by_label']]
    refunded = [d for d in deals if d['refunded']]

    if args.wxids:
        n, p = export_deal_wxids(cfg)
        print(p)
        return

    print('本人 wxid:', self_wxid)
    print('命中成交（标签或备注）:', len(deals), ' 其中带【已成交】标签:', len(by_label))
    print('其中疑似已退费:', len(refunded))
    print('\n--- 带【已成交】标签的客户 ---')
    for i, d in enumerate(by_label, 1):
        flag = '  ⚠退费' if d['refunded'] else ''
        other = [t for t in d['labels'] if t != '已成交']
        print('%3d. %-32s 其他标签=%s%s' % (i, d['remark'][:30], other, flag))
    rest = [d for d in deals if not d['by_label']]
    if rest:
        print('\n--- 仅备注/笔记含「成交」的客户 ---')
        for i, d in enumerate(rest, 1):
            print('%3d. %-32s 标签=%s' % (i, d['remark'][:30], d['labels']))
    n, p = export_deal_wxids(cfg)
    print('\n已写出: %s （%d 个）' % (p, n))


if __name__ == '__main__':
    main()

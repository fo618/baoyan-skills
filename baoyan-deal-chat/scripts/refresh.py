# -*- coding: utf-8 -*-
r"""
刷新聊天数据（每天开工第一步）。

  1) 增量导出最近 N 天单聊 -> <work_dir>/live/   （默认 N 从 config 取，60 天）
  2) --with-deals 时，同步刷新「已成交」标签名单与其完整历史 -> <work_dir>/成交客户原始记录/

密钥已存本机，无需重启微信、无需重新登录。
用法:
  python refresh.py                 # 常规刷新
  python refresh.py --with-deals    # 连成交客户全量一起刷新（较慢）
"""
import argparse
import os
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import wxlib as W  # noqa: E402


def safe_clean(path, must_under, must_name=None):
    """只删自己管的目录，避免误伤。"""
    p = os.path.abspath(path)
    if not p.startswith(os.path.abspath(must_under)):
        raise SystemExit('拒绝清理非工作目录: %s' % p)
    if must_name and os.path.basename(p) != must_name:
        raise SystemExit('拒绝清理非预期目录: %s' % p)
    if os.path.isdir(p):
        shutil.rmtree(p)
    os.makedirs(p, exist_ok=True)


def run_export(cfg, out_dir, since_days=None, wxid_file=None):
    cmd = [cfg['python'], cfg['export_script'],
           '--data-root', cfg['data_root'],
           '--key-file', cfg['key_file'],
           '--out-dir', out_dir,
           '--kinds', cfg.get('kinds', '单聊')]
    if wxid_file:
        cmd += ['--wxid-file', wxid_file, '--min-messages', '1']
    else:
        cmd += ['--since-days', str(since_days if since_days is not None else cfg['live_days']),
                '--min-messages', '2']
    print('  $', ' '.join(cmd))
    r = subprocess.run(cmd, capture_output=True, text=True,
                       encoding='utf-8', errors='replace')
    tail = (r.stdout or '').strip().splitlines()[-12:]
    for line in tail:
        print('   ', line)
    if r.returncode != 0:
        print((r.stderr or '')[-1500:])
        raise SystemExit('导出失败 returncode=%d' % r.returncode)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--days', type=int, help='最近 N 天有消息的会话（默认取 config.live_days）')
    ap.add_argument('--with-deals', action='store_true', help='同时刷新成交客户完整历史')
    ap.add_argument('--skip-live', action='store_true', help='跳过 live 增量（只刷新成交名单时用）')
    args = ap.parse_args()

    cfg = W.load_config()
    wd = cfg['work_dir']
    print('工作目录:', wd)

    if not args.skip_live:
        print('\n[1/2] 增量导出 live/ ...')
        safe_clean(cfg['live_dir'], wd, 'live')
        run_export(cfg, cfg['live_dir'], since_days=args.days)

    if args.with_deals:
        print('\n[2/2] 刷新「已成交」名单 + 完整历史 ...')
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from parse_labels import export_deal_wxids
        n, wxid_file = export_deal_wxids(cfg, out_dir=os.path.join(cfg['out_dir']))
        print('    已成交客户:', n)
        safe_clean(cfg['deal_dir'], wd, '成交客户原始记录')
        run_export(cfg, cfg['deal_dir'], wxid_file=wxid_file)

    print('\n完成。live:', cfg['live_dir'])


if __name__ == '__main__':
    main()

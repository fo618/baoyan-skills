# -*- coding: utf-8 -*-
r"""
baoyan-deal-chat 公共库

职责：
  1. 读 config.json（路径、节奏阈值）
  2. 解密 contact.db -> 联系人档案（备注 / 标签 / 聊单笔记 description）
  3. 合并所有数据集的 all_messages.jsonl -> {wxid: [msg, ...]}
  4. 信号词表 / 风险词表 / 阶段推断 / 静默天数

本模块只读，不写任何微信数据。
"""
import collections
import datetime
import json
import os
import sys
import tempfile

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(SKILL_DIR, 'config.json')
# 兜底目录：优先用 config.json 的 export_script 所在目录，找不到时回落到当前用户的技能目录
FALLBACK_EXPORT_DIR = os.path.join(
    os.path.expanduser('~'), '.workbuddy', 'skills', 'wechat-chat-export', 'scripts')

TZ = datetime.timezone(datetime.timedelta(hours=8))


# ---------------------------------------------------------------- 配置

def load_config(path=None):
    cfg = json.load(open(path or CONFIG_PATH, encoding='utf-8'))
    wd = cfg['work_dir'].rstrip('/')
    cfg['work_dir'] = wd
    cfg.setdefault('live_dir', wd + '/live')
    cfg.setdefault('full_dir', wd + '/导出')
    cfg.setdefault('deal_dir', wd + '/成交客户原始记录')
    cfg.setdefault('daily_dir', wd + '/日报')
    cfg.setdefault('out_dir', os.path.join(SKILL_DIR, 'outbox'))
    if not cfg.get('datasets'):
        cfg['datasets'] = [cfg['live_dir'], cfg['full_dir'], cfg['deal_dir']]
    exp = cfg.get('export_script')
    cfg['export_scripts_dir'] = os.path.dirname(exp) if exp else FALLBACK_EXPORT_DIR
    # 个人回复缓冲：批量导出的 jsonl 首行有 BOM 的容错
    return cfg


def export_helpers(cfg=None):
    """import 上游 wechat-chat-export 的底层函数（decrypt_file / query / hmac_ok_file ...）。"""
    cfg = cfg or load_config()
    d = cfg['export_scripts_dir']
    if d not in sys.path:
        sys.path.insert(0, d)
    import export_chat  # noqa: E402
    return export_chat


# ---------------------------------------------------------------- 通讯录

def _read_varint(b, i):
    shift = 0
    val = 0
    while i < len(b):
        x = b[i]
        i += 1
        val |= (x & 0x7f) << shift
        if not (x & 0x80):
            break
        shift += 7
    return val, i


def parse_proto_labels(buf):
    """解析 contact.extra_buffer（protobuf）里 field 30 = 逗号分隔的标签 ID 列表。"""
    if not buf:
        return []
    i = 0
    out = None
    try:
        while i < len(buf):
            key, i = _read_varint(buf, i)
            fno = key >> 3
            wt = key & 7
            if wt == 0:
                _v, i = _read_varint(buf, i)
            elif wt == 2:
                ln, i = _read_varint(buf, i)
                seg = buf[i:i + ln]
                i += ln
                if fno == 30:
                    out = seg.decode('ascii', 'ignore')
            elif wt == 5:
                i += 4
            elif wt == 1:
                i += 8
            else:
                break
    except Exception:
        pass
    if not out:
        return []
    return [int(x) for x in out.split(',') if x.strip().isdigit()]


def read_contacts(cfg=None, key_hex=None):
    """返回 (contacts, self_wxid)。contacts: {wxid: {...}}"""
    cfg = cfg or load_config()
    ec = export_helpers(cfg)
    key = bytes.fromhex(
        key_hex or open(cfg['key_file'], encoding='ascii').read().strip()
    )

    acct = None
    for d in ec.find_account_dirs(cfg['data_root']):
        p = os.path.join(d, 'db_storage', 'contact', 'contact.db')
        if os.path.exists(p) and ec.hmac_ok_file(key, p):
            acct = d
            break
    if not acct:
        raise SystemExit('密钥未能解锁 %s 下任何账号（key 过期或账号不同）' % cfg['data_root'])

    acct_name = os.path.basename(acct.rstrip('\\/'))
    self_wxid = cfg.get('self_wxid') or ''
    if not self_wxid:
        m = __import__('re').match(r'^(.*)_[0-9a-fA-F]{2,8}$', acct_name)
        self_wxid = m.group(1) if m else ''

    tmp = tempfile.mkdtemp(prefix='wxcontact_')
    cdb = os.path.join(tmp, 'contact.db')
    ec.decrypt_file(key, os.path.join(acct, 'db_storage', 'contact', 'contact.db'), cdb)

    _c, labels = ec.query('SELECT label_id_, label_name_ FROM contact_label', cdb)
    label_map = {int(r[0]): r[1] for r in labels}

    _c, rows = ec.query(
        'SELECT username, remark, nick_name, alias, description, extra_buffer FROM contact', cdb)
    contacts = {}
    for r in rows:
        lids = parse_proto_labels(r[5])
        contacts[r[0]] = {
            'username': r[0],
            'remark': r[1] or '',
            'nick': r[2] or '',
            'alias': r[3] or '',
            'description': r[4] or '',          # 你自己的聊单笔记
            'labels': [label_map.get(i, '?%d' % i) for i in lids],
            'label_ids': lids,
        }
    return contacts, self_wxid


def is_deal(contact):
    if not contact:
        return False
    if '已成交' in (contact.get('labels') or []):
        return True
    txt = (contact.get('remark') or '') + (contact.get('description') or '')
    return '成交' in txt


def is_refunded(contact):
    if not contact:
        return False
    txt = (contact.get('remark') or '') + (contact.get('description') or '')
    return any(w in txt for w in ['已退费', '退费中', '已退款'])


# ---------------------------------------------------------------- 消息集

def iter_jsonl(paths):
    for p in paths:
        if not os.path.isfile(p):
            continue
        for line in open(p, encoding='utf-8-sig'):
            line = line.strip()
            if not line:
                continue
            try:
                yield json.loads(line)
            except Exception:
                continue


def load_chats(cfg=None, wxids=None, dirs=None):
    """合并各数据集的 jsonl。返回 OrderedDict[wxid] = {'chat','wxid','msgs':[...]}"""
    cfg = cfg or load_config()
    dirs = dirs or cfg['datasets']
    paths = [os.path.join(d, 'all_messages.jsonl') for d in dirs]
    chats = collections.OrderedDict()
    seen = set()
    for rec in iter_jsonl(paths):
        wxid = rec.get('chat_wxid') or rec.get('chat')
        if wxids is not None and wxid not in wxids:
            continue
        key = (wxid, rec.get('ts'), rec.get('is_self'), (rec.get('content') or '')[:60])
        if key in seen:
            continue
        seen.add(key)
        c = chats.setdefault(wxid, {'chat': rec.get('chat') or wxid, 'wxid': wxid, 'msgs': []})
        c['msgs'].append(rec)
    for c in chats.values():
        c['msgs'].sort(key=lambda m: m.get('ts') or '')
    return chats


def ts_to_dt(s):
    try:
        return datetime.datetime.strptime(s[:19], '%Y-%m-%d %H:%M:%S').replace(tzinfo=TZ)
    except Exception:
        try:
            return datetime.datetime.strptime(s[:10], '%Y-%m-%d').replace(tzinfo=TZ)
        except Exception:
            return None


def silence_days(msgs, now=None):
    now = now or datetime.datetime.now(TZ)
    if not msgs:
        return None
    last = None
    for m in reversed(msgs):
        d = ts_to_dt(m.get('ts') or '')
        if d:
            last = d
            break
    if not last:
        return None
    return (now - last).days


def side_stats(msgs):
    n = len(msgs)
    me = sum(1 for m in msgs if m.get('is_self'))
    voice = sum(1 for m in msgs if (m.get('content') or '').startswith('[语音'))
    files = sum(1 for m in msgs if (m.get('content') or '').startswith('['))
    return {'total': n, 'me': me, 'other': n - me,
            'me_pct': (me / n * 100) if n else 0,
            'voice': voice, 'attach': files}


# ---------------------------------------------------------------- 信号 / 风险 / 阶段

SIGNALS = [
    ('要合同/问签约', 5, ['合同', '协议', 'e签', '签合同', '怎么签', '甲方', '身份证', '实名', '签署']),
    ('问退费/保障', 5, ['退费', '退款', '不成功', '没上岸', '保障', '保底', '风险', '承诺']),
    ('问价格/优惠', 5, ['优惠', '便宜', '多少钱', '费用', '价格', '报价', '分期', '贵', '预算']),
    ('问名额/开课', 4, ['名额', '开课', '什么时候开始', '什么时候开', '还有位置', '排课', '开班']),
    ('问师资/课时', 4, ['师资', '导师', '课时', '助教', '教授', '老师是', '谁带', '上课']),
    ('聊学情/目标', 3, ['排名', '绩点', '专业人数', '保研率', '目标院校', '想去', '保研', '夏令营']),
    ('对比竞品', 3, ['对比', '别家', '其他机构', '保研岛', '去保研', '哪家', '另一家']),
    ('提决策人', 3, ['孩子', '家长', '妈妈', '爸爸', '商量', '舅舅', '阿姨', '家里']),
]

RISKS = [
    ('退费/投诉', 3, ['退费', '退款', '投诉', '不满意', '差评', '曝光', '维权', '举报']),
    ('明确放弃', 3, ['不考虑', '算了', '不报了', '取消', '不继续', '放弃', '别联系']),
    ('竞品比价', 2, ['对比', '别家', '其他机构', '保研岛', '去保研', '哪家']),
    ('价格受阻', 2, ['太贵', '贵', '预算不够', '负担', '便宜点', '优惠多少']),
    ('决策拖延', 1, ['考虑一下', '考虑考虑', '再看看', '商量一下', '过几天', '回头说', '不急']),
]

# 由晚到早；命中即停（用最近窗口判断当前位置）
STAGES = [
    ('S6 签约交付', ['甲方信息', '身份证', 'e签', '签署', '合同链接', '对公', '账户', '转账', '定金', '付款', '缴费']),
    ('S5 逼单', ['名额', '最后一批', '截止', '月底', '抓紧', '定下来', '敲定', '锁定', '今天定']),
    ('S4 破异议', ['优惠', '太贵', '有点贵', '对比', '别家', '退费', '保障', '成功率', '考虑一下', '商量']),
    ('S3 方案报价', ['价格', '费用', '课时', '方案', '多少钱', '报价', '分期', '项目', '课时数']),
    ('S2 诊断定档', ['保研政策', '竞赛', '绩点', '排名', '目标院校', '论文', '科研', '夏令营', '推免']),
    ('S1 破冰建档', ['信息采集', '哪年入学', '本科院校', '自我介绍', '规划师', '基本情况']),
]

CLIENT_KINDS = ('闪电单', '常规', '长线')

# ---- 成交之后的客户（老客户/复购）走另一套阶段，不要套销售漏斗 ----
POST_STAGES = [
    ('P3 风险处置', ['退费', '退款', '投诉', '不满意', '维权', '要退', '退一下', '曝光', '差评']),
    ('P2 交付期', ['上课', '课时', '辅导', '老师', '安排', '进度', '论文', '竞赛', '夏令营', '群', '排课']),
    ('P1 维护期', []),
]


def post_sale_stage(msgs, window=20):
    """已成交客户：判断处于 风险处置 / 交付期 / 维护期。"""
    text = '\n'.join((m.get('content') or '') for m in msgs[-window:])
    for name, kws in POST_STAGES:
        if not kws:
            return name, ''
        for k in kws:
            if k in text:
                return name, k
    return 'P1 维护期', ''


def clean_content(c):
    """把非文本消息（原始 XML / 卡片）压成一个可读标签，避免刷屏。"""
    c = (c or '').strip()
    if not c:
        return ''
    if c.startswith('<?xml') or c.startswith('<msg'):
        if 'bigheadimgurl' in c or 'wx.qlogo.cn' in c:
            return '[名片/链接卡片]'
        return '[卡片消息]'
    if c.startswith('<') and '>' in c[:40]:
        return '[卡片消息]'
    return c


# ---------------------------------------------------------------- 身份分层

CLIENT_LABELS = {'家长', '学生'}
BOT_KW = ['clawbot', 'bot', '助手', '服务号', '客服', '小助手', '公众号', '订阅号', 'workbuddy',
          '元宝', 'deepseek', 'kimi', '文心', '豆包', '通义', 'chatgpt', 'claude', 'copilot']
# ⚙️ 换成你自己微信里的渠道/同行标签名（通讯录标签，一字不差）。
#    本包已清空原作者的关系网数据；留空不影响其它功能，只是渠道不会被自动剔出客户池。
CHANNEL_LABELS = {
    # 例：'某渠道总', '高报-某老师', '同行-某机构', '同事',
}
# 名字里出现这些词 -> 判为同行/机构方。可按需增删（含自家品牌名时要小心，会误伤同事）
PEER_KW = ['升学规划', '保研规划', '生涯规划', '高考规划', '志愿规划', '留学机构']

# 内部同事 / 渠道协作专用词汇（学生/家长不会这么说）
INTERNAL_KW = ['师资群', '客诉', '加急', '接手', '教研费', '教务教研', '内部加红',
               '和你说下', '咱们内部', '我拉个群', '交付同学', '排课表',
               '来源：', '已转', '能转走', '跟进进度', '抖音私信']

_INTERNAL_CACHE = None


def _internal_names():
    global _INTERNAL_CACHE
    if _INTERNAL_CACHE is None:
        try:
            _INTERNAL_CACHE = set(json.load(open(CONFIG_PATH, encoding='utf-8')).get('internal_names') or [])
        except Exception:
            _INTERNAL_CACHE = set()
    return _INTERNAL_CACHE


def looks_internal(msgs, min_hits=2):
    """对方的话语里全是内部运营词 -> 同事/教务，不是客户。"""
    other = '\n'.join((m.get('content') or '') for m in msgs if not m.get('is_self'))
    return sum(1 for k in INTERNAL_KW if k in other) >= min_hits


def looks_like_peer(name, labels=None):
    """名字/标签像是同行、渠道、机构老师（不是学生/家长）。"""
    name = name or ''
    labels = set(labels or [])
    if '老师' in name or '规划师' in name or any(k in name for k in PEER_KW):
        # 名字里带院校/年级信息 -> 反而可能是客户（如「23届合肥工业电子信息」）
        if any(k in name for k in ['大学', '学院', '校区']) or any(
                k in name for k in ['级', '届']) and not any(k in name for k in PEER_KW):
            return False
        return True
    if any(k in name for k in ['cto', 'CTO', 'CEO', '中介', '渠道']):
        return True
    return bool(labels & CHANNEL_LABELS)


def peer_by_my_words(msgs, name, min_hits=2):
    """我平时管他叫「X老师」-> 多半是同行/渠道，不是客户。"""
    import re as _re
    core = _re.sub(r'[^\u4e00-\u9fa5]', '', name or '')
    if not core:
        return False
    cands = set()
    for i in range(len(core)):
        for j in range(i + 1, min(i + 3, len(core)) + 1):
            cands.add(core[i:j] + '老师')
    mine = '\n'.join((m.get('content') or '') for m in msgs if m.get('is_self'))
    return sum(mine.count(c) for c in cands) >= min_hits


def segment(name, labels=None, msgs=None):
    """返回 '机器人' / '内部同事' / '渠道·同行' / '客户'。"""
    low = (name or '').lower()
    if any(k in low for k in BOT_KW):
        return '机器人'
    if name in _internal_names():
        return '内部同事'
    if looks_like_peer(name, labels):
        return '渠道·同行'
    if msgs and peer_by_my_words(msgs, name):
        return '渠道·同行'
    if msgs and looks_internal(msgs):
        return '内部同事'
    return '客户'


def _hits(text, lexicon):
    out = []
    for name, weight, kws in lexicon:
        n = sum(1 for k in kws if k in text)
        if n:
            out.append((name, weight, n))
    return out


def scan_signals(msgs, window=None):
    """扫描客户侧消息（默认最近 40 条）命中的成交意向信号。"""
    dial = [m for m in msgs if not m.get('is_self')]
    if window:
        dial = dial[-window:]
    text = '\n'.join((m.get('content') or '') for m in dial)
    hits = _hits(text, SIGNALS)
    score = sum(w * min(n, 3) for _n, w, n in hits)
    return {'hits': hits, 'score': score}


def scan_risks(msgs, window=None):
    dial = [m for m in msgs if not m.get('is_self')]
    if window:
        dial = dial[-window:]
    text = '\n'.join((m.get('content') or '') for m in dial)
    hits = _hits(text, RISKS)
    top = max([w * min(n, 2) for _n, w, n in hits], default=0)
    return {'hits': hits, 'level': top}


def infer_stage(msgs, window=14):
    """用最近 window 条（双方）推断当前所处阶段。返回 (阶段名, 证据文本)。"""
    recent = msgs[-window:]
    text = '\n'.join((m.get('content') or '') for m in recent)
    for name, kws in STAGES:
        for k in kws:
            if k in text:
                return name, k
    return 'S0 新线索/未分类', ''


def client_kind(msgs, contacts, wxid):
    """闪电单 / 常规 / 长线 —— 按成交周期规律 + 当前静默天数给跟进节奏。"""
    msgs = [m for m in msgs if m.get('ts')]
    if not msgs:
        return '常规'
    first = ts_to_dt(msgs[0]['ts'])
    last = ts_to_dt(msgs[-1]['ts'])
    span = (last - first).days if (first and last) else 0
    sil = silence_days(msgs)
    if span <= 7 and (sil is not None and sil <= 2):
        return '闪电单'
    if span > 30:
        return '长线'
    return '常规'

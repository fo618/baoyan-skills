# 保研聊单教练 · 装机上手（10 分钟跑通）

一套**只读**系统：读你本机微信聊天 → 排优先级 → AI 出话术，**发送永远由你自己点**。
不看你的聊天记录以外的任何东西，数据全落本地。

---

## 0. 先知道：这个包里没有我的任何客户数据

本包做过隐私清洗，**不含**原作者的：客户 wxid / 姓名 / 昵称 / 院校备注 / 手机号、
95 人成交名单、渠道同行名单、公司收款账户、机构名与自称、成交复盘报告。

所以第一次用，下面两件事**必须你自己做**（做完这套东西才是你的）：

1. 填 `config.json` 里的 5 个路径（第 2 步）
2. 全局搜 `【`，把占位符换成你自己的信息（机构名 / 称呼 / 价目 / 账户，第 4 步）

跑完第 5 步，你会得到属于自己的成交名单、话术库和复盘报告。

---

## 1. 前置依赖（一次装完）

| 依赖 | 怎么来 | 备注 |
| --- | --- | --- |
| Python 3.10+ | 任意 Python，记下解释器**绝对路径** | 后面写进 `config.json` 的 `python` |
| pycryptodome / zstandard | `pip install pycryptodome zstandard` | 解密微信数据库用 |
| `wechat-chat-export` 技能 | 放在 `~/.workbuddy/skills/wechat-chat-export/` | 提供 `export_chat.py` / `export_all.py` / `capture_key.py` |
| 微信 PC 端 | **已登录并保持运行** | 数据库不运行就不更新，导出会是旧的 |

验证依赖装好：`<你的python> -c "import Crypto, zstandard; print('ok')"` 打印 `ok` 即可。

---

## 2. 填 config.json（只改这一处）

打开技能根目录的 `config.json`，把 `<你的用户名>` / `<你的工作目录>` 全部替换掉：

```jsonc
"python":        "C:/Users/<你的用户名>/.workbuddy/binaries/python/envs/default/Scripts/python.exe",
"data_root":     "C:/Users/<你的用户名>/xwechat_files",
"key_file":      "C:/Users/<你的用户名>/<你的工作目录>/key/key.txt",
"export_script": "C:/Users/<你的用户名>/.workbuddy/skills/wechat-chat-export/scripts/export_all.py",
"work_dir":      "C:/Users/<你的用户名>/<你的工作目录>/微信聊天记录",
"datasets":      ["…/微信聊天记录/live", "…/微信聊天记录/成交客户原始记录"]
```

- 路径统一用**正斜杠**，Windows 下 Python 同样认。
- `work_dir` 提前建好空目录，脚本会在下面自动建 `live/`、`成交客户原始记录/`、`日报/`。
- `export_script` 指向的目录里必须有 `export_chat.py`，否则解密函数 import 不到。

---

## 3. 抓微信数据库密钥（一次，长期有效）

```bash
"<python>" ~/.workbuddy/skills/wechat-chat-export/scripts/capture_key.py
```

产物是一个 64 位 hex 的 `key.txt`，把它放到第 2 步填的 `key_file` 路径。
微信保持登录就不用重抓（换账号才需要重抓）。

---

## 4. 把占位符换成你自己的（必做）

在技能目录里全局搜索 `【`，至少改这些地方：

| 位置 | 占位符 | 换成 |
| --- | --- | --- |
| `references/01-成交话术库.md` 头部 | 【机构名】【你的称呼】【你的从业年限】 | 你自己的 |
| 同上 · 信任三件套 | 【机构官网】【累计服务学员数】【行业/媒体奖项】 | 你机构的 |
| 同上 · 签约段 | 【公司支付宝账号】【公司对公账号】【公司全称】【开户行全称】 | **找财务核对后再填**，别凭记忆 |
| 同上 · 报价段 | 【主推计划名】【如 8.98w–9w】等全部金额 | 你的真实价目 |
| 同上 · 异议段 | 【我方公开成功率，以官方口径为准】 | 你机构对外口径 |
| `scripts/wxlib.py` | `CHANNEL_LABELS = {}` | 你微信里渠道/同行的标签名，空着也行 |
| `config.json` | `internal_names: []` | 你同事/教务的备注名 |

⚠️ `references/03` 里的百分比（49% / 22% / 28 天中位数）是**原作者样本的量级**，
不是你的。跑完第 5 步的复盘脚本后用自己的数字覆盖。

---

## 5. 第一次跑（按顺序来）

```bash
PY="<你的python>"      # 换成 config.json 里那个解释器绝对路径
SK="<技能目录>/scripts"
```

```bash
# ① 刷新数据：增量导出最近 60 天的单聊（约 40 秒）
"$PY" "$SK/refresh.py"

# ② 打开工作台：今天该先跟谁
"$PY" "$SK/pipeline.py"

# ③ 看某个客户：他卡在哪、该说哪一节
"$PY" "$SK/copilot.py" --list          # 先列出可匹配的客户名
"$PY" "$SK/copilot.py" "<客户备注名>" --tail 40
```

出话术是 AI 的活：把 `copilot.py` 的档案贴给 AI，让它按 `references/` 四份手册写稿。

---

## 6. 生成属于你自己的资产（跑过一次就齐全）

```bash
"$PY" "$SK/parse_labels.py"            # 成交名单（读微信「已成交」标签）-> outbox/
"$PY" "$SK/refresh.py" --with-deals    # 重导成交客户全量历史（较慢，每周一次够）
"$PY" "$SK/analysis/deal_review.py"    # 画像 / 节奏 / 证据覆盖 / 跨会话复用话术
"$PY" "$SK/analysis/objection_mine.py" # 异议→应答配对 / 逼单话术 / 风险逆转话术
```

拿到产出后让 AI 回填两处（这是这套工具"越用越厚"的地方）：
- `references/01-成交话术库.md` ← 用 `objection_mine.py` 挖出的**你自己的原话**替换示例话术
- `references/02-成交客户共性报告.html` ← 用 `deal_review.py` 的统计生成一份单文件 HTML
  （深红主题、可打印；结构照：一页速览 / 客户画像 / 4 个标准动作 / 成交周期 / 六步链 /
  四类异议应答 / 风控提醒）

`outbox/` 里的成交名单是你自己的隐私数据，别随包外发。

---

## 7. 三个最容易踩的坑

1. **`--json` 输出没过滤**：只有控制台版和 `--md` 版的客户池过滤过 `segment == "客户"`，
   `--json` 是原样输出，渠道/同行/机器人会排在最前。要结构化数据就自己再过滤一遍。
2. **风险标记误报率约一半**：★3 常来自客户转发的微信系统提示（含"我要投诉"）。
   **一律只看最近 20 条原文复核**，脚本结论只能当候选。
3. **别删 `live/` 旧数据**：安全钩子会拦 `safe_clean()`。要重建就把旧目录**改名**
   `mv live live_backup_<日期>`，别删除。

更多排障见 `SKILL.md` 第九节（三段式表格）。

---

## 8. 铁律（改不改都在这，别破）

1. 🛑 **全程只读**：绝不代发、绝不自动回复。出稿，你自己发。
2. **数据不出本机**：聊天记录、密钥、分析产物只落本地。
3. **成交判定只认微信「已成交」标签**，不要用"发过合同"去猜（会低估一个数量级）。
4. **缺信息不编**：金额、排名、成功率没核实就写定性表述。

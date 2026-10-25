# 安装说明（同事分发版）

这个包能干两件事：

1. **给学员背景 → 出 PPT 规划表**（学员信息表 + 冲稳保定档表，原生 PPT 表格，不是截图）
2. **给推免政策文件 → 出一份家长学生解读 HTML**（11 个标准模块，单文件，可直接打印/转 PDF）

> ⚠️ **包里不含 PPT 模板**。原模板带作者品牌页，已剔除。
> **你必须先放一个自己的 pptx 进去**，否则出不了 PPT。详见第 2 步。

---

## 第 1 步：装依赖

```bash
pip install python-pptx      # 主流程（出 PPT），必装
pip install pypdf            # 形态 C 抽 PDF 文本，用到政策解读时再装
```

可选：`pip install pymupdf`（只有在你主动给抓取脚本加 `--pdf` 时才需要，默认不装也能跑）。
`generate_tables.py`（已下线的 Excel 版）需要 `openpyxl`，主流程不用装。

装完验一下，别跳过：

```bash
python -c "import pptx; print('ok')"
```

> **最常见的坑**：系统里可能装了多个 Python，你 `pip install` 装到了 A，
> 跑脚本用的是 B，于是报 `ModuleNotFoundError: No module named 'pptx'`。
> 解决：固定用一个解释器，装和跑都用它的完整路径（venv 里的那个最稳），
> 并把 SKILL.md 第 74 行的 `【你的venv路径】` 替换成它（见第 5 步）。

---

## 第 2 步：放你的 PPT 模板（**不做这步跑不起来**）

把你的版式模板放进：

```
assets/templates/你的模板.pptx
```

- 文件名随便起，**目录里只放这一个**（放多个脚本只取排序第一个）。
- 任何 16:9 / 4:3 的 pptx 都能跑，没有硬性结构要求：脚本按页标题关键词找
  「学员信息」「院校定档」两页，找不到就自动克隆一页同版式页面再放表。
- 想做 26 页那种完整版式，看 `assets/templates/放你的模板到这里.md`。

放好之后，命令行**可以省略模板参数**：

```bash
python scripts/build_ppt_slides.py 我的学员.json
```

没放的话脚本会打印指引并直接退出，不会静默出错。

---

## 第 3 步：装到 skills 目录

把整个 `baoyan-planner-tables` 文件夹拷到：

| 系统 | 路径 |
|---|---|
| Windows | `C:\Users\<你的用户名>\.workbuddy\skills\baoyan-planner-tables` |
| macOS / Linux | `~/.workbuddy/skills/baoyan-planner-tables` |

拷完**重启一次 WorkBuddy** 才会被识别。

---

## 第 4 步：首次运行顺序（照这个走一遍）

```bash
# ① 复制示例数据源，改成你的学员
cp assets/sample_data.json student_data_张三.json

# ② 先空跑，看字号和有没有溢出（不写文件，最稳妥）
python scripts/build_ppt_slides.py student_data_张三.json --dry-run

# ③ 没问题再真跑（不写输出名，脚本自动命名并落到桌面）
python scripts/build_ppt_slides.py student_data_张三.json
```

出现 `⚠ 超限` 就回去**删字**，别靠压字号掩盖（定档表「适配院校」最吃空间，压到 ≤12 字，
长校名换简称，如 `哈尔滨工业大学（环境学院）` → `哈工大（环境学院）`）。

形态 C（政策解读）的两步：

```bash
python scripts/extract_policy_text.py "XX大学推免政策.pdf" --out policy_text.txt
python scripts/build_policy_html.py policy_data_XX大学.json "XX大学推免政策_解读.html"
```

---

## 第 5 步：全局搜 `【` 替换（**必做**）

包里留下了占位符。在你解压后的目录里全局搜 `【`，逐项替换：

| 占位符 | 出现在 | 替换成 |
|---|---|---|
| `【你的称呼】` | SKILL.md（5 处）、README.md（1 处）、test-prompts.json（1 处） | 你的自称/品牌名，例如「李老师」。不想品牌化就删掉这两个字 |
| `【你的venv路径】` | SKILL.md（2 处） | 你装了 python-pptx 的解释器路径，例如 Windows `C:/Users/<你>/.venv/Scripts/python.exe`、mac `~/.venv/bin/python` |

> ⚠️ **别动脚本里的 `【`**！`scripts/build_ppt_slides.py` 第 559~560 行的 `【.*?】`
> 和 `scripts/fetch_school_tuimian_policy.py` 里的 `【{scope}】` 是**正则和格式化字符串**，
> 改了会直接跑崩。只改上面表格里列的 `.md` / `.json`。

---

## 第 6 步：常见坑

| 症状 | 原因 / 处理 |
|---|---|
| `ModuleNotFoundError: No module named 'pptx'` | 装和跑不是同一个 Python。用完整路径跑（见第 1 步） |
| 提示「找不到模板 PPT」 | 第 2 步没做，往 `assets/templates/` 放一个 pptx |
| 输出只有 2 页 / 页数不对 | 正常——你的模板页数和参考版式不同，脚本会自动清掉没投表格的空页。不想清加 `--keep-empty-pages` |
| `× 院校层次自检不通过` | 定档里混进了双非/四非，或某档不是 4 所，或软科位次梯度反了。按打印的明细改数据；确认无误才用 `--force` |
| 表格文字溢出/字号太小 | 回去删字，或加 `--max-font 11` 下调上限 |
| PPT 找不到落哪儿了 | 默认写到**桌面**，文件名 `保研规划_<院校>.pptx`。加 `--no-desktop` 改回当前目录，或用 `--dir <路径>` |
| 抓院校政策拿到 3KB 空壳 | 那是跳转链接不是正文。跟到学院官网抓 `v_news_content` / `vsb_content` 容器，详见 SKILL.md 铁律 19 |
| Bash 里 `ls` / `dirname` command not found | 每条命令前先 `export PATH="/usr/bin:/bin:$PATH"` |

---

## 包里都有什么

```
baoyan-planner-tables/
├─ INSTALL.md                    ← 你正在看这个
├─ README.md                     快速上手：怎么触发、两种交付形态、底层命令、数据格式
├─ SKILL.md                      完整规范（932 行）：铁律、字段说明、排障手册、版本演进
├─ test-prompts.json             3 条自测用例，改完模板后跑一遍验收
├─ assets/
│  ├─ sample_data.json           数据源骨架（XX 占位，复制改自己的学员）
│  ├─ policy_sample_data.json    形态 C 的完整样例（某高校公开推免政策，照格式改）
│  └─ templates/
│     └─ 放你的模板到这里.md      ← 第 2 步要放 pptx 的地方
└─ scripts/
   ├─ build_ppt_slides.py        主流程：两张表写进 PPT
   ├─ generate_tables.py         已下线（Excel 版），仅作回溯
   ├─ extract_policy_text.py     形态 C：PDF → 纯文本
   ├─ build_policy_html.py       形态 C：数据 JSON → 解读 HTML
   ├─ fetch_yanzhaowang_policy.py      抓研招网推免政策
   └─ fetch_school_tuimian_policy.py   抓本校/本院推免方案
```

`README.md` 是日常看的，`SKILL.md` 是出问题或要改规则时查的。两个都不会自动执行任何东西，
改数据改 `assets/*.json` 就行，**不用碰脚本**。

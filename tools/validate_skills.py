#!/usr/bin/env python3
"""保研 AI Skills 静态门禁（CI 与本地通用）。

跑法： python tools/validate_skills.py
退出码：0 = 通过（可能有 warning），1 = 有 error。

**它能做什么**：技能包结构、SKILL.md front matter、test-prompts 结构、
所有 .py 能否编译、所有 .json 能否解析、明显隐私/密钥泄漏、README 图片引用是否存在。

**它不能做什么**（别假装）：真正跑一遍"给学员背景能不能出对 PPT"要模型，属于人工回归；
test-prompts.json 就是那份人工回归清单，改技能后照着跑。
"""
from __future__ import annotations

import json
import py_compile
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
errors: list[str] = []
warnings: list[str] = []

# Windows 控制台默认是 cp1252/GBK，print 中文会 UnicodeEncodeError（本地 GBK 能过、CI cp1252 会炸）。
# 强制 UTF-8 输出，见 2026-10-25 CI 首次红灯排查。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except (AttributeError, ValueError):
        pass

# 明显不该出现在公开仓里的东西（按需增删，别把正常业务词当泄漏）
SECRET_PATTERNS = [
    (re.compile(r"C:[/\\]Users[/\\][A-Za-z0-9._-]{2,}"), "真实用户目录绝对路径"),
    (re.compile(r"wxid_[A-Za-z0-9]{10,}"), "真实微信 wxid"),
    (re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)"), "手机号"),
    (re.compile(r"gh[pousr]_[A-Za-z0-9]{20,}"), "GitHub token"),
    (re.compile(r"sk-[A-Za-z0-9]{24,}"), "API key"),
    (re.compile(r"[A-Za-z0-9._%+-]+@(qq|163|gmail|outlook|foxmail)\.com"), "真实邮箱"),
]
SCAN_EXT = {".md", ".py", ".json", ".ps1", ".txt", ".yaml", ".yml"}
SKIP_DIRS = {".git", "__pycache__", ".github", "node_modules"}


def err(msg: str) -> None:
    errors.append(msg)


def warn(msg: str) -> None:
    warnings.append(msg)


def parse_front_matter(path: Path) -> dict[str, str]:
    """手写解析 YAML front matter（不引第三方依赖，CI 才能零安装跑）。"""
    text = path.read_text(encoding="utf-8")
    if not text.startswith("---"):
        raise ValueError("文件不是以 --- 开头")
    end = text.find("\n---", 3)
    if end == -1:
        raise ValueError("front matter 没有闭合的 ---")
    data: dict[str, str] = {}
    for line in text[3:end].splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if ":" not in line:
            continue
        key, _, value = line.partition(":")
        data[key.strip()] = value.strip()
    return data


def check_skill_dirs() -> list[Path]:
    skills: list[Path] = []
    for d in sorted(ROOT.iterdir()):
        if not d.is_dir() or d.name.startswith(".") or d.name in SKIP_DIRS or d.name in {"docs", "tools"}:
            continue
        skill_md = d / "SKILL.md"
        if not skill_md.exists():
            continue
        skills.append(d)
        try:
            fm = parse_front_matter(skill_md)
        except ValueError as exc:
            err(f"{d.name}/SKILL.md: front matter 解析失败 —— {exc}")
            continue
        name = fm.get("name", "")
        desc = fm.get("description", "")
        if not name:
            err(f"{d.name}/SKILL.md: front matter 缺 name")
        elif name != d.name:
            err(f"{d.name}/SKILL.md: name='{name}' 与目录名不一致（agent 会找不到技能）")
        if not desc:
            err(f"{d.name}/SKILL.md: front matter 缺 description（决定自动触发，不能空）")
        elif len(desc) < 60:
            warn(f"{d.name}/SKILL.md: description 只有 {len(desc)} 字，可能触发不准")
        if len(skill_md.read_text(encoding="utf-8")) < 500:
            warn(f"{d.name}/SKILL.md: 正文偏短（<500 字符），确认不是占位文件")

        tp = d / "test-prompts.json"
        if not tp.exists():
            warn(f"{d.name}: 没有 test-prompts.json（建议补 3 条回归用例）")
        else:
            try:
                cases = json.loads(tp.read_text(encoding="utf-8"))
            except json.JSONDecodeError as exc:
                err(f"{d.name}/test-prompts.json: JSON 解析失败 —— {exc}")
                continue
            if not isinstance(cases, list):
                err(f"{d.name}/test-prompts.json: 顶层必须是数组")
            elif len(cases) < 3:
                err(f"{d.name}/test-prompts.json: 至少 3 条回归用例，当前 {len(cases)} 条")
            else:
                for i, case in enumerate(cases):
                    if not isinstance(case, dict):
                        err(f"{d.name}/test-prompts.json[{i}]: 必须是对象")
                        continue
                    for key in ("id", "prompt", "expected"):
                        if key not in case or not str(case[key]).strip():
                            err(f"{d.name}/test-prompts.json[{i}]: 缺字段 {key}")
    if not skills:
        err("仓库根目录没找到任何含 SKILL.md 的技能目录")
    return skills


def check_python_and_json() -> None:
    for path in ROOT.rglob("*.py"):
        if SKIP_DIRS & set(path.parts):
            continue
        try:
            py_compile.compile(str(path), doraise=True, cfile=str(path) + ".pyc.tmp")
        except py_compile.PyCompileError as exc:
            err(f"{path.relative_to(ROOT)}: Python 语法错误 —— {exc.msg.strip().splitlines()[-1]}")
        finally:
            tmp = Path(str(path) + ".pyc.tmp")
            if tmp.exists():
                tmp.unlink()
    for path in ROOT.rglob("*.json"):
        if SKIP_DIRS & set(path.parts):
            continue
        try:
            json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            err(f"{path.relative_to(ROOT)}: JSON 解析失败 —— {exc}")


def check_secrets() -> None:
    for path in ROOT.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in SCAN_EXT:
            continue
        if SKIP_DIRS & set(path.parts):
            continue
        try:
            text = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        for pattern, label in SECRET_PATTERNS:
            for m in pattern.finditer(text):
                # config.json 里的 <你的用户名> 占位符不算泄漏
                snippet = m.group(0)
                if "<" in snippet or ">" in snippet:
                    continue
                err(f"{path.relative_to(ROOT)}: 疑似{label} —— {snippet[:40]}")


def check_readme_assets() -> None:
    readme = ROOT / "README.md"
    if not readme.exists():
        err("缺 README.md")
        return
    text = readme.read_text(encoding="utf-8")
    refs = re.findall(r"!\[[^\]]*\]\(([^)\s]+)\)", text)
    refs += re.findall(r'<img[^>]+src="([^"]+)"', text)
    for ref in refs:
        if ref.startswith(("http://", "https://")):
            continue
        target = (ROOT / ref).resolve()
        if not target.exists():
            err(f"README 引用的文件不存在：{ref}")


def main() -> int:
    skills = check_skill_dirs()
    check_python_and_json()
    check_secrets()
    check_readme_assets()

    print(f"扫描技能：{[s.name for s in skills]}")
    for w in warnings:
        print(f"  WARN  {w}")
    for e in errors:
        print(f"  ERROR {e}")
    if errors:
        print(f"\n✗ 未通过：{len(errors)} 个 error，{len(warnings)} 个 warning")
        return 1
    print(f"\n✓ 通过：0 个 error，{len(warnings)} 个 warning")
    return 0


if __name__ == "__main__":
    sys.exit(main())

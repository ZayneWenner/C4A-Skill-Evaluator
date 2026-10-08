#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
c4a_evaluate.py — C4 技能提交自动评审器 (C4A Skill Submission Evaluator)

把 wechat-doc-mapper 的"扫描→分类→缺口分析"能力升级为"自动阅卷"：
  ① 文件采集与识别  ② 提交完整性检查  ③ 技能质量评审(四条件)  ④ 评审报告生成

设计要点（详见 方案设计.md）：
  - 规则为主 + 可选 LLM 深审 的「混合」架构，但默认纯规则即可离线、确定性地跑通。
  - 所有检测信号从 references/c4_rubric.yaml 加载，避免把判断逻辑硬编码进代码。
  - 评分标准完全对齐 c4_rubric.yaml 的 scoring 段，保证"评分标准清晰、可复核"。

用法：
  python c4a_evaluate.py --folder <提交文件夹> [--report <报告.md>] [--excel <表.xlsx>]
                          [--rubric references/c4_rubric.yaml] [--top-n 3]

退出码：0 = 正常；2 = 未找到任何 C4 文件。
"""

import argparse
import datetime
import io
import json
import os
import re
import sys
import zipfile

try:
    import yaml
except ImportError:  # pragma: no cover
    sys.stderr.write("[WARN] PyYAML 未安装，使用内嵌默认 rubric。建议: pip install pyyaml\n")
    yaml = None

# openpyxl 为可选依赖：仅在 --excel 时用到
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    _HAS_OPENPYXL = True
except ImportError:  # pragma: no cover
    _HAS_OPENPYXL = False


# --------------------------------------------------------------------------- #
# 0. 默认 rubric（当外部 yaml 不可用时兜底；优先加载同目录 references/c4_rubric.yaml）
# --------------------------------------------------------------------------- #
DEFAULT_RUBRIC = {
    "required_deliverables": {
        "skill_doc": {"label_cn": "Skill 说明文档", "detection": {
            "filename_patterns": ["skill说明", "skill_doc", "skill-doc", "skill_description"],
            "content_signals": ["使用场景", "输入", "输出", "解决什么问题", "use case", "input", "output"],
            "preferred_extensions": [".md", ".pdf", ".docx"]}},
        "executable_content": {"label_cn": "可执行内容", "detection": {
            "filename_patterns": ["技能", "skill"],
            "content_signals": ["```", "def ", "class ", "import ", "---\nname:"],
            "preferred_extensions": [".skill", ".py", ".md", ".zip"]}},
        "demo": {"label_cn": "Demo", "detection": {
            "filename_patterns": ["demo", "演示", "截图", "screen", "录屏"],
            "preferred_extensions": [".mp4", ".mov", ".webm", ".png", ".jpg", ".gif"]}},
        "teaching_doc": {"label_cn": "教学说明", "detection": {
            "filename_patterns": ["教学说明", "教学", "tutorial", "teaching", "上手指南", "quickstart", "how-to"],
            "content_signals": ["上手", "步骤", "常见坑", "安装", "注意事项", "getting started", "step by step"],
            "preferred_extensions": [".md", ".pdf", ".docx"]}},
        "ai_log": {"label_cn": "AI 日志", "detection": {
            "filename_patterns": ["AI日志", "AI_log", "ai-log", "ai_日志"],
            "content_signals": ["使用的 AI", "AI 工具", "prompt", "迭代次数", "迭代", "ChatGPT", "Claude", "DeepSeek"],
            "preferred_extensions": [".md", ".pdf", ".docx"]}},
    },
    "quality_criteria": {
        "reusable": {"label_cn": "可复用", "weight": 0.25, "negative_signals": [
            "/Users/", "/home/", "C:\\", "我的电脑", "my_secret", "api_key = ", "apikey = "]},
        "executable": {"label_cn": "可执行", "weight": 0.25},
        "verifiable": {"label_cn": "可验证", "weight": 0.25},
        "clear_io": {"label_cn": "IO 明确", "weight": 0.25,
                     "io_pattern": "输入.*输出|input.*output|接受.*返回|给定.*得到"},
    },
    "scoring": {"completeness": {}, "quality_per_criterion": {},
                "composite": {"completeness_weight": 0.4, "quality_weight": 0.6}},
}


# --------------------------------------------------------------------------- #
# 1. 文件内容读取（按扩展名选择解析器，UTF-8 为主、GBK 回退）
# --------------------------------------------------------------------------- #
def _read_bytes(path, max_bytes=8_000_000):
    if os.path.getsize(path) > max_bytes:
        return None  # 超大文件：仅做文件名匹配
    with open(path, "rb") as f:
        return f.read()


def _decode(raw):
    if raw is None:
        return ""
    for enc in ("utf-8", "gbk", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def read_text(path):
    """返回文件的纯文本内容；无法解析的二进制返回 ''。"""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in (".md", ".txt", ".py", ".yaml", ".yml", ".json", ".csv"):
            return _decode(_read_bytes(path))
        if ext == ".skill" or ext == ".zip":
            return _read_skill_zip(path)
        if ext == ".pdf":
            return _read_pdf(path)
        if ext in (".docx", ".pptx"):
            return _read_office(path)
    except Exception as e:  # 解析失败不致命，降级为文件名匹配
        sys.stderr.write(f"[WARN] 读取失败 {path}: {e}\n")
    return ""


def _read_skill_zip(path):
    """读取 .skill/.zip 内的 SKILL.md 与全部文本文件内容。"""
    out = []
    try:
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if n.lower().endswith(("skill.md", ".md", ".txt", ".py")):
                    try:
                        out.append(_decode(z.read(n)))
                    except Exception:
                        pass
    except zipfile.BadZipFile:
        return ""
    return "\n".join(out)


def _read_pdf(path):
    try:
        from pypdf import PdfReader
    except ImportError:
        return ""
    try:
        r = PdfReader(path)
        return "\n".join((p.extract_text() or "") for p in r.pages)
    except Exception:
        return ""


def _read_office(path):
    try:
        from docx import Document
    except ImportError:
        return ""
    try:
        d = Document(path)
        return "\n".join(par.text for par in d.paragraphs)
    except Exception:
        return ""


# --------------------------------------------------------------------------- #
# 2. 文件采集与作者识别
# --------------------------------------------------------------------------- #
C4_MARK = "_C4_"
MEDIA_EXT = {".mp4", ".mov", ".webm", ".png", ".jpg", ".jpeg", ".gif"}


def scan_folder(folder):
    """递归扫描，返回 (submissions, non_c4)。
    submissions: {author: [file_info, ...]}；non_c4: [file_info, ...]
    file_info: {path, name, ext, rel, dir, size}
    """
    all_files = []
    for root, _dirs, files in os.walk(folder):
        for fn in files:
            p = os.path.join(root, fn)
            all_files.append({
                "path": p, "name": fn, "ext": os.path.splitext(fn)[1].lower(),
                "rel": os.path.relpath(p, folder), "dir": root, "size": os.path.getsize(p),
            })

    # 第一遍：识别含 _C4_ 标记的文件及其作者、所在子目录
    author_by_subdir = {}
    submissions = {}
    non_c4 = []
    for fi in all_files:
        if C4_MARK in fi["name"]:
            author = fi["name"].split(C4_MARK)[0].strip()
            submissions.setdefault(author or "Unknown", []).append(fi)
            # 记录该文件所在子目录，便于把同目录的 demo 媒体归入作者
            rel_dir = os.path.dirname(fi["rel"])
            if rel_dir and rel_dir not in ("", "."):
                author_by_subdir.setdefault(rel_dir, author or "Unknown")
        elif C4_MARK in fi["rel"]:
            # 父目录含 _C4_，例如 Author_C4_xxx/ 子文件夹
            part = [s for s in fi["rel"].split(os.sep) if C4_MARK in s]
            author = part[0].split(C4_MARK)[0].strip() if part else "Unknown"
            submissions.setdefault(author or "Unknown", []).append(fi)
        else:
            non_c4.append(fi)

    # 第二遍：把含 _C4_ 的子目录里的"无名"媒体文件(如 demo.mp4)归入作者
    for fi in non_c4[:]:
        rel_dir = os.path.dirname(fi["rel"])
        if rel_dir in author_by_subdir and fi["ext"] in MEDIA_EXT:
            submissions.setdefault(author_by_subdir[rel_dir], []).append(fi)
            non_c4.remove(fi)

    return submissions, non_c4


def extract_author_from_content(author_files):
    """文件名/子目录都无法识别时的内容回退：在 .md 中找作者字段。"""
    for fi in author_files:
        if fi["ext"] in (".md", ".txt"):
            txt = read_text(fi["path"])
            m = re.search(r"(?:作者|姓名|author)\s*[:：]\s*([^\n]{1,30})", txt, re.I)
            if m:
                return m.group(1).strip()
    return "Unknown"


# --------------------------------------------------------------------------- #
# 3. 完整性检查（5 必需文件）
# --------------------------------------------------------------------------- #
def _file_matches(fi, det):
    name_l = fi["name"].lower()
    if any(p.lower() in name_l for p in det.get("filename_patterns", [])):
        return True
    if any(name_l.endswith(e) for e in det.get("preferred_extensions", [])):
        # 仅凭扩展名还不够，需内容信号佐证（避免把所有 .md 都当 skill_doc）
        sigs = det.get("content_signals", [])
        if sigs:
            txt = read_text(fi["path"]).lower()
            if any(s.lower() in txt for s in sigs):
                return True
    return False


def completeness_check(author_files, rubric):
    results = {}
    for key, spec in rubric["required_deliverables"].items():
        det = spec["detection"]
        matched = [fi for fi in author_files if _file_matches(fi, det)]
        results[key] = {
            "label": spec["label_cn"],
            "present": bool(matched),
            "matched_files": [fi["name"] for fi in matched],
        }
    present = sum(1 for v in results.values() if v["present"])
    if present >= 5:
        status = "✅ 齐全"
    elif present >= 3:
        status = "⚠️ 部分缺失"
    else:
        status = "❌ 严重缺失"
    return results, present, status


# --------------------------------------------------------------------------- #
# 4. 质量评审（四条件）— 规则引擎 + 可选 LLM 深审占位
# --------------------------------------------------------------------------- #
def _has(txt, *subs):
    t = txt.lower()
    return any(s.lower() in t for s in subs)


def _check_reusable(txt, code_files, rubric):
    neg = rubric["quality_criteria"]["reusable"].get("negative_signals", [])
    items = {}
    items["有安装说明"] = _has(txt, "安装", "install", "pip install", "npm install", "setup", "配置步骤")
    items["无硬编码绝对路径"] = not any(n.lower() in txt.lower() for n in neg)
    items["列出环境要求"] = _has(txt, "环境要求", "requirements", "dependencies", "python", "node", "版本", "运行时")
    items["平台无关/注明平台"] = _has(txt, "兼容", "compatible", "跨平台", "windows", "macos", "linux", "平台")
    return items


def _check_executable(author_files, txt, rubric):
    items = {}
    items["含可运行代码/prompt/workflow"] = bool(
        re.search(r"```", txt) or _has(txt, "def ", "class ", "import ", "workflow", "pipeline", "步骤"))
    # .skill 包结构是否有效
    skill_valid = None
    for fi in author_files:
        if fi["ext"] in (".skill", ".zip"):
            try:
                with zipfile.ZipFile(fi["path"]) as z:
                    names = [n.lower() for n in z.namelist()]
                    skill_valid = any(n.endswith("skill.md") for n in names)
            except zipfile.BadZipFile:
                skill_valid = False
    items[".skill 包结构有效(如适用)"] = (skill_valid is True) or (skill_valid is None and _has(txt, "---", "name:"))
    # 语法检查（.py 用 py_compile；SKILL.md frontmatter 用 yaml）
    syntax_ok = True
    checked = False
    for fi in author_files:
        if fi["ext"] == ".py":
            checked = True
            try:
                import py_compile
                py_compile.compile(fi["path"], doraise=True)
            except Exception:
                syntax_ok = False
    items["无明显语法错误"] = (not checked) or syntax_ok
    # YAML frontmatter
    items["含 YAML frontmatter(如 SKILL.md)"] = bool(re.search(r"^---\s*\n.*?name\s*:", txt, re.S | re.M)) or _has(txt, "---\nname:")
    return items


def _check_verifiable(txt, has_demo, rubric):
    items = {}
    items["有测试用例/示例"] = _has(txt, "测试", "test", "example", "示例", "用例", "demo")
    items["定义预期输出"] = _has(txt, "预期", "expected", "输出格式", "返回格式", "结果应为", "期望")
    items["Demo 展示真实结果"] = has_demo or _has(txt, "截图", "录屏", "运行结果", "效果", "输出示例")
    items["成功/失败标准清晰"] = _has(txt, "成功", "失败", "判定", "通过", "assert", "校验", "验证")
    return items


def _check_clear_io(txt, rubric):
    pat = rubric["quality_criteria"]["clear_io"].get("io_pattern",
                                                     r"输入.*输出|input.*output|接受.*返回|给定.*得到")
    items = {}
    items["有『输入X，输出Y』一句话"] = bool(re.search(pat, txt, re.I))
    items["指定输入类型/格式"] = _has(txt, "输入类型", "input type", "接受", "入参", "参数")
    items["指定输出类型/格式"] = _has(txt, "输出类型", "output type", "返回", "出参", "产物")
    items["注明边界/异常输入"] = _has(txt, "边界", "异常", "edge case", "容错", "报错", "非法")
    return items


def _rate(items):
    """对齐 c4_rubric.yaml：✅=2+ 项满足, ⚠️=1 项, ❌=0 项。"""
    n = sum(1 for v in items.values() if v)
    if n >= 2:
        return "✅", 1.0, n
    if n == 1:
        return "⚠️", 0.5, n
    return "❌", 0.0, n


def quality_eval(author_files, rubric):
    txt = "\n".join(read_text(fi["path"]) for fi in author_files)
    has_demo = any(fi["ext"] in MEDIA_EXT for fi in author_files) or _has(txt, "demo", "演示")
    crit = {
        "reusable": _check_reusable(txt, author_files, rubric),
        "executable": _check_executable(author_files, txt, rubric),
        "verifiable": _check_verifiable(txt, has_demo, rubric),
        "clear_io": _check_clear_io(txt, rubric),
    }
    out = {}
    for k, items in crit.items():
        rate, score, n = _rate(items)
        out[k] = {"label": rubric["quality_criteria"][k]["label_cn"], "rate": rate,
                  "score": score, "satisfied": n, "items": items}
    return out


# --------------------------------------------------------------------------- #
# 5. 报告生成
# --------------------------------------------------------------------------- #
def _composite(completeness_present, quality):
    c_score = completeness_present / 5.0
    q_score = sum(v["score"] for v in quality.values()) / 4.0
    w = 0.4
    return round(c_score * w + q_score * (1 - w), 3), round(c_score, 3), round(q_score, 3)


def build_markdown(folder, submissions, non_c4, rubric, top_n=3):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    lines = []
    lines.append("# C4 提交自动评审报告\n")
    lines.append(f"- 生成时间：**{now}**")
    lines.append(f"- 扫描路径：`{folder}`")
    total_authors = len(submissions)
    total_files = sum(len(v) for v in submissions.values())
    lines.append(f"- 识别提交：**{total_authors}** 位作者，**{total_files}** 个文件\n")
    lines.append("> 注：本报告由评审器对输入文件夹**自动生成**。输入无论是真实微信群同步目录、"
                 "手动下载目录还是其他含群文件的目录，处理逻辑完全一致；本仓库 `samples/` "
                 "内置 4 份不同质量水平的样例提交（夹具）用于自测与演示。\n")

    # 预计算
    rows = []
    for author, files in submissions.items():
        if author == "Unknown":
            author_real = extract_author_from_content(files)
        else:
            author_real = author
        comp, present, cstatus = completeness_check(files, rubric)
        qual = quality_eval(files, rubric)
        compo, c_score, q_score = _composite(present, qual)
        rows.append({
            "author": author_real, "files": files, "comp": comp, "present": present,
            "cstatus": cstatus, "qual": qual, "compo": compo, "c_score": c_score, "q_score": q_score,
        })
    # 排序：综合分降序，其次完整性
    rows.sort(key=lambda r: (r["compo"], r["present"]), reverse=True)

    avg_q = round(sum(r["q_score"] for r in rows) / len(rows), 2) if rows else 0
    n_complete = sum(1 for r in rows if r["present"] == 5)
    n_partial = sum(1 for r in rows if 3 <= r["present"] < 5)

    # ---- 一、班级总览 ----
    lines.append("## 一、班级总览\n")
    lines.append("| 指标 | 数值 |")
    lines.append("|------|------|")
    lines.append(f"| 总提交人数 | {total_authors} |")
    lines.append(f"| 完整提交（5/5） | {n_complete} |")
    lines.append(f"| 部分提交（3-4/5） | {n_partial} |")
    lines.append(f"| 平均质量分 | {avg_q}/4.0 |")
    dist = {}
    for r in rows:
        dist[r["cstatus"]] = dist.get(r["cstatus"], 0) + 1
    lines.append(f"| 完整性分布 | " + "，".join(f"{k} {v}" for k, v in dist.items()) + " |\n")

    # ---- 二、作者详情 ----
    lines.append("## 二、作者详情\n")
    for r in rows:
        versions = sorted(set(re.findall(r"_v(\d+)", " ".join(f["name"] for f in r["files"]))))
        vtxt = f"（版本：v{', v'.join(versions)}）" if versions else ""
        lines.append(f"### {r['author']} {vtxt}\n")
        lines.append("**完整性检查：**\n")
        lines.append("| 文件 | 状态 | 匹配文件 |")
        lines.append("|------|------|----------|")
        for key, v in r["comp"].items():
            st = "✅" if v["present"] else "❌"
            mf = "、".join(v["matched_files"]) if v["matched_files"] else "—"
            lines.append(f"| {v['label']} | {st} | {mf} |")
        lines.append("")
        lines.append("**质量评审（四条件）：**\n")
        lines.append("| 条件 | 评级 | 满足项/依据 |")
        lines.append("|------|------|-----------|")
        for key, v in r["qual"].items():
            detail = "；".join(f"{'✔' if ok else '✘'}{k}" for k, ok in v["items"].items())
            lines.append(f"| {v['label']} | {v['rate']} | {detail} |")
        lines.append("")
        # 改进建议
        suggestions = _suggestions(r)
        if suggestions:
            lines.append("**改进建议：**")
            for s in suggestions:
                lines.append(f"1. {s}")
            lines.append("")
        lines.append("---\n")

    # ---- 三、排名 ----
    lines.append("## 三、综合排名\n")
    lines.append("| 排名 | 作者 | 完整性 | 质量分 | 综合分 |")
    lines.append("|------|------|--------|--------|--------|")
    for i, r in enumerate(rows, 1):
        lines.append(f"| {i} | {r['author']} | {r['present']}/5 | {r['q_score']}/4.0 | {r['compo']} |")
    lines.append("")

    # ---- 四、全班改进建议 ----
    lines.append("## 四、全班改进建议\n")
    miss = {}
    for r in rows:
        for key, v in r["comp"].items():
            if not v["present"]:
                miss[v["label"]] = miss.get(v["label"], 0) + 1
    if miss:
        top = sorted(miss.items(), key=lambda x: -x[1])[0]
        lines.append(f"- 最常见缺失文件：**{top[0]}**（{top[1]} 人缺失）")
    weak = {}
    for r in rows:
        for key, v in r["qual"].items():
            if v["rate"] != "✅":
                weak[v["label"]] = weak.get(v["label"], 0) + 1
    if weak:
        topw = sorted(weak.items(), key=lambda x: -x[1])[0]
        lines.append(f"- 最弱质量维度：**{topw[0]}**（{topw[1]} 人未达 ✅）")
    lines.append("- 建议下次提交前，用本评审器对文件夹自检后再提交。")
    lines.append("")

    if non_c4:
        lines.append("## 附：扫描到的非 C4 文件\n")
        for fi in non_c4[:50]:
            lines.append(f"- `{fi['rel']}`")
        lines.append("")

    return "\n".join(lines)


def _suggestions(r):
    out = []
    for key, v in r["comp"].items():
        if not v["present"]:
            out.append(f"补齐缺失的「{v['label']}」")
    for key, v in r["qual"].items():
        if v["rate"] == "❌":
            out.append(f"「{v['label']}」维度基本未满足，需补充对应内容（见上方满足项明细）")
        elif v["rate"] == "⚠️":
            out.append(f"「{v['label']}」仅部分满足，建议补全："
                       + "、".join(k for k, ok in v["items"].items() if not ok))
    return out[:6]


def build_excel(path, submissions, rubric):
    if not _HAS_OPENPYXL:
        sys.stderr.write("[WARN] openpyxl 不可用，跳过 Excel 生成（pip install openpyxl）\n")
        return False
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "评审详表"
    hdr = ["作者", "完整性", "完整性状态", "可复用", "可执行", "可验证", "IO明确", "质量分", "综合分"]
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    rows = []
    for author, files in submissions.items():
        a = author if author != "Unknown" else extract_author_from_content(files)
        comp, present, cstatus = completeness_check(files, rubric)
        qual = quality_eval(files, rubric)
        compo, c_score, q_score = _composite(present, qual)
        rows.append([a, f"{present}/5", cstatus, qual["reusable"]["rate"], qual["executable"]["rate"],
                     qual["verifiable"]["rate"], qual["clear_io"]["rate"], q_score, compo])
    rows.sort(key=lambda x: x[-1], reverse=True)
    for row in rows:
        ws.append(row)
    wb.save(path)
    return True


# --------------------------------------------------------------------------- #
# 6. 入口
# --------------------------------------------------------------------------- #
def load_rubric(path):
    if path and os.path.exists(path) and yaml:
        with open(path, "r", encoding="utf-8") as f:
            return yaml.safe_load(f)
    return DEFAULT_RUBRIC


def main():
    ap = argparse.ArgumentParser(description="C4 技能提交自动评审器")
    ap.add_argument("--folder", required=True, help="包含 C4 提交的本地文件夹路径")
    ap.add_argument("--report", default=None, help="输出 Markdown 报告路径")
    ap.add_argument("--excel", default=None, help="输出 Excel 详表路径（可选）")
    ap.add_argument("--rubric", default=None, help="c4_rubric.yaml 路径（默认加载同目录 references/）")
    ap.add_argument("--top-n", type=int, default=3, help="（预留）仅展示前 N 名")
    args = ap.parse_args()

    if not os.path.isdir(args.folder):
        sys.stderr.write(f"[ERROR] 文件夹不存在: {args.folder}\n")
        return 2

    # 默认 rubric 路径：脚本同目录/references/c4_rubric.yaml
    if not args.rubric:
        cand = os.path.join(os.path.dirname(os.path.abspath(__file__)), "references", "c4_rubric.yaml")
        if os.path.exists(cand):
            args.rubric = cand
    rubric = load_rubric(args.rubric)

    submissions, non_c4 = scan_folder(args.folder)
    if not submissions:
        print(f"[INFO] 未在 {args.folder} 找到任何含 '_C4_' 标记的提交文件。")
        print("       请确认文件命名遵循 姓名拼音_C4_内容描述.扩展名，或放在含 _C4_ 的子目录中。")
        return 2

    md = build_markdown(args.folder, submissions, non_c4, rubric, args.top_n)
    if args.report:
        with open(args.report, "w", encoding="utf-8") as f:
            f.write(md)
        print(f"[OK] 报告已写入: {args.report}")
    else:
        print(md)

    if args.excel:
        if build_excel(args.excel, submissions, rubric):
            print(f"[OK] Excel 已写入: {args.excel}")

    return 0


if __name__ == "__main__":
    sys.exit(main())

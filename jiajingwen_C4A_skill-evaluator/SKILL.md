---
name: c4a-skill-evaluator
description: >
  C4A 技能提交自动评审器。扫描一个本地文件夹（WeChat 群文件同步目录 / 手动下载目录 / 任意含群文件的目录），
  自动识别每位作者的 C4 提交，检查 5 个必需文件是否齐全，按 C4 四条件（可复用 / 可执行 / 可验证 / IO 明确）
  做规则化质量评审，并生成结构化评审报告（班级总览 + 作者详情 + 综合排名 + 改进建议）+ 可选 Excel 详表。
  这是 wechat-doc-mapper 的"自动阅卷"升级版——把收发室变成阅卷系统。
  触发词：评审C4提交、检查技能提交、C4评审报告、evaluate C4 submissions、check C4 completeness、
  review skill submissions，或用户给出一个文件夹路径并希望评审/复查其中的技能文件。
---

# C4A 技能提交自动评审器 (Automated Skill Submission Evaluator)

## 这是什么

把 wechat-doc-mapper 的"扫描 → 分类 → 缺口分析"能力升级为"自动阅卷"：
不只是分信（谁交了、交了什么），而是**阅卷**（交齐没、质量如何、怎么改）。

## 架构选择（为什么是「规则为主 + 可选 LLM 深审」）

| 方案 | 选用 | 理由 |
|------|------|------|
| 纯规则（关键词/正则） | ✅ 默认 | 确定性强、可离线、可复现、零 API 成本；助教可反复跑、出同样结果 |
| 纯 LLM | ❌ | 不确定、慢、贵、结果难复核，不适合"评分标准清晰"的硬要求 |
| 规则 + LLM 混合 | ✅ 预留 | 规则快筛给确定性结论，LLM 仅在需语义深审时介入（脚本已留 `quality_eval` 扩展点） |

核心原则：**所有检测信号从 `references/c4_rubric.yaml` 加载**，判断逻辑不写死在代码里——改评分标准只改 yaml。

## 流水线

```
本地文件夹（含 _C4_ 文件）
   ↓ ① scan_folder     递归扫描 + 作者识别（文件名_C4_ → 子目录 → 内容回退）
   ↓ ② completeness_check  5 必需文件（skill_doc/executable/demo/teaching/ai_log）
   ↓ ③ quality_eval    四条件规则评分（✅/⚠️/❌）
   ↓ ④ build_report    班级总览 + 作者详情 + 排名 + 改进建议（+Excel）
```

## 输入

- 一个本地文件夹路径（无论文件怎么来：手动下载、WeChat 同步、导出工具，本技能都接受）
- 文件命名遵循 `姓名拼音_C4_内容描述.扩展名` 最佳；不规范时回退到子目录 / 内容识别

## 输出

- Markdown 评审报告（默认打印到 stdout，或用 `--report` 落盘）
- 可选 Excel 详表（`--excel`）
- 综合分 = 完整性×0.4 + 质量×0.6；质量分 = 四条件均值（✅=1.0, ⚠️=0.5, ❌=0.0）

## 使用方式

```bash
# 进入技能目录
cd jiajingwen_C4A_skill-evaluator

# 安装依赖（仅首次）
pip install pyyaml openpyxl

# 评审一个文件夹，报告落盘 + 生成 Excel
python c4a_evaluate.py --folder /path/to/C4_submissions \
                       --report ../jiajingwen_C4A_评审报告.md \
                       --excel ../jiajingwen_C4A_评审详表.xlsx
```

依赖说明：
- `pyyaml` 必需（加载评分信号）
- `openpyxl` 仅在使用 `--excel` 时必需
- `.pdf`/`.docx` 内容读取需 `pypdf` / `python-docx`（可选；缺失时自动降级为文件名匹配）

## 评审维度（可检测信号摘要）

**完整性（5 文件）**：见 `references/c4_rubric.yaml → required_deliverables`，文件名模式 + 内容信号双判。

**质量（四条件）**：见 `references/c4_rubric.yaml → quality_criteria`，每条检查项：
- 可复用：有安装说明、无硬编码绝对路径、列环境要求、平台无关/注明平台
- 可执行：含可运行代码/工作流、.skill 结构有效、无明显语法错误、含 YAML frontmatter
- 可验证：有测试/示例、定义预期输出、Demo 展示真实结果、成功失败标准清晰
- IO 明确：有"输入X输出Y"一句话、指定输入类型、指定输出类型、注明边界/异常

评级规则（对齐 rubric）：满足 ≥2 项 → ✅；恰 1 项 → ⚠️；0 项 → ❌。

## 边界情况

| 情况 | 处理 |
|------|------|
| 空 / 无 _C4_ 文件 | 提示命名规范，退出码 2 |
| 命名不规范 | 回退子目录名 → 内容识别 → 标记 Unknown |
| 超 50MB 文件 | 跳过内容分析，仅文件名匹配 |
| 二进制（.mp4 等） | 仅文件名匹配；同 _C4_ 子目录的媒体自动归入作者 |
| 同作者多版本 | 识别 `_v2`/`_v3`，按文件名分组（演示夹具中可体现） |

## 文件清单

```
jiajingwen_C4A_skill-evaluator/
├── SKILL.md                  # 本文件
├── c4a_evaluate.py           # 评审器主程序
├── references/
│   └── c4_rubric.yaml        # 机器可读评分标准（唯一事实来源）
└── samples/                  # 自测夹具（3+ 不同质量样例提交）
```

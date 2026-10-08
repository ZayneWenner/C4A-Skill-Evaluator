# C4A 技能提交自动评审 —— Skill 评审器

EduSeed 挑战「C4A 技能提交自动评审」提交材料。
challengeId: `ch-20260717031432-5jvqje` ｜ 截止：2026-12-31 23:59

## 挑战要做什么

把「收发室」升级成「阅卷系统」：给定一个含多个作者提交的本地文件夹，自动识别每位作者的 C4 提交、检查 5 个必需文件是否齐全，并按 C4 四条件（可复用 / 可执行 / 可验证 / IO 明确）做规则化质量评审，输出班级总览 + 作者详情 + 排名 + 改进建议。

## 交付物

| 文件 | 说明 |
|---|---|
| `jiajingwen_C4A_方案设计.md` | 方案设计：评审维度、评分算法、误判控制、可解释输出 |
| `jiajingwen_C4A_skill-evaluator/` | 可运行技能包（`SKILL.md` + `c4a_evaluate.py` + `references/c4_rubric.yaml` + `samples/`） |
| `jiajingwen_C4A_评审报告.md` | 用评审器对样例提交给出的评审报告（班级总览 + 作者详情 + 排名 + 建议） |
| `jiajingwen_C4A_评审详表.xlsx` | 样例 × 各维度逐项评分明细 |
| `jiajingwen_C4A_AI日志.md` | AI 协作日志 |
| `jiajingwen_C4A_拿来说明.md` | 交付物使用说明 |
| `jiajingwen_C4A_教学说明.md` | 教学说明 |
| `CHALLENGE.md` | 挑战原文（供评审对照） |

## 运行方式

```bash
cd jiajingwen_C4A_skill-evaluator
pip install pyyaml openpyxl

python c4a_evaluate.py --folder <含 _C4_ 文件的目录> \
                       --report ../jiajingwen_C4A_评审报告.md \
                       --excel  ../jiajingwen_C4A_评审详表.xlsx
```

- `pyyaml` 必需（加载评分信号）；`openpyxl` 仅 `--excel` 时必需；`pypdf` / `python-docx` 可选（缺失时降级为文件名匹配）。
- 自测夹具：`jiajingwen_C4A_skill-evaluator/samples/`。

## 设计要点

- **规则为主 + 可选 LLM 深审**：默认纯规则，确定、可离线、可复现、零 API 成本。
- **评分标准不写死在代码里**：所有检测信号从 `references/c4_rubric.yaml` 加载，改标准只改 yaml。
- **综合分** = 完整性 × 0.4 + 质量 × 0.6；质量分 = 四条件均值（✅=1.0 / ⚠️=0.5 / ❌=0.0）。

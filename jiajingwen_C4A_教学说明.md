# 贾静文_C4A_教学说明（评审器安装与使用）

> 配套技能包：`jiajingwen_C4A_skill-evaluator/`
> 适用人群：老师 / 助教 / C4 学员（想提交前自检的人）
> 一句话：把一个"装了 C4 提交的文件夹"拖进来，吐出一份可复核的评审报告。

---

## 一、它能做什么

给定一个本地文件夹（WeChat 群文件同步目录、手动下载目录、或任何含群文件的目录），自动：

1. 识别每位作者的 C4 提交并分组；
2. 检查 5 个必需文件是否齐全（Skill 说明 / 可执行内容 / Demo / 教学说明 / AI 日志）；
3. 按 C4 四条件做质量评审（可复用 / 可执行 / 可验证 / IO 明确），每项给 ✅/⚠️/❌ + 依据；
4. 生成结构化报告：班级总览 + 作者详情 + 综合排名 + 改进建议（+ 可选 Excel）。

---

## 二、安装

### 环境要求
- Python 3.10+（已在 3.13 验证）
- 操作系统无关（Windows / macOS / Linux 均可）

### 步骤
```bash
# 1. 进入技能目录
cd jiajingwen_C4A_skill-evaluator

# 2. 安装依赖（仅两条，零其他依赖）
pip install pyyaml openpyxl
#   说明：pyyaml 必需（加载评分信号）；openpyxl 仅在使用 --excel 时需要

# 3.（可选）如需读取 .pdf / .docx 正文做内容判断
pip install pypdf python-docx
#   缺失时自动降级为"仅文件名匹配"，不影响主流程
```

> 不依赖任何外部 API / 网络，**完全离线可运行**。

---

## 三、怎么用（输入 → 输出）

### 最简用法（报告打印到屏幕）
```bash
python c4a_evaluate.py --folder /path/to/C4_submissions
```

### 常用用法（报告落盘 + 生成 Excel 详表）
```bash
python c4a_evaluate.py \
    --folder /path/to/C4_submissions \
    --report jiajingwen_C4A_评审报告.md \
    --excel jiajingwen_C4A_评审详表.xlsx
```

### 参数说明
| 参数 | 必填 | 说明 |
|------|------|------|
| `--folder` | ✅ | 包含 C4 提交的本地文件夹路径（递归扫描） |
| `--report` | ❌ | Markdown 报告输出路径；不填则打印到 stdout |
| `--excel` | ❌ | Excel 详表输出路径（需 openpyxl） |
| `--rubric` | ❌ | 自定义评分 yaml；默认加载 `references/c4_rubric.yaml` |
| `--top-n` | ❌ | 预留，当前展示全部作者 |

### 退出码
- `0`：正常完成
- `2`：未找到任何含 `_C4_` 的提交（提示检查命名规范）

---

## 四、输入文件该怎么命名（避免被误判）

技能优先按 `姓名拼音_C4_内容描述.扩展名` 识别作者与文件类型：
- `liming_C4_skill说明.md` → 作者 liming，识别为 Skill 说明
- `liming_C4_weather-skill.skill` → 可执行内容（.skill 包）
- `liming_C4_demo.png` → Demo
- `liming_C4_教学说明.md` → 教学说明
- `liming_C4_AI日志.md` → AI 日志

**不规范命名也能跑**：若文件名没按规范，技能会回退到「子目录名」→「文件正文里的 作者：xxx 字段」→ 标记 Unknown 待人工确认。**Demo 视频/截图即使没写 _C4_，只要放在含 _C4_ 文件的子目录里，也会自动归入对应作者。**

---

## 五、输出示例（节选）

```
# C4 提交自动评审报告
- 识别提交：4 位作者，16 个文件

## 一、班级总览
| 总提交人数 | 4 |
| 完整提交（5/5） | 2 |
| 部分提交（3-4/5） | 1 |

## 三、综合排名
| 排名 | 作者 | 完整性 | 质量分 | 综合分 |
| 1 | chenhao | 5/5 | 1.0/4.0 | 1.0 |
| 2 | liming   | 5/5 | 1.0/4.0 | 1.0 |
| 3 | wangfang | 3/5 | 0.625   | 0.615 |
| 4 | zhaolei  | 1/5 | 0.25    | 0.23 |
```

---

## 六、常见坑 / FAQ

| 现象 | 原因 | 解决 |
|------|------|------|
| 提示"未找到任何 C4 文件" | 文件名/子目录都不含 `_C4_` | 按规范重命名，或把文件放进 `xxx_C4_xxx` 子目录 |
| 某文件被判 ❌ 但你觉得有 | 关键词信号没命中（如说明文档没写"使用场景/输入/输出"） | 在文档里补齐信号词；或改 `c4_rubric.yaml` 的信号 |
| 想调评分权重 | 默认 0.4/0.6 写在 rubric 的 scoring 段 | 改 yaml 即可，不必动代码 |
| PDF/DOCX 内容没被分析 | 没装 `pypdf`/`python-docx` | `pip install pypdf python-docx` |
| 想换评审维度 | 信号全在 `references/c4_rubric.yaml` | 直接编辑 yaml，代码自动生效 |

---

## 七、给学员的"提交前自检"建议

把你的提交文件夹跑一遍本评审器，按报告里的「改进建议」逐条补，再提交到群——能显著减少被打回的概率。这也是 C4A 的"助人者"价值：评审器本身就是最好的自查工具。

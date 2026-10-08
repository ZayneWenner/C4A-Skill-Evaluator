#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 4 份不同质量水平的样例 C4 提交（夹具 fixtures），用于自测评审器。
说明：本环境无法访问真实微信群，故以结构典型的样例提交代替真实群文件。
评审器对真实文件夹与对 fixtures 文件夹的处理完全一致（输入只是一个本地路径）。
"""
import os, zipfile

BASE = os.path.dirname(os.path.abspath(__file__))


def w(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def make_skill_zip(path, skill_md):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("SKILL.md", skill_md)


# ---------------- A. liming — 优秀 (5/5, 四条件✅) ----------------
liming_skill = """# 天气查询技能

## 使用场景
用户输入城市名，自动返回未来 3 天天气，用于出行决策。

## 输入与输出（一句话）
输入：城市名（字符串），输出：未来 3 天天气（JSON）。

## 安装
```bash
pip install requests
```
## 环境要求
Python 3.10+，依赖 requests。兼容 Windows / macOS / Linux。

## 测试示例
输入 `北京`，预期返回包含 temperature 字段的 JSON。
边界：城市名为空时返回错误提示。
"""
w(os.path.join(BASE, "liming_C4_skill说明.md"), liming_skill)

liming_skill_zip_md = """---
name: weather-skill
description: 输入城市名返回未来3天天气
---
# Weather Skill
## Workflow
1. 接受城市名
2. 调用天气 API
3. 返回 JSON

```python
def get_weather(city: str) -> dict:
    return {"city": city, "forecast": []}
```
"""
make_skill_zip(os.path.join(BASE, "liming_C4_weather-skill.skill"), liming_skill_zip_md)
w(os.path.join(BASE, "liming_C4_demo.png"), "PNG-PLACEHOLDER")  # 媒体仅文件名匹配
w(os.path.join(BASE, "liming_C4_教学说明.md"),
  "# 教学说明\n## 上手步骤\n1. 安装依赖 2. 运行示例。\n## 常见坑\nAPI key 需配置环境变量。\n## 安装注意事项\n请使用虚拟环境。")
w(os.path.join(BASE, "liming_C4_AI日志.md"),
  "# AI 日志\n使用的 AI：ChatGPT + Claude。\nprompt：生成天气技能骨架。\n迭代次数：3 次。")

# ---------------- B. wangfang — 部分 (3/5, 可复用❌ 因硬编码路径) ----------------
wangfang_skill = """# 文件整理技能
## 使用场景
整理桌面文件。
用户输入文件夹，输出整理后的结构。
参数：文件夹路径(字符串)
返回结果：整理报告。
注意：目前只在我的电脑上跑过，路径写死在 C:\\Users\\wangfang\\projects\\sort\\config.ini
"""
w(os.path.join(BASE, "wangfang_C4_skill说明.md"), wangfang_skill)
w(os.path.join(BASE, "wangfang_C4_mybot.py"),
  "def sort_files(folder):\n    # TODO: 实现\n    return []\n\nclass Sorter:\n    pass\n")
w(os.path.join(BASE, "wangfang_C4_demo.mp4"), "MP4-PLACEHOLDER")
# 故意缺失：教学说明、AI日志

# ---------------- C. zhaolei — 薄弱 (1/5, 多条件❌) ----------------
w(os.path.join(BASE, "zhaolei_C4_说明.md"),
  "# 我的技能想法\n使用场景：帮用户整理文件。\n这个工具还在开发中，暂时没有代码。\n")
w(os.path.join(BASE, "zhaolei_C4_笔记.txt"), "一些零散想法，还没想清楚怎么做。\n")

# ---------------- D. chenhao — 版本迭代 (v1→v2, 5/5, 四条件✅) ----------------
chen_v1 = "# 翻译技能 v1\n使用场景：翻译句子。\n输出结果。\n"  # 早期：缺安装/缺示例
w(os.path.join(BASE, "chenhao_C4_v1_skill说明.md"), chen_v1)
chen_v2 = """# 翻译技能 v2
## 使用场景
调用翻译模型把中文译成英文。

## 输入与输出（一句话）
输入：中文文本（字符串），输出：英文文本（字符串）。

## 安装
```bash
pip install openai
```
## 环境要求
Python 3.11，需 API key（环境变量）。跨平台。

## 测试示例
输入 `你好`，预期输出 `Hello`。
边界：空字符串返回原样。
"""
w(os.path.join(BASE, "chenhao_C4_v2_skill说明.md"), chen_v2)
w(os.path.join(BASE, "chenhao_C4_v2_executable.py"),
  "def translate(text: str) -> str:\n    return text  # 简化示例\n")
w(os.path.join(BASE, "chenhao_C4_v2_demo.png"), "PNG-PLACEHOLDER")
w(os.path.join(BASE, "chenhao_C4_v2_教学说明.md"),
  "# 教学说明\n## 上手步骤\n1. 配 key 2. 运行。\n## 常见坑\nkey 不要写进代码。\n")
w(os.path.join(BASE, "chenhao_C4_v2_AI日志.md"),
  "# AI 日志\n使用的 AI：DeepSeek。\nprompt：写翻译技能。\n迭代次数：2 次。")

print("samples generated under:", BASE)

# 天气查询技能

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

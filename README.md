# 宏观框架分析

本仓库用于沉淀宏观日报、TREE 宏观分析表和相关自动化更新流程。日常数据通过 WPS/Wind 刷新，AI 或脚本再按固定 workflow 同步到 TREE 表，并完成趋势解析、趋势图和交付审计。

## 当前核心文件

| 文件 | 用途 |
| --- | --- |
| `20260706日报数据版V1.xlsx` | 当前最新日报表，由用户手动刷新后保存。 |
| `（0706V1）TREE宏观分析最新版.xlsx` | 当前最新 TREE 表，已同步 2026-07-06 日报数据。 |
| `20260624月报宏观数据_扩展2024历史_Wind刷新.xlsx` | 当前月报宏观数据源，用于历史分位、同比口径和颜色判断。 |
| `（5月28日）TREE宏观分析数据资讯_指标科普说明_单独版_V1.xlsx` | TREE 指标科普说明表。 |

## 重要入口

- `docs/project_info/AI_CONTEXT.md`：后续 AI 接手时优先阅读的项目上下文。
- `docs/AI_WORKFLOWS.md`：重复性任务总入口。
- `docs/ai_workflows/TREE_UPDATE_WORKFLOW.md`：TREE 更新流程手册。
- `tools/tree_daily_update/TREE_UPDATE_RULES.md`：TREE 更新计算、趋势图和审计规则。
- `tools/tree_daily_update/RUNBOOK.md`：TREE 例行更新低消耗执行手册。
- `logs/WORK_LOG.md`：详细工作记录。
- `logs/CHANGELOG.md`：面向项目版本的变更记录。

## 当前 TREE 更新口径

- 例行 TREE 更新默认使用用户手动刷新后的最新日报。
- 默认不自动刷新日报，也不更新日报前台趋势图，除非用户明确要求。
- TREE 趋势图只更新 TREE 自身图表缓存，不直接复制日报趋势图。
- `写入时间` 只记录本地写入/捕获日期，格式为 `yyyy-mm-dd`。
- 趋势解析必须基于日报或补充源的底层历史序列，不只看前台当前值或图形外观。

## GitHub 同步范围

本仓库默认跟踪项目规则、文档、工具脚本、日志和当前核心 Excel 文件。历史归档、临时输出、检查缓存、`node_modules/` 和大体量 `outputs/` 默认不纳入 Git，以避免仓库过大。

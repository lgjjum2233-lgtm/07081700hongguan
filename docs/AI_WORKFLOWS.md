# AI Workflows

本目录入口用于让后续 AI 模型快速读取项目中已经沉淀下来的重复工作流程。遇到同类任务时，先读本文件，再进入对应 workflow 手册，最后再读取详细日志和脚本规则。

## Workflow 入口

| 工作 | 入口文件 | 适用场景 |
| --- | --- | --- |
| 更新 TREE 表 | `docs/ai_workflows/TREE_UPDATE_WORKFLOW.md` | 用户要求按流程更新 TREE、同步最新日报数据、更新写入时间、趋势解析和 TREE 趋势图时使用。 |
| 输出宏观日报结论 | `docs/ai_workflows/MACRO_CONCLUSION_WORKFLOW.md` | 用户要求生成/复核宏观结论、更新辅助判断表、写入 TREE 结论或更新宏观日报 PPT 时使用。 |

## 权威来源

这些文件是 workflow 手册背后的详细依据：

- `tools/tree_daily_update/TREE_UPDATE_RULES.md`：TREE 更新的计算、颜色、趋势图、审计和交付规则。
- `logs/WORK_LOG.md`：历次 TREE 更新、规则修订、问题定位和验证记录。
- `logs/宏观日报结论输出工作日志_20260703.md`：宏观日报结论输出链条、辅助判断表、筛选规则和 PPT 样式规则。

## AI 使用顺序

1. 先读 `docs/project_info/AI_CONTEXT.md`，理解项目目标、当前文件口径和通用约定。
2. 再读本文件，判断任务属于 TREE 更新还是宏观结论输出。
3. 阅读对应 workflow 手册。
4. 对 TREE 更新，继续读取 `tools/tree_daily_update/TREE_UPDATE_RULES.md`。
5. 对结论输出，继续读取 `logs/宏观日报结论输出工作日志_20260703.md`。
6. 执行后更新 `logs/WORK_LOG.md`；如果改变了项目对外口径，也更新 `logs/CHANGELOG.md`。

## 边界

- workflow 手册是可执行工作说明，不替代原始日志。
- 涉及 Excel、PPT 或图表时，必须按对应 workflow 的验证清单执行。
- 不确定最新输入文件时，不要只按文件修改时间判断；必须复核 workbook 结构、关键 sheet、表头和底层数据可用性。

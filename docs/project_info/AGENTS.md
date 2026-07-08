# AI Agent 工作规则

此项目是一个基于 Excel/WPS/Wind 的宏观数据日报项目。AI 接手时请遵守以下规则。

## 优先阅读

1. `docs/project_info/AI_CONTEXT.md`
2. `docs/AI_WORKFLOWS.md`
3. `docs/project_info/README.md`
4. `docs/WORKBOOK_SPEC.md`
5. `logs/WORK_LOG.md`

## 固化 Workflow

- 更新 TREE 表前，先读 `docs/ai_workflows/TREE_UPDATE_WORKFLOW.md`，并遵守 `tools/tree_daily_update/TREE_UPDATE_RULES.md`。
- 输出宏观日报结论、辅助判断表或 PPT 结论前，先读 `docs/ai_workflows/MACRO_CONCLUSION_WORKFLOW.md`。
- 以上 workflow 是给 GitHub 和后续 AI 模型读取的稳定操作手册；详细历史仍以 `logs/WORK_LOG.md` 和对应工作日志为准。

## 文件版本

- 不覆盖原文件。
- 新输出文件必须带版本号。
- 小改动使用 `V1.1`、`V1.2`。
- 大版本或结构变化使用 `V2`、`V3`。
- 文件名示例：`20260526宏观分析汇总_matched_V1.xlsx`。

## Excel 操作

- Wind 数据刷新应通过 WPS 顶部 `Wind -> 刷新数据` 完成。
- 例行 TREE 更新默认使用用户手动刷新后的最新日报；不自动刷新日报，也不更新日报前台趋势图，除非用户明确要求。
- 如果通过脚本修改 `.xlsx`，必须保留原公式、图表和样式。
- 修改趋势图后，要用 WPS 打开验证图表对象和关键单元格。
- 匹配不到的数据填 `—`。
- TREE 中 A股、港股等股市指数的边际变化必须用日涨跌幅，不用点位差；后续脚本必须按表头适配。
- TREE `写入时间` 只记录本地写入/捕获日期，格式为 `yyyy-mm-dd`，不能写小时、分钟或秒。

## 脚本和临时文件

- 正式脚本放在 `scripts/`。
- 临时探索脚本、截图和中间文件放在 `codex_tmp/`，不要作为正式项目入口。
- 如需新增脚本，请在 `scripts/README.md` 记录用途、输入、输出和验证方式。

## 日志

每次完成有意义的修改后，更新：

- `logs/CHANGELOG.md`：记录面向项目的版本变化。
- `logs/WORK_LOG.md`：记录具体操作过程、输入文件、输出文件、验证结果和遗留问题。

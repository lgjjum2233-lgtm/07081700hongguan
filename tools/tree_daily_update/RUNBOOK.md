# TREE 例行更新低消耗流程

这个目录用于把 TREE / 日报更新流程固定下来，减少每次重新推理和大段读取工作簿带来的 token 消耗。

权威规则：

- `docs/ai_workflows/TREE_UPDATE_WORKFLOW.md`
- `tools/tree_daily_update/TREE_UPDATE_RULES.md`
- `logs/WORK_LOG.md`

## 标准流程

1. 识别最新权威 TREE 和最新手动刷新后的日报。
2. 复核日报 xlsx 完整性、底层 sheet 和近期非空观测。
3. 保持 TREE 输入结构：简化版仍简化，复杂版仍复杂，只更新目标主表。
4. 从日报或补充源的底层历史序列同步 TREE 当前数据、数据日期、环比、同比、同比增量和宏观颜色。
5. 只有当前数据或真实数据日期变化时，刷新 `写入时间`；默认只写日期级 `yyyy-mm-dd`。
6. 基于同一底层历史序列更新趋势解析文字。
7. 只更新 TREE 自身趋势图缓存，默认不刷新日报前台趋势图。
8. 清理 TREE 趋势图所有显示文字元素，包括数据标签、系列名、点位标签和图例。
9. 运行 xlsx、公式、日期、写入时间、趋势解析、图表 XML 和视觉渲染审计。
10. 审计通过后交付新版 TREE，并更新工作日志。

## 日报刷新规则

- 例行 TREE 更新默认不运行 `refresh_daily.ps1`。
- 用户已经在 WPS/Wind 手动刷新日报时，只识别和复核最新日报，然后进入 TREE 更新。
- 例行 TREE 更新默认不更新日报前台趋势图。
- 只有用户明确要求 Codex 刷新日报时，才运行 `refresh_daily.ps1` 或同类刷新脚本。

## 可复用脚本

- `sync_tree_new_indicators_from_daily.py`：TREE 数据和趋势图缓存同步的核心工具之一，已内置趋势图标签/图例清理。
- `update_tree_v3_latest_layout.py`：适配 V3 最新布局的数据同步与样式处理工具。
- `audit_workbooks.py`、`comprehensive_recheck.py`、`macro_change_history_audit.py`：历史审计工具，使用前需确认列位和当前 workflow 一致。
- `run_update.py`、`sync_tree.py`、`postprocess_daily.py` 是早期低消耗流程入口，复用前必须按当前 `TREE_UPDATE_RULES.md` 复核是否适配。

## 强制闸口

- 日报必备数据源必须存在。
- TREE 当前数据、数据日期、环比、同比、同比增量必须与底层历史序列口径一致。
- `写入时间` 只作为本地写入/捕获日期，不代表官方发布日期或 Wind 后台发布时间。
- TREE 主表不能有公式错误。
- 当日更新的边际变化颜色必须通过深度审计。
- TREE 趋势图必须是干净红色单折线，保持原锚点和尺寸。
- TREE 趋势图 XML 审计必须满足 `total_dLbls=0`、`total_legend=0`，有效来源行缓存点数大于 0。
- 必须完成代表性区域视觉预览；不能只凭 XML 审计声称视觉通过。

## 降低 token 的原则

- 默认只输出摘要，不展开几千行 JSON。
- 只有异常行才展开细节。
- 先读 workflow 和规则文件，再读取必要 sheet；不要反复全量展开工作簿。
- 最新日报已由用户手动刷新时，跳过自动刷新和日报趋势图修复，直接复核底层数据。
- 新增指标、颜色基准和白名单集中放在 `config.json` 或对应正式规则文件，避免每次重新判断。

# 工作区文件索引 V1

更新日期：2026-07-08

本索引用于快速判断项目根目录、归档目录和临时目录中各类文件的用途。整理原则是：根目录只放当前最新版主表、当前说明表和当前演示稿；项目说明进入 `docs/project_info/`；项目配置进入 `docs/project_config/`；旧版本、基准文件和参考材料进入 `archive/`；探索脚本、截图和中间文件进入 `codex_tmp/` 的任务分组。

## 根目录保留文件

| 文件 | 类型 | 用途 |
| --- | --- | --- |
| `20260706日报数据版V1.xlsx` | 当前日报 | 当前日报主文件，由用户手动完成 7月6日刷新并保存，作为 TREE 例行更新的数据源。 |
| `20260624月报宏观数据_扩展2024历史_Wind刷新.xlsx` | 当前月报数据源 | 当前月度宏观数据源，扩展 2024 历史并完成 6月24日 Wind 刷新。 |
| `（0706V1）TREE宏观分析最新版.xlsx` | 当前 TREE 文件 | 当前 TREE 精简交付版，基于 0703V2 按新流程同步 7月6日日报，保留可见主表和复核表，图表 XML 已审计通过。 |
| `（5月28日）TREE宏观分析数据资讯_指标科普说明_单独版_V1.xlsx` | TREE 说明表 | 只保留 `指标科普说明` 一个 sheet，打开即可直接查看指标解释。 |
| `宏观传导框架_资产版_V6.3_数据趋势更新_对齐修复.pptx` | 当前演示稿 | 当前资产版宏观传导框架演示稿，基于 V6 原稿更新数据、较上期、趋势图并修复小范围对齐。 |

## 说明与配置目录

| 目录 | 内容 |
| --- | --- |
| `docs/project_info/` | `README.md`、`AI_CONTEXT.md`、`AGENTS.md`，用于项目说明和 AI 接手规则。 |
| `docs/project_config/` | `.gitignore` 等项目配置归档。若后续正式启用 Git，需把 `.gitignore` 放回根目录才能自动生效。 |
| `tools/tree_daily_update/` | 新流程日常更新工具目录，包含日报刷新、图表缓存修正、TREE 同步和审计复核脚本。 |
| `node_modules/` | 本地 Node 工具依赖缓存，供 `tools/` 和结论表生成脚本使用。 |
| `outputs/macro_result/` | 宏观结果图片输出目录，存放从根目录移出的当前图片成果。 |
| `outputs/macro_result_ppt/` | 宏观日报结论 PPT 输出目录，最新根目录移出版本为 `0623宏观日报结论输出最新版_面级方向提示版.pptx`。 |
| `outputs/investment_strategy/` | AI 产业框架和投资策略类正式输出工作簿。 |
| `outputs/` | 其他正式输出文件，按主题子目录管理。 |

## 归档目录

| 目录 | 内容 |
| --- | --- |
| `archive/old_daily_versions/` | 旧版日报、日报基准文件、WPS 刷新中间版和财政赤字脉冲前一版日报。 |
| `archive/monthly_data_versions/` | 旧版月报宏观数据源、待刷新版本、图表缓存修正版和边际变化复核版。 |
| `archive/old_summary_versions/` | 旧版宏观分析汇总表、汇总表基准文件、原 `宏观分析汇总/` 文件夹中的历史表和已清空文件夹归档。 |
| `archive/tree_versions/` | TREE 原始输入、补码版、权重补充版、数据更新中间版和财政赤字脉冲前一版。 |
| `archive/audit_reports/` | TREE 和日报相关复核表、审计表、口径判定和着色校验输出。 |
| `archive/ppt_versions/` | 旧版和基准版资产配置框架演示稿，包括 V6 原稿、V6.1 和 V6.2。 |
| `archive/report_versions/` | 旧版宏观日报结论表、精简版、细分维度版和合并结论版。 |
| `archive/reference_inputs/` | `新增宏观数据.xlsx`、`日报V30新增(1).xlsx`、TREE 新框架参考表、5月29日指标补充跟踪表和宏观分析框架优化建议文档。 |

## 临时目录

| 目录 | 内容 |
| --- | --- |
| `codex_tmp/cache/` | 工具运行产生的缓存归档，如 `tree_daily_update_pycache_20260706/`。 |
| `codex_tmp/20260526_wps_refresh/` | WPS/Wind 点击、截图、窗口检查和打开验证相关临时脚本。 |
| `codex_tmp/20260526_add12_macro/` | 新增 12 个中国宏观指标和趋势图修复相关探索脚本。 |
| `codex_tmp/20260526_summary_match/` | 日报匹配宏观分析汇总表相关探索脚本。 |
| `codex_tmp/20260527_summary_v30_mapping/` | V30 指标与汇总表框架对应、验证相关临时文件。 |
| `codex_tmp/20260527_daily_v30_expand/` | V30 指标接入日报、扩展展示页和验证相关临时文件。 |
| `codex_tmp/20260528_v25_weight_update/` | TREE V2.5 权重补充、权重校准和渲染检查相关临时文件。 |
| `codex_tmp/20260528_data_update/` | 5月28日日报 WPS/Wind 刷新和 TREE 数据同步相关临时文件。 |
| `codex_tmp/20260528_fiscal_impulse/` | 财政赤字脉冲接入日报和 TREE 同步相关临时文件。 |
| `codex_tmp/20260528_percent_format/` | TREE 百分比格式检查和修正相关临时文件。 |
| `codex_tmp/20260528_indicator_guide/` | TREE 指标科普说明表生成脚本和校验相关临时文件。 |
| `codex_tmp/20260529_tree_update/` | 5月29日 WPS/Wind 刷新后 TREE 同步、权重理由清理、美股指数更新和校验相关临时文件。 |
| `codex_tmp/20260529_tree_trends/` | TREE 趋势图同步脚本和 5月29日/6月1日趋势图复跑相关临时文件。 |
| `codex_tmp/20260529_tree_weight_style/` | TREE 权重阈值红黑样式处理脚本和复跑相关临时文件。 |
| `codex_tmp/20260601_refresh_daily/` | 6月1日 WPS/Wind 刷新日报、TREE 同步和中间输出归档相关临时文件。 |
| `codex_tmp/20260601_tree_text_audit/` | 6月1日 TREE 权重理由、定义、定义降维、怎么用相关校订脚本、审计 JSON 和中间输出。 |
| `codex_tmp/20260603_tree_update/` | 6月3日用户放入新版 TREE 的 Wind 刷新同步、无图版、趋势图重建、权重样式和校验预览相关临时文件。 |
| `codex_tmp/tree_code_fill/` | TREE 指标代码补充和当前数据同步的临时工作区。 |
| `codex_tmp/failed_outputs/` | 失败或废弃的中间 Excel 输出，保留用于追溯。 |

## 使用建议

- 每日更新优先从根目录的当前日报开始。
- 新生成的正式 Excel 文件仍放根目录，文件名必须带版本号。
- 确认新版本稳定后，把上一轮旧版本移入 `archive/old_daily_versions/`、`archive/old_summary_versions/` 或 `archive/tree_versions/`，根目录只留最新版。
- 参考输入默认放在 `archive/reference_inputs/`；复跑历史脚本前先确认脚本是否需要调整输入路径。
- 临时探索脚本继续放 `codex_tmp/`，建议按日期和任务新建子目录。
- 正式可复用脚本只放 `scripts/`，新增正式脚本后同步更新 `scripts/README.md`。

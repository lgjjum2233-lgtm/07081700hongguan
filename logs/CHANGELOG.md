# 更新日志

## 2026-07-08

### Changed

- 将 TREE `写入时间` 规则从分钟级改为日期级；后续默认只写 `yyyy-mm-dd`，不写小时、分钟或秒。
- 更新 `tools/tree_daily_update/TREE_UPDATE_RULES.md`、`docs/ai_workflows/TREE_UPDATE_WORKFLOW.md` 和近期 TREE 写入时间包装脚本，避免后续继续输出分钟级时间。
- 同步更新 `tools/tree_daily_update/RUNBOOK.md`、`docs/WORKBOOK_SPEC.md`、`docs/project_info/AI_CONTEXT.md`、`docs/project_info/README.md`、`docs/project_info/AGENTS.md` 和 `scripts/README.md`，清理旧列位和旧日报刷新流程表述。
- 新增根目录 `README.md`、`.gitignore` 和 `.gitattributes`，并同步到 GitHub 仓库 `lgjjum2233-lgtm/07081700hongguan`；GitHub 跟踪范围包括项目文档、工具脚本、日志和当前最新日报/TREE 等核心文件，默认排除历史归档、临时输出和检查缓存。
- 更新宏观日报结论输出 workflow，补充数据一致性、口径优先级和旧 `指标判断` 冲突处理规则，避免后续根据旧文字误判最新数据。
- 修正宏观日报结论输出工作日志的引用路径为 `logs/宏观日报结论输出工作日志_20260703.md`。

## 2026-07-06

### Added

- 生成 `（0706V1）TREE宏观分析最新版_数据同步.xlsx`，作为 7月6日 TREE 数据同步中间版。
- 生成 `（0706V1）TREE宏观分析最新版.xlsx`，作为 7月6日按新流程更新后的当前 TREE。
- 新增 `codex_tmp/tree_data_sync_0706v1.json`、`codex_tmp/trend_analysis_0706v1.json`、`codex_tmp/audit_tree_0706v1.json` 和 `codex_tmp/render_0706v1/`，用于记录本次同步、趋势解析、审计和视觉渲染结果。

### Changed

- 当前日报切换为 `20260706日报数据版V1.xlsx`。
- 当前 TREE 切换为 `（0706V1）TREE宏观分析最新版.xlsx`。
- 按新 TREE workflow 跳过自动日报刷新，直接使用用户手动更新后的日报作为底层数据源。
- TREE 主表同步更新 179 行，更新 TREE 趋势图 178 个，趋势解析更新 185 行，`写入时间` 刷新 41 行。
- 更新 `docs/project_info/README.md`、`docs/project_info/AI_CONTEXT.md`、`docs/WORKSPACE_INDEX_V1.md` 和 `docs/WORKBOOK_SPEC.md`，同步 7月6日当前文件口径。
- 按既有工作区规则整理 7月1日到 7月6日新增日报、TREE 中间版、检查文件和宏观日报结论输出工作日志。
- 将旧日报移入 `archive/old_daily_versions/`，旧 TREE 和数据同步中间版移入 `archive/tree_versions/`，检查 `ndjson` 移入 `archive/audit_reports/`。
- 将 `宏观日报结论输出工作日志_20260703.md` 移入 `logs/`，将新生成的 `tools/tree_daily_update/__pycache__/` 移入 `codex_tmp/cache/tree_daily_update_pycache_20260706/`。

## 2026-07-03

### Added

- 新增 `docs/AI_WORKFLOWS.md`，作为 GitHub/AI 可读的重复工作流程总入口。
- 新增 `docs/ai_workflows/TREE_UPDATE_WORKFLOW.md`，固化 TREE 表更新流程、输入识别、趋势图规则、审计清单和交付说明。
- 新增 `docs/ai_workflows/MACRO_CONCLUSION_WORKFLOW.md`，固化宏观日报结论输出流程、辅助判断表、两轮筛选、反向证据检查、PPT 展示和标准句式。
- 新增 `docs/superpowers/specs/2026-07-03-ai-workflows-design.md` 和 `docs/superpowers/plans/2026-07-03-ai-workflows-docs.md`，记录本次文档化设计与实施计划。

### Changed

- 更新 `docs/project_info/AI_CONTEXT.md` 和 `docs/project_info/AGENTS.md`，将 `docs/AI_WORKFLOWS.md` 加入后续 AI 接手阅读顺序。
- 更新 `logs/WORK_LOG.md`，记录本次 workflow 文档固化过程和验证结果。

## 2026-06-30

### Changed

- 按既有工作区规则整理 6月21日到 6月30日新增日报、月报数据源、TREE、宏观日报结论 PPT、策略类工作簿和检查文件。
- 当前日报切换为 `20260630日报数据版V1_Wind刷新_最终版_图表缓存修正.xlsx`。
- 当前月报宏观数据源保留为 `20260624月报宏观数据_扩展2024历史_Wind刷新.xlsx`。
- 当前 TREE 切换为 `（0630V3）TREE宏观分析最新版.xlsx`。
- 将旧日报移入 `archive/old_daily_versions/`，旧月报数据源移入 `archive/monthly_data_versions/`，旧 TREE 移入 `archive/tree_versions/`。
- 将根目录宏观日报结论 PPT 移入 `outputs/macro_result_ppt/`，策略类工作簿移入 `outputs/investment_strategy/`，检查 `ndjson` 移入 `archive/audit_reports/`。
- 将旧 TREE 宏观日报结论表移入 `archive/report_versions/`。
- 将新生成的 `tools/tree_daily_update/__pycache__/` 移入 `codex_tmp/cache/tree_daily_update_pycache_20260630/`。
- 更新 `docs/WORKSPACE_INDEX_V1.md`、`docs/project_info/README.md`、`docs/project_info/AI_CONTEXT.md` 和 `docs/WORKBOOK_SPEC.md`，同步 6月30日当前文件口径。

## 2026-06-18

### Changed

- 按既有工作区规则整理 6月15日到 6月17日新增日报、月报数据源、TREE、宏观日报结论 PPT 和检查文件。
- 当前日报切换为 `20260617日报数据版V2_Wind刷新_PE修正_图表缓存修正.xlsx`。
- 当前月报宏观数据源切换为 `20260617月报宏观数据_扩展2024历史_Wind刷新.xlsx`。
- 当前 TREE 切换为 `（6月17日V2）TREE宏观分析最新版_PE估值修正版.xlsx`。
- 将旧日报移入 `archive/old_daily_versions/`，旧月报数据源移入 `archive/monthly_data_versions/`，旧 TREE 移入 `archive/tree_versions/`。
- 将根目录宏观日报结论 PPT 移入 `outputs/macro_result_ppt/`，检查 `ndjson` 和辅助表移入 `archive/audit_reports/`。
- 将 `宏观结果输出.png` 移入 `outputs/macro_result/`。
- 将新生成的 `tools/tree_daily_update/__pycache__/` 移入 `codex_tmp/cache/tree_daily_update_pycache_20260618/`。
- 更新 `docs/WORKSPACE_INDEX_V1.md`、`docs/project_info/README.md`、`docs/project_info/AI_CONTEXT.md` 和 `docs/WORKBOOK_SPEC.md`，同步 6月18日当前文件口径。

## 2026-06-14

### Changed

- 按既有工作区规则整理 6月12日和 6月13日 TREE 后续版本。
- 当前 TREE 切换为 `（6月13日）TREE宏观分析_总结论精简版_美国通胀表述修正.xlsx`。
- 将 6月12日 V15-V17、6月13日中间 TREE 版本移入 `archive/tree_versions/`。
- 将 6月12日复核表移入 `archive/audit_reports/`。
- 将 `宏观结果输出.png` 移入 `outputs/macro_result/`。
- 将新生成的 `tools/tree_daily_update/__pycache__/` 移入 `codex_tmp/cache/tree_daily_update_pycache_20260614/`。
- 更新 `docs/WORKSPACE_INDEX_V1.md`、`docs/project_info/README.md`、`docs/project_info/AI_CONTEXT.md` 和 `docs/WORKBOOK_SPEC.md`，同步 6月14日当前文件口径。

## 2026-06-12

### Added

- 新增 `archive/monthly_data_versions/`，用于归档旧版月报宏观数据源。
- 新增 `archive/audit_reports/`，用于归档 TREE 和日报相关复核表、审计表和口径校验输出。

### Changed

- 按既有工作区规则整理 6月10日、6月11日和 6月12日新增日报、月报数据源、TREE、结论表和复核表。
- 当前日报切换为 `20260611日报数据版V2_新增三项宏观指标_Wind刷新_图表缓存修正.xlsx`。
- 当前月报宏观数据源切换为 `20260611月报宏观数据_扩展2024历史_Wind刷新.xlsx`。
- 当前 TREE 切换为 `（6月12日V15）TREE宏观分析_当前数据红高绿低.xlsx`。
- 当前结论表切换为 `20260610_TREE宏观日报_趋势研判与资产含义.xlsx`。
- 旧日报、旧月报数据源、旧 TREE、旧结论表和复核表分别移入 `archive/old_daily_versions/`、`archive/monthly_data_versions/`、`archive/tree_versions/`、`archive/report_versions/` 和 `archive/audit_reports/`。
- 按实际存在文件修正工作区索引：当前根目录未保留宏观分析汇总表，TREE 指标科普说明表仅保留单独版。
- 将新生成的 `tools/tree_daily_update/__pycache__/` 移入 `codex_tmp/cache/tree_daily_update_pycache_20260612/`。
- 更新 `docs/WORKSPACE_INDEX_V1.md`、`docs/project_info/README.md`、`docs/project_info/AI_CONTEXT.md` 和 `docs/WORKBOOK_SPEC.md`，同步 6月12日当前文件口径。

## 2026-06-10

### Added

- 新增 `archive/report_versions/`，用于归档旧版 TREE 宏观日报结论表。

### Changed

- 按既有工作区规则整理 6月9日和 6月10日新增日报、TREE 和结论表版本。
- 当前日报切换为 `20260610日报数据版V1_Wind刷新_图表缓存修正.xlsx`，6月8日、6月9日和 6月10日未修正日报已移入 `archive/old_daily_versions/`。
- 当前 TREE 切换为 `（6月10日V1）TREE宏观分析.xlsx`，0608、6月9日和 6月10日中间 TREE 版本已移入 `archive/tree_versions/`。
- 当前结论表保留 `20260610_TREE宏观日报_最终结论_自然表述版.xlsx`，其他结论表版本已移入 `archive/report_versions/`。
- 将新生成的 `tools/tree_daily_update/__pycache__/` 移入 `codex_tmp/cache/tree_daily_update_pycache_20260610/`。
- 在工作区说明中补充 `node_modules/` 作为本地 Node 工具依赖缓存。
- 更新 `docs/WORKSPACE_INDEX_V1.md`、`docs/project_info/README.md`、`docs/project_info/AI_CONTEXT.md` 和 `docs/WORKBOOK_SPEC.md`，同步 6月10日当前文件口径。

## 2026-06-09

### Changed

- 按既有工作区规则整理 6月8日 TREE 后续修正版，根目录当前 TREE 切换为 `（0608）TREE宏观分析数据资讯V3_边际变化补色版.xlsx`。
- 将 `（0608）TREE宏观分析数据资讯V1_新流程Wind更新复核版.xlsx` 和 `（0608）TREE宏观分析数据资讯V2_美股边际变化修正版.xlsx` 移入 `archive/tree_versions/`。
- 将新生成的 `tools/tree_daily_update/__pycache__/` 移入 `codex_tmp/cache/tree_daily_update_pycache_20260609/`。
- 更新 `docs/WORKSPACE_INDEX_V1.md`、`docs/project_info/README.md`、`docs/project_info/AI_CONTEXT.md` 和 `docs/WORKBOOK_SPEC.md`，同步当前 TREE V3 口径。

## 2026-06-08

### Changed

- 按既有工作区规则整理 6月5日和 6月8日新增日报、TREE 文件，根目录当前正式文件切换到 6月8日版本。
- 当前日报切换为 `20260608日报数据版V1_新流程Wind刷新_图表缓存修正.xlsx`，旧版 6月4日、6月5日和 6月8日未修正日报已移入 `archive/old_daily_versions/`。
- 当前 TREE 切换为 `（0608）TREE宏观分析数据资讯V1_新流程Wind更新复核版.xlsx`，旧版 0604/0605 TREE 已移入 `archive/tree_versions/`。
- 在工作区说明中新增 `tools/tree_daily_update/`，作为新流程日常更新工具目录。
- 将 `tools/tree_daily_update/__pycache__/` 移入 `codex_tmp/cache/`。
- 此前被占用的旧文件 `20260605日报数据版V8_新流程Wind刷新_图表缓存修正.xlsx` 重试后已移入 `archive/old_daily_versions/`，对应 `~$` WPS 临时锁文件已消失。
- 更新 `docs/WORKSPACE_INDEX_V1.md`、`docs/project_info/README.md`、`docs/project_info/AI_CONTEXT.md` 和 `docs/WORKBOOK_SPEC.md`，同步 6月8日当前文件口径。

## 2026-06-04

### Added

- 新增 `archive/ppt_versions/`，用于归档旧版和基准版资产配置框架演示稿。

### Changed

- 按同一整理规则处理当天新增日报和 TREE 版本，根目录继续只保留当前正式文件。
- 当前日报切换为 `20260604日报数据版V3_宏观数据扩展_财政赤字脉冲_新增四指标_AH溢价指数_ETF净流入_趋势图样式统一_V5.xlsx`；当天前序 V1-V4、AH 溢价指数单独版和 6月3日新增四指标版已移入 `archive/old_daily_versions/`。
- 当前 TREE 切换为 `（0604）TREE宏观分析数据资讯_日报数据同步_新增指标_趋势图旧样式_保留原图XML_V2.xlsx`；0604 前序 TREE、0604 原始输入和 6月3日 TREE 已移入 `archive/tree_versions/`。
- 再次更新 `docs/WORKSPACE_INDEX_V1.md`、`docs/project_info/README.md`、`docs/project_info/AI_CONTEXT.md` 和 `docs/WORKBOOK_SPEC.md`，同步 6月4日最新文件口径。
- 按“根目录只保留当前最新版主表、当前说明表和当前演示稿”的规则整理根目录。
- 当前日报切换为 `20260603日报数据版V3_宏观数据扩展_财政赤字脉冲_新增四指标_V1.xlsx`，原 `20260603日报数据版V3_宏观数据扩展_财政赤字脉冲_WPS刷新_V1.xlsx` 已移入 `archive/old_daily_versions/`。
- 当前演示稿切换为 `宏观传导框架_资产版_V6.3_数据趋势更新_对齐修复.pptx`，V6 原稿、V6.1 和 V6.2 已移入 `archive/ppt_versions/`。
- 将根目录旧版 6月1日/6月2日 TREE、6月3日原始 TREE 输入、旧版日报和 `TREE宏观分析数据跟踪表_指标补充-5月29日(1) (1).xlsx` 分别移入 `archive/tree_versions/`、`archive/old_daily_versions/` 和 `archive/reference_inputs/`。
- 更新 `docs/WORKSPACE_INDEX_V1.md`、`docs/project_info/README.md`、`docs/project_info/AI_CONTEXT.md` 和 `docs/WORKBOOK_SPEC.md`，同步整理后的当前文件口径。

## 2026-06-03

### Added

- 生成 `20260603日报数据版V3_宏观数据扩展_财政赤字脉冲_WPS刷新_V1.xlsx`，作为用户登录 Wind 后通过 WPS 刷新的 6月3日快照。
- 生成 `（6月3日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_V1.xlsx`，基于用户放入的 `（6月3日）TREE宏观分析数据资讯 .xlsx` 同步刷新后日报缓存并重建趋势图。
- 生成 `（6月3日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_V1.1.xlsx`，修正 V1 中前台展示页旧值优先导致的 PMI 等底层序列漏更新。
- 生成 `（6月3日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_V1.2.xlsx`，将趋势图锚点从 O 列修正到 N 列对应行。

### Changed

- TREE 主表基于 6月3日刷新日报更新 104 行当前数据/数据日期/边际变化，未匹配 15 行；从无图版重建 105 张趋势图，趋势图未匹配 20 行，并恢复权重大于 2% 的红色加粗样式。
- 初版 TREE 文件为 `（6月3日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_V1.xlsx`，后续已由 V1.1 替代；当前日报切换为 `20260603日报数据版V3_宏观数据扩展_财政赤字脉冲_WPS刷新_V1.xlsx`。
- 当前 TREE 使用用户 6月3日提供版的旧列结构：H 为 `怎么用的（重点）`，J/K/L/M 分别为 `指标代码`、`当前数据`、`数据日期`、`边际变化`，N 为 `数据折线图`。
- 修正数据源优先级：当日报前台展示页仍为旧日期、底层 `宏观数据` 同代码序列已有更新时，优先采用底层新序列同步 K/L/M 和趋势图；PMI 组更新至 `2026-05-31`，短端利率/中债收益率/海外日度指标等同步至对应最新日期。
- 当前 TREE 文件切换为 `（6月3日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_V1.1.xlsx`；漏项审计仅保留 `港股日度成交额（亿港元）`，因底层序列单位与 TREE 亿港元口径不一致而暂不覆盖。
- 将存在 PMI 漏更新的旧 V1 移入 `archive/tree_versions/`，根目录保留修正后的 V1.1。
- 当前 TREE 文件切换为 `（6月3日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_V1.2.xlsx`；105 个趋势图锚点均位于 N 列，旧 V1.1 已移入 `archive/tree_versions/`。

## 2026-06-02

### Added

- 生成 `20260602日报数据版V3_宏观数据扩展_财政赤字脉冲_WPS刷新_V1.1.xlsx`，作为用户登录 Wind 后通过 WPS 刷新的 6月2日快照。
- 生成 `（6月2日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_指标说明校订_美国定义优化_V1.xlsx`，同步 6月2日刷新后日报缓存并重建趋势图。

### Changed

- TREE 主表基于 6月2日刷新日报更新 105 行当前数据/数据日期/边际变化，未匹配 17 行；从无图版重建 105 张趋势图，趋势图未匹配 22 行，并恢复权重大于 2% 的红色加粗样式。
- 当前 TREE 文件切换为 `（6月2日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_指标说明校订_美国定义优化_V1.xlsx`；当前日报切换为 `20260602日报数据版V3_宏观数据扩展_财政赤字脉冲_WPS刷新_V1.1.xlsx`。

## 2026-06-01

### Added

- 生成 `（6月1日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_指标说明校订_美国定义优化_V1.4.xlsx`，重点优化美国相关指标的 `定义` 和 `定义降维`。
- 生成 `宏观传导框架_资产版_V6.3_数据趋势更新_对齐修复.pptx`，以用户手工调整后的 `宏观传导框架_资产版_V6.pptx` 为底稿，仅更新数据、较上期、趋势图和小范围对齐。
- 生成 `宏观传导框架_资产版_V6.2_左侧指标校准_排版修复.pptx`，按 TREE 权重口径校准 PPT 左侧宏观数据区，并修复 V6.1 版排版拥挤和文字遮挡问题。
- 新增 `codex_tmp/20260601_ppt_left_check/left_indicator_check.md`，记录本次 PPT 左侧纳入指标清单和排除说明。
- 生成 `（6月1日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_指标说明校订_定义打磨_定义降维详写_V1.3.xlsx`，将 H 列 `定义降维` 详写为更完整的“指标是什么”描述。
- 生成 `（6月1日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_指标说明校订_定义打磨_V1.2.xlsx`，全量复核并打磨 TREE 主表 G 列 `定义`。
- 生成 `（6月1日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_指标说明校订_定义降维_V1.1.xlsx`，将 `定义降维` 重写为“指标是什么”的白话定义。
- 生成 `（6月1日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_指标说明校订_定义降维_V1.xlsx`，在 TREE 主表 `定义` 右侧新增 `定义降维` 列。
- 生成 `（6月1日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_指标说明校订_V1.2.xlsx`，全量校订 TREE 主表 `权重理由`、`定义`、`怎么用的（重点）` 三列。
- 生成 `20260601日报数据版V3_宏观数据扩展_财政赤字脉冲_WPS刷新_V1.xlsx`，作为用户登录 Wind 后通过 WPS 刷新的日报快照。
- 生成 `（6月1日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_V1.xlsx`，同步刷新后日报缓存，并重新生成趋势图和权重样式。
- 生成 `（6月1日）TREE宏观分析数据资讯_数据更新_V1.xlsx`，基于当前工作区最新 TREE 成品和最新 `20260529` 日报缓存同步主表数据。

### Changed

- TREE 主表第 78-132 行美国相关指标定义再次复核；重点澄清 TGA 是财政部在美联储的一般账户现金余额、ON RRP 是合格交易对手隔夜借给美联储的现金余额/工具、IORB 是美联储支付给银行准备金余额的管理利率、SOFR 是美国国债回购隔夜融资利率、SOFR-IORB 是两者差值。
- 当前 TREE 文件切换为 `（6月1日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_指标说明校订_美国定义优化_V1.4.xlsx`。
- TREE 主表第 6-132 行 `权重理由`、`定义`、`怎么用的（重点）` 三列覆盖 127 行、381 个单元格；修正美国隔夜逆回购利率、SOFR-IORB 利差、美国房屋销售、核心 PCE、财政赤字脉冲和产业分组等口径问题。
- G 列 `定义` 再次全量重写为更完整的专业口径描述，覆盖 G6:G132 共 127 条；美国零售销售、房屋销售、GDP、核心PCE、TIPS实际利率等底层代码口径已明确写入定义。
- 拆除 G 列目标区 5 组合并单元格，保证每个指标行都能独立读取定义。
- H 列 `定义降维` 再次详写，覆盖 H6:H132 共 127 条，补充统计对象、范围和口径边界，同时保持不写怎么用。
- TREE 主表在 `定义` 右侧新增 H 列 `定义降维`，写入 127 条降维解释；原 H:O 整体右移为 I:P，趋势图/绘图锚点同步右移。
- `定义降维` 口径修正为只描述“指标是什么”，不描述怎么用；当时 TREE 文件切换为 `（6月1日）TREE宏观分析数据资讯_Wind刷新同步_趋势图权重样式_指标说明校订_定义打磨_定义降维详写_V1.3.xlsx`，后续已由 V1.4 替代。
- `宏观传导框架_资产版_V6.3_数据趋势更新_对齐修复.pptx` 同步 6月1日 TREE 中左侧宏观指标及右侧固定资产配置指标，更新趋势图、当前值、较上期和关键结论文字；缺少 TREE 当前值的右侧交易总金额清为 `—`，避免沿用旧数。
- `宏观传导框架_资产版_V6.2_左侧指标校准_排版修复.pptx` 左侧宏观区按 `权重占比 >= 2%` 纳入指标；中国 `关注货币政策的走向` 作为结论文字不进指标表，美国 109-111 行产业基本面指标按用户确认不放入左侧数据区。
- 保持 PPT 右侧资产配置指标框架不变，仅重排和更新中国/美国左侧流动性、经济面、财政政策数据区。
- 按工作区整理规则归档旧版日报和旧版 TREE，根目录保留 6月1日当前日报、当前 TREE、当前汇总表、TREE 指标科普说明表和当前演示稿。
- 更新 `docs/WORKSPACE_INDEX_V1.md`、`docs/project_info/README.md`、`docs/project_info/AI_CONTEXT.md` 和 `docs/WORKBOOK_SPEC.md`，同步 6月1日当前文件口径。
- 刷新后日报同步更新 TREE 主表 108 行，未匹配 17 行；趋势图同步匹配 105 个图表。
- `融资融券余额` 已使用刷新后的 Wind 有效值，不再沿用上一有效值保护。
- 上证指数、深成指数、沪深300、创业板指、恒生指数、恒生科技指数、VIX、美元指数、COMEX黄金等行更新到刷新后日报缓存日期 `46171`。
- TREE `重点策略跟踪情况(V2.5)(1)` 本次同步更新 107 行，18 行仍无可确认日报来源或按既有保护规则保留旧值。
- A股、港股、美股 9 个股市指数继续使用日涨跌幅口径；当前定义降维版为 N 列，保持百分比格式。
- 保留 5月29日成品中的趋势图、权重样式和工作簿结构；目标页校验未发现新增公式错误。

## 2026-05-29

### Added

- 生成 `20260529日报数据版V3_宏观数据扩展_财政赤字脉冲_WPS刷新_V1.xlsx`，作为通过 WPS 顶部 `Wind -> 刷新数据` 刷新后的日报快照。
- 生成 `（5月29日）TREE宏观分析数据资讯_数据更新_权重理由清理_V1.xlsx`，同步刷新后日报缓存，并清理权重理由前缀。

### Changed

- 生成 `（5月29日）TREE宏观分析数据资讯_数据更新_权重理由清理_V1.2.xlsx`，将中国基本面政策面由约 `11.11%` 下调至 `8%`，流动性面由约 `38.89%` 上调至 `42%`。
- 中国基本面增长和价格利润分组保持不变，产业继续不纳入权重合计。
- 生成 `（5月29日）TREE宏观分析数据资讯_数据更新_权重理由清理_V1.1.xlsx`，按用户规则将中国基本面 `产业` 部分排除出权重合计。
- TREE 中国基本面第 44-49 行产业权重占比改为 `-`，第 6-43 行按原相对权重重新归一，合计为 100%。
- TREE `重点策略跟踪情况(V2.5)(1)` 共清理 127 行权重理由，删除 `权重较高，因为`、`权重中等，因为`、`权重较低，因为` 等前缀，只保留后面的解释部分。
- A股、港股、美股 9 个股市指数继续统一为日涨跌幅口径，M 列保持百分比格式；美股三大指数更新至 AP 发布的 2026-05-28 收盘数据。
- 发现刷新后日报中 `融资融券余额` 最新日 Wind 值为空、前台公式返回 0；TREE 中保留上一有效值，避免把缺值误写为真实 0。
- 将 5月28日旧 TREE 和旧日报移入 `archive/tree_versions/`、`archive/old_daily_versions/`，根目录保留 5月29日最新主表。

## 2026-05-28

### Added

- 生成 `（5月28日）TREE宏观分析数据资讯_指标科普说明_V1.xlsx`，作为独立说明版文件新增 `指标科普说明` sheet，覆盖 TREE 主表 122 个指标的定义、作用、跟踪方式、上涨/下跌含义和注意事项。
- 生成 `（5月28日）TREE宏观分析数据资讯_数据更新_财政赤字脉冲_V1.3.xlsx`，将 A股、港股、美股股市指数的边际变化从点位差改为日涨跌幅。
- 新增正式脚本 `scripts/update_tree_stock_index_pct_change.ps1`，用于后续 TREE 更新后自动把股票指数边际变化转换为涨跌幅。
- 生成 `（5月28日）TREE宏观分析数据资讯_数据更新_财政赤字脉冲_V1.2.xlsx`，按上证指数同类取数口径补齐 `标普500指数`、`纳斯达克指数` 和 `道琼工业指数`。
- 生成 `（5月28日）TREE宏观分析数据资讯_数据更新_财政赤字脉冲_V1.1.xlsx`，修正 TREE 中百分比数据的显示格式。
- 生成 `20260528日报数据版V3_宏观数据扩展_财政赤字脉冲_V1.xlsx`，在日报 `中国宏观` 新增 `财政赤字脉冲` 指标，并接入收入、支出、GDP 三条底层 Wind 代码。
- 生成 `（5月28日）TREE宏观分析数据资讯_数据更新_财政赤字脉冲_V1.xlsx`，将 `财政赤字脉冲` 同步进 TREE V2.5 主表。
- 生成 `20260528日报数据版V3_宏观数据扩展_WPS刷新_V1.xlsx`，作为 5月28日通过 WPS/Wind 刷新后的日报副本。
- 生成 `（5月28日）TREE宏观分析数据资讯_数据更新_V1.xlsx`，将刷新后日报数据同步到 V2.5 权重总表。
- 生成 `（5月28日）TREE宏观分析数据资讯_权重补充_V1.xlsx`，在 `重点策略跟踪情况(V2.5)(1)` 中补齐 127 行权重占比和权重理由。
- 生成 `（5月28日）TREE宏观分析数据资讯_权重补充_V1.2.xlsx`，将权重理由改为“权重较高/中等/较低 + 原因”的简明解释口径。
- 生成 `（5月28日）TREE宏观分析数据资讯_权重补充_V1.3.xlsx`，结合专业宏观/金融条件框架重新校准权重，并新增 `权重校验说明` sheet。
- 新增 `docs/WORKSPACE_INDEX_V1.md`，归纳根目录、`archive/` 和 `codex_tmp/` 的文件用途与保留规则。
- 新增 `archive/` 分类归档目录：`old_daily_versions/`、`old_summary_versions/`、`tree_versions/`、`reference_inputs/`。

### Changed

- 按日报前台单位列为 `%` 的口径，统一 TREE V2.5 主表中 46 行百分比指标的 K/M 列显示为 `0.00%`；`财政赤字脉冲` 修正为 `K27=-0.81%`、`M27=0.12%`。
- 按“同类表格归档、根目录只留最新版”的规则，将旧版日报、旧版 TREE、旧版汇总表和参考输入移入 `archive/` 对应分类目录。
- 根目录当前最新版主表更新为 `20260528日报数据版V3_宏观数据扩展_财政赤字脉冲_V1.xlsx`、`20260526宏观分析汇总_matched_V1.2.xlsx` 和 `（5月28日）TREE宏观分析数据资讯_数据更新_财政赤字脉冲_V1.3.xlsx`。
- 将 `.gitignore` 移入 `docs/project_config/`，将 `AGENTS.md`、`AI_CONTEXT.md`、`README.md` 移入 `docs/project_info/`，根目录进一步清理为只保留当前主表和文件夹。
- 将已清空的 `宏观分析汇总/` 文件夹移入 `archive/old_summary_versions/宏观分析汇总_empty_folder/`。
- 扩展日报 `宏观数据` 的 WSD 公式范围至 `E2:E146`，新增 `M0046168`、`M0046166`、`M0001395`，并用公式计算 `财政赤字脉冲 = (一般公共预算收入 - 一般公共预算支出) / 去年GDP`，前台按百分比口径展示。
- TREE 同步后可匹配指标由 104 行增至 105 行，未匹配行由 18 行降至 17 行；`财政赤字脉冲` 第 27 行写入代码 `M0046168-M0046166/M0001395`。
- 使用 WPS 顶部 `Wind -> 刷新数据` 刷新日报副本后，读取日报 `A股港股`、`中债`、`中国宏观`、`海外数据`、`宏观数据` 和 `指数走势` 的缓存值，同步更新 TREE V2.5 主表 K/L/M 列。
- 本次 TREE 数据同步共更新 104 行，18 行无可确认数据源继续填 `—`；权重列、权重理由、趋势图列和 `权重校验说明` 未改动。
- 在权重补充版中删除 `重点策略跟踪情况(V3)`，仅保留并更新 V2.5 主表；按 A 列顶层模块分别归一，使中国基本面、A股、港股、短债、中长债、美国经济基本面、美股、美债、黄金各自合计 100%。
- 清理权重理由中的重复权重表述，避免与“较高/中等/较低”标签冲突。
- V1.3 上调中国信用周期核心指标、美国金融条件与通胀核心指标、债券曲线指标和黄金实际利率/降息预期指标；下调低频、结构性、重复性或噪音较大的辅助指标。
- 将 `codex_tmp/` 顶层临时脚本按任务批次整理到 5 个子目录。
- 更新 `README.md`、`AI_CONTEXT.md`、`docs/WORKBOOK_SPEC.md`、`docs/WORKSPACE_INDEX_V1.md` 和 `scripts/README.md`，同步 5月28日最新文件、参考输入归档位置和复跑历史脚本注意事项。

## 2026-05-27

### Added

- 生成 `（补充代码）TREE宏观分析数据资讯_V1.2.xlsx`，将最新日报扩展表中的当前数据、数据日期、边际变化同步到 TREE 宏观分析数据资讯。
- 生成 `（补充代码）TREE宏观分析数据资讯_V1.1.xlsx`，基于最新日报扩展表继续补充 TREE 指标代码。
- 生成 `（补充代码）TREE宏观分析数据资讯_V1.xlsx`，在 `重点策略跟踪情况(V2.5)` 的 `指标代码` 列补充宏观代码。
- 生成 `20260526宏观分析汇总_matched_V1.1.xlsx`，在“中国基本面 / 流动性面 / 实体流动性”中补入 `新增外币贷款（同比多增）`。
- 为新增指标写入当前数据、数据日期、边际变化，并新增同风格 R 列趋势图。
- 生成 `20260526宏观分析汇总_matched_V1.2.xlsx`，完成 `日报V30新增(1).xlsx` 中 32 个指标与宏观分析汇总框架的对应。
- 在美国流动性框架中补充 6 个缺失指标：`美联储储备银行信贷-贷款`、`SOFR成交额`、`穆迪Aaa企业债收益率`、`银行信贷余额（非季调）`、`居民部门贷款总额`、`商业银行房地产贷款余额`。

### Changed

- 匹配日报 `A股港股`、`中债`、`中国宏观`、`海外数据` 和 `指数走势` 的缓存数据，更新 TREE 首张表 P/Q/R 列；本次未更新折线图列。
- 将日报 `A股港股`、`中债`、`中国宏观`、`海外数据` 和 `指数走势` 的可用代码纳入 TREE 补码来源，替换 7 个原为 `—` 的代码项。
- 以 `20260527日报数据版V3_宏观数据扩展_V1.2.xlsx` 的 `宏观数据` sheet 为主要代码源，补齐 90 个空白指标代码；在该 sheet 中无对应项的 31 个监测指标填 `—`。
- 参考 `（5月15日）TREE宏观分析汇总.xlsx` 与 `宏观分析框架优化建议.docx` 校验新增指标归类；避免重复添加已在框架中的社融、PMI、地产、利润类指标。
- 将 32 个 V30 新增指标对应的汇总表 I 列监测指标文字标为蓝色，新增行暂不更新 O/P/Q/R 数据。

### Current Files

- 日报：`20260527日报数据版V3_宏观数据扩展_V1.2.xlsx`
- 汇总表：`20260526宏观分析汇总_matched_V1.2.xlsx`
- TREE 补码表：`（补充代码）TREE宏观分析数据资讯_V1.2.xlsx`

## 2026-05-26

### Added

- 建立项目文档结构：`README.md`、`AI_CONTEXT.md`、`AGENTS.md`、`docs/`、`logs/`、`scripts/`。
- 新增 12 个中国宏观指标到日报。
- 新增 `新增宏观数据.xlsx`，保存从图片识别出的新增指标清单。
- 新增宏观分析汇总表匹配脚本。

### Changed

- 修复新增 12 个指标的趋势图，使其与原有趋势图风格一致。
- 将日报数据匹配到宏观分析汇总表，更新当前数据、数据日期、边际变化和趋势图。
- 统一后续文件命名规则：输出文件必须带 `V1`、`V1.1`、`V2` 等版本号。

### Current Files

- 日报：`20260526日报数据版V3_add12macro_趋势图修复.xlsx`
- 汇总表：`20260526宏观分析汇总_matched_V1.xlsx`

## 2026-05-27

### Added

- 新增日报扩展文件：`20260527日报数据版V3_宏观数据扩展_V1.2.xlsx`。
- 新增 `海外数据` sheet，用于承接美国/海外相关宏观指标。
- 新增脚本：`scripts/expand_daily_macro_v30.ps1`。

### Changed

- 将 `日报V30新增(1).xlsx` 中 18 条尚未存在的美国指标加入 `宏观数据`。
- 将 `宏观数据` 中未展示的指标补充到 `A股港股`、`中债`、`中国宏观` 和 `海外数据`。
- `宏观数据` 扩展到 143 行，图表部件总数增加到 160。

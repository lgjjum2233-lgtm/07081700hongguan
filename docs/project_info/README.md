# 宏观框架分析

这是一个用于跟踪 A 股、港股、中债、中国宏观及后续全球宏观指标的日报分析项目。项目核心是通过 WPS 表格中的 Wind 插件刷新指标数据，并把日报数据同步到宏观分析汇总表中，方便每日观察当前数据、数据日期、边际变化和趋势图。

## 当前核心文件

| 文件 | 用途 |
| --- | --- |
| `20260706日报数据版V1.xlsx` | 当前日报主文件，由用户手动完成 7月6日刷新并保存，作为 TREE 例行更新的数据源。 |
| `20260624月报宏观数据_扩展2024历史_Wind刷新.xlsx` | 当前月报宏观数据源，扩展 2024 历史并完成 6月24日 Wind 刷新。 |
| `（0706V1）TREE宏观分析最新版.xlsx` | 当前 TREE 精简交付版，基于 0703V2 按新流程同步 7月6日日报，保留可见主表和复核表。 |
| `（5月28日）TREE宏观分析数据资讯_指标科普说明_单独版_V1.xlsx` | 当前 TREE 指标科普说明表。 |
| `宏观传导框架_资产版_V6.3_数据趋势更新_对齐修复.pptx` | 当前资产版宏观传导框架演示稿。 |
| `outputs/macro_result_ppt/0623宏观日报结论输出最新版_面级方向提示版.pptx` | 最新从根目录移出的宏观日报结论 PPT 输出。 |
| `outputs/investment_strategy/` | AI 产业框架和投资策略类正式输出工作簿目录。 |
| `archive/reference_inputs/新增宏观数据.xlsx` | 从桌面两张宏观数据图片识别整理出的新增指标清单，作为参考输入归档。 |
| `archive/reference_inputs/日报V30新增(1).xlsx` | V30 新增指标清单，作为参考输入归档。 |
| `archive/old_daily_versions/` | 旧版日报和日报基准文件归档。 |
| `archive/old_summary_versions/` | 旧版宏观分析汇总表和汇总表基准文件归档。 |

后续所有新生成文件都应使用版本号命名，例如：

- `20260526宏观分析汇总_matched_V1.xlsx`
- `20260526宏观分析汇总_matched_V1.1.xlsx`
- `20260528日报数据版V3_V1.xlsx`

## 项目结构

```text
.
├── docs/
│   ├── WORKBOOK_SPEC.md
│   ├── WORKSPACE_INDEX_V1.md
│   ├── UPDATE_WORKFLOW.md
│   ├── GITHUB_SETUP.md
│   ├── project_info/
│   │   ├── README.md
│   │   ├── AI_CONTEXT.md
│   │   └── AGENTS.md
│   └── project_config/
│       └── .gitignore
├── logs/
│   ├── CHANGELOG.md
│   └── WORK_LOG.md
├── scripts/
│   ├── README.md
│   ├── add_12_macro_indicators.ps1
│   ├── fix_trends_openxml.ps1
│   ├── fix_trend_row_heights.ps1
│   ├── expand_daily_macro_v30.ps1
│   ├── match_daily_to_macro_summary.ps1
│   └── verify_wps_open_summary.ps1
├── tools/
│   └── tree_daily_update/
├── node_modules/
├── outputs/
│   ├── macro_result_ppt/
│   ├── investment_strategy/
│   └── macro_result/
├── archive/
│   ├── old_daily_versions/
│   ├── monthly_data_versions/
│   ├── old_summary_versions/
│   ├── tree_versions/
│   ├── audit_reports/
│   ├── ppt_versions/
│   ├── report_versions/
│   └── reference_inputs/
└── *.xlsx / *.pptx  # 仅保留当前最新版主表、说明表和当前演示稿
```

详细文件用途见 `docs/WORKSPACE_INDEX_V1.md`。

## 快速开始

1. 打开最新日报文件。
2. 在 WPS 顶部菜单点击 `Wind`。
3. 点击 `刷新数据`，等待 Wind 数据刷新完成。
4. 保存为新的日期和版本号文件。
5. 如需同步到宏观分析汇总表，运行 `scripts/match_daily_to_macro_summary.ps1` 或让 AI 按 `docs/project_info/AI_CONTEXT.md` 的说明操作。
6. 更新 `logs/CHANGELOG.md` 和 `logs/WORK_LOG.md`。

## 依赖环境

- Windows
- WPS Office 表格
- WPS 的 Wind 插件，并已登录可用账号
- PowerShell
- 可选：Git / GitHub Desktop / Git LFS

## 重要原则

- 不覆盖原始 Excel 文件，所有输出另存为带版本号的新文件。
- TREE 例行更新默认使用用户手动刷新后的最新日报；不自动刷新日报，也不更新日报前台趋势图，除非用户明确要求。
- 匹配不到的数据统一填 `—`。
- TREE 中 A股、港股等股市指数的边际变化统一使用日涨跌幅，即 `今日收盘 / 昨日收盘 - 1`，显示为百分比；后续脚本必须按表头定位。
- TREE 权重理由列只保留解释内容，不再保留 `权重较高，因为`、`权重中等，因为`、`权重较低，因为` 这类前缀。
- TREE 当前主表以后续交付版的表头为准；不要依赖早期固定列位，后续脚本必须按表头定位。
- TREE `写入时间` 只记录本地写入/捕获日期，格式为 `yyyy-mm-dd`，不是官方发布日期或 Wind 后台发布时间。
- TREE 中国基本面中 `产业` 部分不纳入权重合计，产业行权重占比用 `-` 表示，其余中国基本面指标重算至合计 100%。
- TREE 中国基本面权重当前口径：流动性面 `42%`、政策面 `8%`、增长 `27.78%`、价格利润 `22.22%`、产业 `-`。
- 修改公式、图表或指标映射后，需要用 WPS 打开验证。
- 临时文件放在 `codex_tmp/`，正式脚本放在 `scripts/`。
- 旧版本和参考输入放在 `archive/`，项目说明放在 `docs/project_info/`，根目录只保留当前最新版主表、当前说明表和当前演示稿。

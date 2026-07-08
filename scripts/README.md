# 脚本说明

这些脚本是当前项目中可复用的正式脚本。运行前请先确认输入文件和输出文件命名。

注意：部分历史脚本会在项目根目录按文件名模式自动查找输入文件。当前整理规则是根目录只保留最新版主表；旧版、基准文件和参考输入已进入 `archive/`。复跑历史脚本前，请先按 `docs/WORKSPACE_INDEX_V1.md` 确认输入路径，必要时调整脚本的查找目录。

## `match_daily_to_macro_summary.ps1`

用途：把日报数据匹配到宏观分析汇总表。

主要动作：

- 读取最新日报文件。
- 读取宏观分析汇总表。
- 更新汇总表 O/P/Q/R 列。
- 为匹配成功的指标生成趋势图。
- 匹配不到的指标填 `—`。

运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\match_daily_to_macro_summary.ps1
```

## `expand_daily_macro_v30.ps1`

用途：把 `archive/reference_inputs/日报V30新增(1).xlsx` 中的新指标接入日报。

主要动作：

- 将 V30 中尚未存在的指标追加到 `宏观数据`。
- 扩展 Wind WSD 公式引用范围。
- 把 `宏观数据` 中未展示的指标补充到 `A股港股`、`中债`、`中国宏观`。
- 新增 `海外数据` sheet，并放入美国/海外相关指标。
- 为新增展示行写入当前值、日期、变化、上期值、上期日期公式和趋势图。

运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\expand_daily_macro_v30.ps1
```

## `add_12_macro_indicators.ps1`

用途：把 `archive/reference_inputs/新增宏观数据.xlsx` 中的 12 个指标加入日报 `中国宏观` 页。

注意：该脚本依赖 WPS COM，即 `KET.Application`。

## `fix_trends_openxml.ps1`

用途：修复新增指标的趋势图引用和图表结构。

## `fix_trend_row_heights.ps1`

用途：修复新增指标行高，使趋势图显示尺寸与原有指标一致。

## `verify_wps_open_summary.ps1`

用途：用 WPS 打开汇总表并验证图表对象和关键单元格。

运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\verify_wps_open_summary.ps1
```

## `update_tree_stock_index_pct_change.ps1`

用途：把 TREE 主表中股票指数行的边际变化从点位差转换为日涨跌幅。

注意：这是早期 TREE 结构的辅助脚本，当前 TREE 更新应优先按 `docs/ai_workflows/TREE_UPDATE_WORKFLOW.md` 和 `tools/tree_daily_update/TREE_UPDATE_RULES.md` 执行；复用本脚本前必须按表头重新确认输出列。

适用行：

- A股指数：上证指数、深成指数、沪深300、创业板指。
- 港股指数：恒生指数、恒生科技指数。
- 美股指数：标普500指数、纳斯达克指数、道琼工业指数。

计算口径：

- `涨跌幅 = 点位差 / (当前点位 - 点位差)`。
- 输出到当前 TREE 的环比/边际变化列，并设置为 `0.00%` 百分比格式；不要依赖旧版固定 M 列。

运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\update_tree_stock_index_pct_change.ps1
```

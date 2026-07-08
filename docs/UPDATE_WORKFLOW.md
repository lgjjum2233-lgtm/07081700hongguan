# 日常更新流程

## 1. 刷新日报

1. 打开最新日报文件。
2. 在 WPS 顶部点击 `Wind`。
3. 点击 `刷新数据`。
4. 等待刷新完成。
5. 检查 `当前数据`、`数据日期`、`变化` 和趋势图是否更新。
6. 另存为新日期和版本号文件。

命名示例：

```text
20260527日报数据版V3_V1.xlsx
20260527日报数据版V3_V1.1.xlsx
```

## 2. 更新宏观分析汇总表

运行或参考脚本：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\match_daily_to_macro_summary.ps1
```

脚本会：

- 找到最新日报文件。
- 找到原始或最新汇总表。
- 匹配日报数据到汇总表。
- 更新 O/P/Q/R 列。
- 生成新的趋势图。
- 输出新版本文件。

注意：脚本当前以 2026-05-26 的文件结构为基准，后续如果增加 `全球宏观` 页，需要同步更新映射逻辑。

## 3. 验证

建议至少做三类验证：

- 用 WPS 打开文件，确认没有修复提示。
- 检查关键单元格，例如 O6、P6、Q6。
- 检查图表对象数量是否与匹配指标数量一致。

可使用：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\verify_wps_open_summary.ps1
```

## 4. 更新日志

每次更新后，记录：

- 输入文件
- 输出文件
- 修改内容
- 匹配数量
- 未匹配数量
- 验证结果
- 遗留问题

对应文件：

- `logs/CHANGELOG.md`
- `logs/WORK_LOG.md`


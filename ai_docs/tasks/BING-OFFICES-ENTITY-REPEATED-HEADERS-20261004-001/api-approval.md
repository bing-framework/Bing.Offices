# 公共 API 成员审批与迁移

- taskId: BING-OFFICES-ENTITY-REPEATED-HEADERS-20261004-001
- approvedBy: user（明确要求实施 Entity 单据重复页眉与打印标题计划）
- approvedAt: 2026-10-04 (Asia/Shanghai)
- 范围：仅增加 Entity 打印标题、实际重复页眉及 Provider SPI；不删除公共成员或改变既有签名。

## 批准的新增成员

| 类型 | 新增成员 |
|---|---|
| `ExcelEntityLayoutBuilder<TEntity>` | `PrintTitleRows(string sheetName, string rows)` |
| `ExcelEntityLayout<TEntity>` | `PrintTitleRows`（Provider SPI） |
| `ExcelEntityListRegionBuilder<TItem>` | `PageHeader(string sourceRange, string markerAddress, string markerText, Action<ExcelEntityPageHeaderBuilder<TItem>> configure = null)` |
| `ExcelEntityListRegion<TEntity>` | `PageHeader`（Provider SPI） |
| `ExcelEntityPageHeaderBuilder<TItem>` | 类型、公开无参构造、`Cell(string, object, ExcelCellStyle, string)`、`Cell<TValue>(string, Func<ExcelEntityPageHeaderContext<TItem>, TValue>, ExcelCellStyle, string)`、`Merge(string)` |
| `ExcelEntityPageHeaderContext<TItem>` | 类型、`PageNumber`、`PageCount`、`Items` |
| `ExcelEntityPageHeaderCell` | 类型、`Reference`、`Style`、`NumberFormat`、`GetValue(int, int, IReadOnlyList<object>)`（Provider SPI） |
| `ExcelEntityPageHeaderDefinition` | 类型、`SourceRange`、`Marker`、`MarkerText`、`Cells`、`Merges`、`RowCount`（Provider SPI） |
| `BingOfficesConfigurationException` | 新增带 `BingOfficesStage`、Sheet、原始一基行列坐标的构造重载；保留旧构造，用于页眉 marker 结构化错误 |

## 兼容性与迁移

现有布局、构造和 Fluent 调用不变。只需打印标题时使用 `PrintTitleRows("明细", "1:3")`；需要实际复制单据头时，在分页列表区域使用 `PageBreak(rowsPerPage)` 与 `PageHeader("A1:D3", "A1", "HEADER", ...)`。同一 Sheet 的两种模式互斥。实际页眉的 marker 用于导入跳过，不属于业务明细；模板源区域须紧邻首个明细行，合并必须完整包含于源区域。公式只复制文本，绘图对象不复制。

本阶段原有双 TFM 成员差异为 `Bing.Offices.Abstractions` 各新增 25 个成员、删除 0 个。本轮 Review 修复补充一个结构化配置异常构造，按相同基线合计新增 26 个、删除 0 个；其他生产程序集无成员差异。最终候选身份和快照比较见执行记录。

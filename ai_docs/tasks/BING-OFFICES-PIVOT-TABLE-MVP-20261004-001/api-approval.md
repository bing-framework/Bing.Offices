# 公共 API 成员审批与迁移

- taskId: `BING-OFFICES-PIVOT-TABLE-MVP-20261004-001`
- approvedBy: user（继续按 TODO 顺序实现能力）
- approvedAt: 2026-10-04（Asia/Shanghai）
- 范围：Workbook 透视表定义、Fluent 入口、Provider SPI 和细粒度能力位；不删除或改变现有公开签名。

## 批准的新增成员

| 类型 | 成员 |
|---|---|
| `ExcelPivotTableDefinition` | 公开类型、无参构造、`Name`、`SourceSheet`、`SourceRange`、`TargetSheet`、`TargetRow`、`TargetColumn`、`RowField`、`ValueField`、`Validate()` |
| `ExcelWorkbookExportBuilder` | `PivotTable(ExcelPivotTableDefinition definition)` |
| `ExcelWorkbookExportRequest` | `PivotTables`（隐藏于 IntelliSense 的 Provider SPI） |
| `ExcelProviderFeatures` | `PivotTables` |

双 TFM 对旧基线各新增 14 个成员、删除 0 个；仅涉及 `Bing.Offices.Abstractions`。候选差异保存于 `artifacts/api-before/api-diff.json`。

## 兼容性与迁移

原有 Workbook 构建调用保持不变。需要生成透视表时，在已声明的来源 Sheet 后调用 `PivotTable(...)`；来源区域包含表头和全部数据行，目标 Sheet 自动追加。ClosedXML XLSX 支持当前模型；其他 Provider 返回结构化 Unsupported。模板中已有透视表仍不在保留契约内。完整示例见 `docs/excel/pivot-tables.md`。

## Review 修复的 API 边界

FIX-001～006 修正业务来源名单、工作表名称及展开边界、辅助 Sheet 名称预留、XML 注释和验收矩阵，不新增或删除公共成员。最终候选重新构建与打包后，应先证明双 TFM 成员快照与当前批准基线一致，再刷新源码、程序集与包身份并执行完整比较。此次身份刷新不扩大原成员审批范围。

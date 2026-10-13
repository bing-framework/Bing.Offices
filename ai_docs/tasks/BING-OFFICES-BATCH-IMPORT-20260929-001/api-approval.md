# 公共 API 成员审批与迁移

- taskId: BING-OFFICES-BATCH-IMPORT-20260929-001
- approvedBy: user（要求继续按 TODO 顺序实现分批导入扩展）
- approvedAt: 2026-09-29 (Asia/Shanghai)
- 范围：增加独立 Workbook 多 Sheet 分批导入接口和结果描述；既有单 Sheet 接口、构造函数和签名不删除、不改签名。

## 批准的新增成员

- `IExcelWorkbookBatchImporter` 的同步和异步 Workbook 多 Sheet 方法。
- `ExcelWorkbookImportBatch` 的 Sheet 名称、实体类型、全局批次号、首行、实体和错误属性。
- `ExcelBatchImportSheetSummary`、`ExcelWorkbookBatchImportSummary` 类型及其只读聚合属性。
- `ExcelImportBatch<TItem>.SheetName` 和带 Sheet 名称的新构造函数；保留现有四参数构造函数。
- `ExcelProviderFeatures.MultiSheetBatchImport` 能力位。
- `ExcelSheetImportBuilder<TItem>.BatchSize` 及其执行描述的隐藏 `BatchSize` 属性，用于为 Workbook 分批 Sheet 设置批次大小。

## 兼容性与迁移

旧调用继续使用 `IExcelBatchImporter.ImportBatches`。需要一次处理 Workbook 中多个 Sheet 时，改从 DI 获取 `IExcelWorkbookBatchImporter`，提交已有的 `ExcelWorkbookImportRequest<TWorkbook>`；回调通过 `ExcelWorkbookImportBatch.SheetName` 和 `ItemType` 路由。每个 Sheet 仍遵守既有固定列、配置校验和资源限制边界。

双 TFM 候选快照预期：`Bing.Offices.Abstractions` 新增 31 个成员、删除 0 个；`Bing.Offices.ExcelDataReader` 新增 2 个方法并因新增接口产生 1 条实现类型快照变化。其他生产程序集成员差异为零。

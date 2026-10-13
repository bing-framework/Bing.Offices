# 生产符号到测试方法映射

| 生产符号 | 测试项目 | 测试方法 | 关键行为 |
|---|---|---|---|
| `Bing.Offices.Imports.IExcelWorkbookBatchImporter` | `Bing.Offices.ExcelDataReader.Tests` | `ImportWorkbookBatches_ShouldDeliverSheetsInOrder` | 多 Sheet 同步批次顺序与来源标识 |
| `Bing.Offices.Imports.IExcelWorkbookBatchImporter` | `Bing.Offices.ExcelDataReader.Tests` | `ImportWorkbookBatchesAsync_ShouldAwaitEachBatchAndPreserveSource` | 异步回调等待、输入流所有权 |
| `Bing.Offices.Imports.ExcelWorkbookImportBatch` | `Bing.Offices.ExcelDataReader.Tests` | `ImportWorkbookBatches_ShouldSupportHeterogeneousSheetTypes` | 异构 Sheet 实体类型与对象批次 |
| `Bing.Offices.Imports.ExcelWorkbookBatchImportSummary` | `Bing.Offices.ExcelDataReader.Tests` | `ImportWorkbookBatches_ShouldShareResourceBudgetAcrossSheets` | Sheet 摘要与聚合统计、共享资源预算 |
| `Bing.Offices.Imports.ExcelImportBatch{TItem}.SheetName` | `Bing.Offices.ExcelDataReader.Tests` | `Import_ShouldReadRealXlsbFixtureAndDeliverBatches` | 单 Sheet 批次来源定位 |
| `Bing.Offices.ExcelDataReader.ExcelDataReaderExcelImporter` | `Bing.Offices.ExcelDataReader.Tests` | `AddBingOfficesExcelDataReader_ShouldRegisterReadAndBatchServices` | DI 注册新契约 |

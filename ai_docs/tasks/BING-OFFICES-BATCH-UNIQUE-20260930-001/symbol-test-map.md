# 生产符号到测试方法映射

| 生产符号 | 测试项目 | 测试方法 | 关键行为 |
|---|---|---|---|
| `Bing.Offices.ExcelDataReader.ExcelDataReaderExcelImporter.ReadBatches` | `Bing.Offices.ExcelDataReader.Tests` | `ImportBatches_UniqueValues_ShouldRejectDuplicatesAcrossBatches` | 单 Sheet 跨批次唯一值判重、批次实体与错误交付 |
| `Bing.Offices.ExcelDataReader.ExcelDataReaderExcelImporter.MaterializeTypedRow` | `Bing.Offices.ExcelDataReader.Tests` | `ImportBatches_UniqueValues_ShouldRollbackFailedRowReservation` | 当前行失败回滚唯一值预留 |
| `Bing.Offices.Providers.UniqueTracker` | `Bing.Offices.ExcelDataReader.Tests` | `ImportBatches_UniqueValues_ShouldReportResourceLimit` | 唯一值数量上限与结构化 ResourceLimit |
| `Bing.Offices.ExcelDataReader.ExcelDataReaderExcelImporter.ReadWorkbookSheetBatchesAsync` | `Bing.Offices.ExcelDataReader.Tests` | `ImportWorkbookBatchesAsync_UniqueValues_ShouldIsolateSheetsAndRejectDuplicates` | 异步 Workbook 批次跨行判重、按 Sheet 隔离和物理坐标 |

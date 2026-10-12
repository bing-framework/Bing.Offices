# 生产符号到测试映射

| 生产符号 | 关键行为 | 测试项目 | 测试方法 |
| --- | --- | --- | --- |
| `ExcelDataReaderExcelImporter.ImportBatches` | XLSX 原生显式列表、错误坐标、合法行保留 | `Bing.Offices.ExcelDataReader.Tests` | `ExcelDataReaderBoundaryTest.ImportBatches_WorkbookValidation_ShouldRejectInvalidListValue` |
| `ExcelDataReaderExcelImporter.ImportWorkbookBatchesAsync` | 多 Sheet 名称列表、数值规则、异步批次和错误隔离 | `Bing.Offices.ExcelDataReader.Tests` | `ExcelDataReaderBoundaryTest.ImportWorkbookBatchesAsync_WorkbookValidation_ShouldUseNamedListAndNumbers` |
| `ExcelDataReaderExcelImporter.CreateWorkbookValidation` | XLSX 之外的原生校验预检边界 | `Bing.Offices.ExcelDataReader.Tests` | `ExcelDataReaderBoundaryTest.ImportBatches_WorkbookValidationOnXls_ShouldBeRejectedAtPreflight` |
| `ExcelDataReaderExcelImporter` | 唯一值跨批次、回滚和资源限制回归 | `Bing.Offices.ExcelDataReader.Tests` | `ImportBatches_UniqueValues_ShouldRejectDuplicatesAcrossBatches`、`ImportBatches_UniqueValues_ShouldRollbackFailedRowReservation`、`ImportBatches_UniqueValues_ShouldReportResourceLimit` |

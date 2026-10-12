# 生产符号到测试方法追溯

| 生产符号 | 测试项目 | 测试方法 |
| --- | --- | --- |
| `ExcelDataReaderExcelImporter.ImportBatches` | `Bing.Offices.ExcelDataReader.Tests` | `ImportBatches_DynamicColumns_ShouldIsolateValuesAndRollbackUnique` |
| `ExcelDataReaderExcelImporter.ImportBatchesAsync` | `Bing.Offices.ExcelDataReader.Tests` | `ImportBatchesAsync_DynamicColumns_ShouldConvertAndWriteTarget` |
| `ExcelDataReaderExcelImporter.ImportWorkbookBatches` | `Bing.Offices.ExcelDataReader.Tests` | `ImportWorkbookBatches_DynamicColumns_ShouldRejectUnknownHeader`; `ImportWorkbookBatches_DynamicColumns_ShouldApplyWorkbookValidation` |
| `ExcelDataReaderExcelImporter.ImportWorkbookBatchesAsync` | `Bing.Offices.ExcelDataReader.Tests` | `ImportWorkbookBatchesAsync_DynamicColumns_ShouldConvertAndWriteTarget` |

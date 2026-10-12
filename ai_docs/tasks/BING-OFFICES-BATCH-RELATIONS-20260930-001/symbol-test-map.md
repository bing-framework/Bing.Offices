# 生产符号到测试方法追溯

| 生产符号 | 关键行为 | 测试项目 | 测试方法 |
| --- | --- | --- | --- |
| `ExcelDataReaderExcelImporter.ImportWorkbookBatches` | XLS/XLSX 乱序 Sheet 关系绑定 | `Bing.Offices.ExcelDataReader.Tests` | `ImportWorkbookBatches_Relations_ShouldBindAfterOutOfOrderSheets` |
| `ExcelDataReaderExcelImporter.ImportWorkbookBatchesAsync` | 父行失败、子批次交付和关系错误诊断 | `Bing.Offices.ExcelDataReader.Tests` | `ImportWorkbookBatchesAsync_Relations_ShouldReportParentFailureAndMissingParent` |
| `ExcelDataReaderBatchRelationCoordinator<TWorkbook>` | 延迟绑定、父项索引、缺失父项错误 | `Bing.Offices.ExcelDataReader.Tests` | 同上两个关系职责级测试 |

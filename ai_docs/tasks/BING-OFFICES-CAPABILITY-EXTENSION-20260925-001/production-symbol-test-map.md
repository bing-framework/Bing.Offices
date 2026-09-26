# Production Symbol -> Test Map

| 生产符号/行为 | 直接测试方法 | 测试项目 |
|---|---|---|
| `ExcelFormat.Xlsb` | `Import_ShouldReadRealXlsbFixtureAndDeliverBatches`、`Capabilities_ShouldDescribeReadOnlyXlsbAndBatchSupport` | `Bing.Offices.ExcelDataReader.Tests` |
| `IExcelProviderCapabilityDescriptor` | `ExcelDataReader_ShouldRunCompleteAndBatchImportContracts` | `Bing.Offices.ProviderContract.Tests` |
| `IExcelBatchImporter.ImportBatches` | `ImportBatches_ShouldDeliverBatchesAndKeepSourceOwnedByCaller` | `Bing.Offices.ExcelDataReader.Tests` |
| `IExcelBatchImporter.ImportBatchesAsync` | `ImportBatchesAsync_ShouldAwaitEachBatchAndHonorCancellation` | `Bing.Offices.ExcelDataReader.Tests` |
| 分批缺失表头、正文起始行、动态列/唯一性边界 | `ExcelDataReaderBoundaryTest` 中对应同步/异步职责测试 | `Bing.Offices.ExcelDataReader.Tests` |
| 分批结构化异常、回调取消与暂存清理 | `ImportBatches_MalformedWorkbook_ShouldUseStructuredProviderException`、`ImportBatchesAsync_CallbackCancellation_ShouldPropagateAndCleanStaging` | `Bing.Offices.ExcelDataReader.Tests` |
| 旧 Provider 的 XLSB 拒绝 | `NpoiFormatBoundaryTest`、`Xlsb_ShouldBeRejectedAsUnsupported`、`UnsupportedXlsb_ShouldFailBeforeWriting` | `Bing.Offices.Npoi.Tests`、`Bing.Offices.MiniExcel.Tests`、`Bing.Offices.ClosedXml.Tests` |
| 完整导入共享 Workbook `MaxRows` | `Import_ShouldShareRowBudgetAcrossSheetsAndHideEarlierEntities` | `Bing.Offices.ExcelDataReader.Tests` |
| ExcelDataReader Workbook 级 `MaxSheets`/`MaxCells` 与表头前物理 Cell 预检 | `Import_ShouldApplySheetColumnAndCellBudgetsWithoutPartialEntities`、`Import_ShouldCountWorkbookCellsOnUnselectedSheetsAndBeforeHeader`、`ImportAsync_ShouldCountWorkbookCellsOnUnselectedSheets` | `Bing.Offices.ExcelDataReader.Tests` |
| XLSX 精确物理 Cell 边界与 XLSB 限制拒绝 | `Import_ShouldCountWorkbookCellsOnUnselectedSheetsAndBeforeHeader`、`Import_XlsbPhysicalCellLimits_ShouldBeExplicitlyUnsupported` | `Bing.Offices.ExcelDataReader.Tests` |
| XLSX 映射/转换 | `Import_ShouldReadNpoiGeneratedXlsxAndMapValues` | `Bing.Offices.ExcelDataReader.Tests` |
| XLS 映射 | `Import_ShouldReadNpoiGeneratedXls` | `Bing.Offices.ExcelDataReader.Tests` |
| 中文、日期、空值 | `Import_ShouldPreserveUnicodeDateAndEmptyValue` | `Bing.Offices.ExcelDataReader.Tests` |
| 1900/1904 日期系统 | `Import_ShouldRespectWorkbookDateSystem` | `Bing.Offices.ExcelDataReader.Tests` |
| 公式缓存值 | `Import_ShouldReadCachedFormulaValue` | `Bing.Offices.ExcelDataReader.Tests` |
| 损坏文件结构化异常 | `Import_MalformedWorkbook_ShouldReturnStructuredProviderException` | `Bing.Offices.ExcelDataReader.Tests` |
| Failure Workbook 预检拒绝 | `Import_ShouldRejectFailureWorkbookBeforeCreatingOutput` | `Bing.Offices.ExcelDataReader.Tests` |
| 只读 Provider 无导出器 | `ImportOnlyProviderContractTest.ExcelDataReader_ShouldRunCompleteAndBatchImportContracts` | `Bing.Offices.ProviderContract.Tests` |
| DI 仅注册读取服务 | `AddBingOfficesExcelDataReader_ShouldRegisterReadAndBatchServices` | `Bing.Offices.ExcelDataReader.Tests` |
| Public API 增量和双 TFM identity | `PublicApiContractTest`、`ApiSnapshot` compare、package consumer | `Bing.Offices.Tests.Integration`、`build/ApiSnapshot`、Consumer.Net6/Net8 |

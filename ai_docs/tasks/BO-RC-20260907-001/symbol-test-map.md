# 生产符号到测试方法映射

| 生产符号/职责 | 测试项目 | 测试方法 |
| --- | --- | --- |
| `IExcelImporter.ImportAsync` / `NpoiExcelImporter.ImportAsync` | `Bing.Offices.Tests` | `AsyncPipelineTest.ExcelAsync_ShouldUseAsyncOuterStream_AndMatchSyncImport` |
| Excel Failure Workbook Async staging | `Bing.Offices.Tests` | `AsyncPipelineTest.ExcelAsync_FailureWorkbook_ShouldCopyToAsyncDestination` |
| `IExcelExporter.ExportAsync` | `Bing.Offices.Tests` | `AsyncPipelineTest.ExcelAsync_ShouldUseAsyncOuterStream_AndMatchSyncImport` |
| `IExcelExporter.ExportToFileAsync` | `Bing.Offices.Tests` | `AsyncPipelineTest.AsyncFileExport_PreCanceled_ShouldNotCreateTarget` |
| `ICsvImporter.ImportAsync` | `Bing.Offices.Tests` | `AsyncPipelineTest.CsvAsync_ShouldUseAsyncReadWrite_AndMatchSyncResult`; `AsyncPipelineTest.CsvAsync_PreCanceled_ShouldPreserveCancellation`; `AsyncPipelineTest.CsvAsync_MidReadAndMidWriteCancellation_ShouldPropagate` |
| `ICsvExporter.ExportAsync` | `Bing.Offices.Tests` | `AsyncPipelineTest.CsvAsync_ShouldUseAsyncReadWrite_AndMatchSyncResult`; `AsyncPipelineTest.CsvAsync_MidReadAndMidWriteCancellation_ShouldPropagate` |
| `ICsvExporter.ExportToFileAsync` | `Bing.Offices.Tests` | `AsyncPipelineTest.AsyncFileExport_PreCanceled_ShouldNotCreateTarget` |
| Async error parity / exporter Observer | `Bing.Offices.Tests` | `AsyncPipelineTest.CsvAsync_ErrorResult_ShouldMatchSyncResult`; `AsyncPipelineTest.ExcelAsync_FileCommitFailure_ShouldObserveSameExceptionOnce` |
| `IFileExportCommitter.CommitAsync` / `DefaultFileExportCommitter` | `Bing.Offices.Tests` | `DefaultFileExportCommitterTest.CommitAsync_NewTarget_ShouldWriteAndMove`; `CommitAsync_ExistingTarget_ShouldReplaceAfterSuccessfulWrite`; `CommitAsync_WriteFailure_ShouldPreserveExceptionAndTarget`; `CommitAsync_PreCanceled_ShouldNotCreateFiles`; `CommitAsync_CanceledAfterWrite_ShouldCleanTemporaryFile` |
| `NpoiStreamCopier.CopyAsync` | `Bing.Offices.Tests` | AsyncOnly read/write assertions in `AsyncPipelineTest`; existing sync size/cancellation cases in `StreamPipelineTest` |
| Core/NPOI 85 个公开扩展完整签名 | `Bing.Offices.Tests` | `PublicExtensionCoverageTest.PublicExtensions_ShouldHaveDirectBehaviorTestForEverySignature`；逐签名到职责测试的方法映射见 `artifacts/reports/public-extension-coverage.md` |
| API classification/IVT | `Bing.Offices.Tests` | `PublicApiContractTest.PublicApi_ExportedTypes_ShouldHaveGovernedClassification`; `PublicApi_ProductionAssemblies_ShouldNotExposeProductionFriendAssemblies` |
| Sync Excel/CSV core parity | Unit/Integration | Existing `StreamPipelineTest`, `CsvTest`, `ExcelP0RegressionTest`, `ExcelImporterIntegrationTest` on net6/net8 |

API baseline 已由维护者 `jian玄冰` 批准；net6/net8 API snapshot compare 与 `PublicApiContractTest` 均通过。Phase 9 的完整 Unit TRX 为 `artifacts/test-results/BO-RC-phase9-net6-unit.trx`、`BO-RC-phase9-net8-unit.trx`，定向扩展 TRX 为对应的 `*-extensions.trx`；它们替代早期 `BO-RC-continued-*` 证据。

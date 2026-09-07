# Unit Test 报告

## 结果

命令：`dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 -c Release --no-restore`

- `492 passed / 0 skipped / 1 failed / 493 total`
- 唯一失败：`PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`，明确报 baseline `approvedBy/approvedAt` 缺失；计划禁止用 Skip 掩盖审批门禁。
- Build：net8.0 Release 成功；本轮 solution build 为 `0 warning / 0 error`。

## 直接覆盖映射

| 生产职责 | 直接测试 |
| --- | --- |
| exporter 文件提交、Observer 同实例/单次 | `ReviewFixRegressionTest.ExporterFileCommitFailure_ShouldObserveSameExceptionOnce`、`ExcelWorkbookRequestTest.ExportToFile_*`、`CsvTest` 文件导出测试 |
| 默认文件提交器/SPI | `DefaultFileExportCommitterTest.Commit_NewTarget_ShouldWriteAndMove`、`Commit_ExistingTarget_ShouldReplaceAfterSuccessfulWrite`、`Commit_WriteFailure_ShouldPreserveExceptionAndTarget`、`Commit_PreCanceled_ShouldNotCreateFiles`；Exporter 同实例与 DI 注册由 `ReviewFixRegressionTest.ExporterFileCommitFailure_ShouldObserveSameExceptionOnce`、`MappingPlanFactory_DiDefaultAndReplacement_ShouldPreserveOwnershipBoundary` 覆盖 |
| TryAddPicture 前置和后置失败 | `StreamPipelineTest.SheetExtensions_TryAddPictureInvalidArguments_ShouldThrowWithoutMutation`、`SheetExtensions_TryAddPicturePostMutationFailure_ShouldThrowExplicitExportException`、`SheetExtensions_TryAddPictureResizeFailure_ShouldThrowExplicitExportException`、`SheetExtensions_TryAddPictureValidImage_ShouldSupportBothProviders` |
| JSON/XML loader 删除 diagnostics | `ExcelMappingConfigurationLoaderTest.RemovedDiagnosticsSurface_ShouldNotBePublished` 及现有 JSON/XML round-trip 测试 |
| API canonicalizer | `PublicApiContractTest.PublicApiSnapshot_CanonicalLines_ShouldIncludeGovernedMetadata` |
| 资源/取消/文件清理 | `ExcelP0RegressionTest.Import_FailureWorkbook_CandidateRowBudget_ShouldRejectBeforeMutation`、现有 AtomicFileCommitter 系列和取消测试 |
| Failure Workbook preflight | `NpoiFailureWorkbookPreflightTest` 覆盖错误行、单元格、图片数量/字节、目标对象估算的边界与超限，以及新预算非正值配置 |
| Failure Workbook annotation/summary | `ExcelP0RegressionTest` Failure Workbook 138-test slice, including annotation, cleanup and serialization contracts |
| CSV/Mapping职责拆分 | CSV `CsvTest` 39 tests；`ReviewFixRegressionTest`/Mapping slice 83 tests；结构移动后 solution rebuild 0 warning |

SQL 元数据规则不适用：本任务未修改 `Bing.Data.Sql`。

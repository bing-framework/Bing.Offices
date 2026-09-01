# Unit Test Report

状态：`VERIFIED`

| 范围 | 结果 |
| --- | --- |
| selector | VERIFIED；定向通过，单一解析结果贯穿 plan/import/failure mapping。 |
| 原子文件提交 | VERIFIED；当前 net6/net8 相关 Unit 通过，包含 Move/Replace 与 cleanup 诊断。 |
| Failure Workbook | VERIFIED；当前 net6/net8 相关 Unit 通过，包含目录、copy、MaxBytes、取消和清理诊断。 |
| metadata | VERIFIED；新增 XLS 默认 preserve 与 64 路 XLS/XLSX 请求隔离测试。 |
| CSV/Excel 资源与释放 | VERIFIED；本轮增加 Excel non-seekable 超限、调用方流所有权和反射异常解包主链测试；CSV/Excel 限制、取消、leave-open 与 Failure Workbook 清理矩阵持续通过。 |
| 全量 net6/net8 | VERIFIED；各 `360/360`，0 失败，0 跳过。 |

测试方法遵循英文命名、中文 XML 目的和 AAA 结构；不通过弱化断言或删除测试解决失败。

## 生产符号追溯

- `NpoiResolvedSheet` / `NpoiExcelImporter.ResolveSheet` -> `Import_SheetSelectorByIndex_ShouldReadSelectedSheet`、`Import_DuplicateSheetSelectors_ShouldFailBeforePlanExecution`。
- `NpoiImportPlanBuilder.Create` -> `Import_WorkbookRequest_ShouldUsePerSheetHeaderAndDynamicDefinitions` 及全量 Workbook import tests。
- `ExcelWorkbookMetadataOptions` / `Metadata(...)` -> `Export_WorkbookRequest_ShouldSnapshotMetadata`、`Export_TemplateMetadata_ShouldPreserveByDefaultAndOverrideExplicitly`、`Export_XlsTemplateMetadata_WhenNotSpecified_ShouldPreserveTemplateValues`、`Export_MetadataConcurrentRequests_ShouldRemainIsolated`。
- `AtomicFileCommitter` -> `ReviewFixRegressionTest` atomic failure matrix。
- `NpoiFailureWorkbookWriter` -> `ExcelP0RegressionTest` and integration Failure Workbook matrix。

## RH31-105 符号级追溯矩阵

| 生产符号/责任 | 直接测试方法 | 覆盖契约 |
| --- | --- | --- |
| `NpoiStreamCopier.Copy` | `NpoiStreamCopier_Copy_ShouldPreserveDataAndStreamOwnership`; `NpoiStreamCopier_Copy_WhenMaxBytesExceeded_ShouldThrowBeforeWrite`; `NpoiStreamCopier_Copy_WhenCancelled_ShouldThrowBeforeIo` | 数据保真、输入上限、取消前置检查、源/目标所有权 |
| `NpoiExcelImporter.Import` | `Import_NonSeekableStream_ShouldKeepSourceOpenAndReturnItems`; `Import_NonSeekableInputOverLimit_ShouldRejectAndKeepSourceOpen`; `Import_XlsAndInvalidWorkbook_ShouldSupportFormatAndKeepSourceOpen`; `StreamPipeline_PreCancelledToken_ShouldThrowOperationCanceledException`; `StreamPipeline_CancelDuringProcessing_ShouldPreserveCallerStreams` | seekable/non-seekable、MaxInputBytes、格式失败、取消和源流保留 |
| `NpoiImportPlanBuilder.Create` / `CreateWorkbookPlan` | `Import_WorkbookRequest_ShouldUsePerSheetHeaderAndDynamicDefinitions`; `Import_MappingFactoryFailure_ShouldPreserveOriginalExceptionType` | 多 Sheet 计划复用、反射异常保持原始类型 |
| `NpoiImportRowMaterializer` | `Import_InvalidAndDuplicateValues_ShouldReturnStructuredErrors`; `Import_ConversionFailure_ShouldNotPolluteDuplicateValidationState`; `Import_ThrowingValidationRule_ShouldReturnStructuredValidationError`; `Import_UnboundCustomValidationAttribute_ShouldThrowConfigurationException` | 转换、校验、异常分类和失败行回滚 |
| `CsvEntityImporter.Import` | `EntityPipeline_ResourceLimits_ShouldClassifyAndTruncate`; `EntityPipeline_MaxInputBytes_ShouldRejectSeekableAndNonSeekableOverflow`; `EntityPipeline_MaxErrors_ShouldTruncateWithoutExceedingLimit`; `EntityPipeline_PreCancelledOperation_ShouldThrow` | 输入/行/列/字段/错误上限和取消 |
| `CsvEntityExporter.Export` | `EntityPipeline_EscapedValues_ShouldRoundTripAndKeepStreamsOpen`; `EntityPipeline_BomNullValuesAndNonSeekableStream_ShouldRoundTripAndKeepStreamOpen`; `EntityPipeline_PreCancelledOperation_ShouldThrow` | 转义、编码、不可寻址目标、leave-open 和取消 |
| `CsvRecordReader.Read` / `CsvLimitedReadStream.Dispose` | `EntityPipeline_BomNullValuesAndNonSeekableStream_ShouldRoundTripAndKeepStreamOpen`; `EntityPipeline_MaxInputBytes_ShouldRejectSeekableAndNonSeekableOverflow` | parser 读取、非寻址输入、限制包装和调用方流所有权 |
| `CsvRecordWriter.Write` | `EntityPipeline_EscapedValues_ShouldRoundTripAndKeepStreamsOpen`; `EntityPipeline_BomNullValuesAndNonSeekableStream_ShouldRoundTripAndKeepStreamOpen` | writer flush、转义和目标流保留 |
| `NpoiFailureWorkbookWriter.Write` | `Import_AnnotatedFailureWorkbook_ShouldAnnotateOriginal`; `FailureWorkbook_CancellationDuringCopy_ShouldCleanTemporaryFile`; `FailureWorkbook_PrimaryFailureWithCleanupFailure_ShouldPreserveCleanupException` | 失败工作簿、取消清理、主异常与清理异常并存 |
| `AtomicFileCommitter.Commit` | `AtomicFileCommitter_PrimaryFailureWithCleanupFailure_ShouldPreserveBothExceptions`; `AtomicFileCommitter_CommitFailure_ShouldKeepExistingTargetAndCleanup` 及 Move/Replace 双故障矩阵 | 原子提交、目标保留、staging 清理和双异常诊断 |
| `ExcelMappingPlanFactory.Create` / `CreateWorkbook` | `MappingPlan_WorkbookAndTenantVersion_ShouldBeImmutableAndIsolated`; `MappingPlan_DynamicConfigurationMutation_ShouldNotChangeExistingPlan`; `MappingPlan_ConcurrentReads_ShouldReuseSamePlan`; `MappingPlan_CacheEviction_ShouldRebuildProductionPlan` | 快照隔离、并发命中、淘汰和重建 |
| `PublicApiContractTest` API 门禁 | `PublicApi_ReleaseAssemblies_ShouldMatchApprovedBaseline`; `PublicApi_ExportedTypes_ShouldHaveGovernedClassification`; `PublicApi_PublicMembers_ShouldNotExposeNpoiTypes`; `PublicApi_ProductionAssemblies_ShouldNotExposeProductionFriendAssemblies`; `PublicApi_NpoiAssembly_ShouldMatchExactMemberBaseline` | 导出面、分类、NPOI 泄漏、IVT 和 NPOI 注册入口 |
| 独立资源探针 | `Import_ResourceProbe_ShouldRunInIndependentProcess` | 独立进程、LOH、峰值工作集和 16 场景资源上限 |

上述矩阵在 net6.0 和 net8.0 Unit 目标上执行；新增本轮直接测试在 net8.0 定向执行通过，完整多 TFM 结果以本任务最终验证命令为准。

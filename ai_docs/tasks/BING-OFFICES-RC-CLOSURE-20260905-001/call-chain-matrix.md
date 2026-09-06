# 调用链矩阵

| 公共入口 | Mapping/Plan | Core/Provider | IO/序列化 | 异常边界 | 测试状态 |
| --- | --- | --- | --- | --- | --- |
| `IExcelExporter.Export` | `NpoiExportPlanBuilder` 缓存泛型委托 | `NpoiExcelExporter` / `NpoiExportSheetWriter` | NPOI Workbook.Write | exporter dispatcher | 多 Sheet、异常、Benchmark 已验证 |
| `ExcelStreamExtensions.ExportToFile` | 同上 | `AtomicFileCommitter` | temp + Flush + Replace/Move | 内容失败原样传播；文件系统失败为 FileCommit | Unit/Integration 已验证 |
| `IExcelImporter.Import` | mapping plan factory + 缓存 Sheet/关系委托 | `NpoiExcelImporter` / row materializer | MaxInputBytes、XLSX preflight 后 WorkbookFactory | importer dispatcher | Unit/Integration/ResourceProbe 已验证 |
| `ICsvExporter.Export` | mapping plan factory | `CsvEntityExporter` | StreamWriter/CsvRecordWriter | exporter dispatcher | 基线覆盖；file path 同 Atomic 风险 |
| `ICsvImporter.Import` | mapping plan factory | `CsvEntityImporter` | CsvHelper reader | importer dispatcher/结构化行错误 | 基线覆盖 |
| `IExcelMappingConfigurationLoader` | v2 parse/merge | Core loader provider | 安全 JSON/XML reader | DI 入口观察同实例；静态入口纯解析 | 独立合同 4/4 |
| `SheetExtensions.TryAddPicture` | 参数/范围 | HSSF/XSSF drawing | picture bytes | 只吞明确可恢复状态 | HSSF/XSSF 与参数边界已验证 |
| `SheetExtensions.MovePictures` | 范围 | HSSF/XSSF | drawing anchors | null 参数；未知实现 UnsupportedFeature | Provider 成功/失败分支已验证 |
| Failure Workbook | error selection | `NpoiFailureWorkbookWriter` | 源/目标 DOM + serialize | 可选行元数据降级发结构化诊断 | 正确性已验证；完整双 DOM 资源矩阵待预算 |

## 追踪要求

每次修改后追加“最终生产符号 -> 测试项目/方法”的职责级映射；公共 API、Provider 分支、缓存、格式化与 Builder 路径不得只依赖综合测试。

## 已建立的生产符号到测试方法追踪

| 最终生产符号/行为 | 测试项目 | 方法 |
| --- | --- | --- |
| `AtomicFileCommitter.Commit` 内容失败保留实例 | Bing.Offices.Tests | `AtomicFileCommitter_ContentFailure_ShouldPreserveExceptionInstance` |
| `AtomicFileCommitter.Commit` Move/Replace/File cleanup | Bing.Offices.Tests | `AtomicFileCommitter_CommitFailure_ShouldKeepExistingTargetAndCleanup`、`AtomicFileCommitter_MoveFailureWithCleanupFailure_ShouldPreserveDiagnostics`、`AtomicFileCommitter_ReplaceFailureWithCleanupFailure_ShouldKeepExistingTarget` |
| Excel/CSV `ExportToFile` 内容失败不误归 FileCommit | Bing.Offices.Tests | `ExportToFile_Failure_ShouldKeepExistingTarget`、`StreamExtensions_FileExportFailure_ShouldKeepExistingTarget` |
| `SheetExtensions.TryAddPicture` 参数边界 | Bing.Offices.Tests | `SheetExtensions_TryAddPictureInvalidArguments_ShouldThrowWithoutMutation` |
| `SheetExtensions.TryAddPicture` XSSF/HSSF 成功 | Bing.Offices.Tests | `SheetExtensions_TryAddPictureValidImage_ShouldSupportBothProviders` |
| `CellExtensions.SetValue(DateTimeOffset)` XLSX/XLS | Bing.Offices.Tests | `Export_DateTimeOffset_ShouldWriteStableOffsetText` |
| `CsvEntityExporter` DateTimeOffset 往返 | Bing.Offices.Tests | `EntityPipeline_DateTimeOffset_ShouldRoundTripWithStableOffsetText` |
| `ExcelDateParser` explicit/fixed offset | Bing.Offices.Tests | `TryParse_DateTimeOffset_ShouldRequireExplicitOrConfiguredOffset` |
| `DefaultExcelMappingConfigurationLoader` Observer 同实例 | Bing.Offices.Tests | `DefaultLoader_InvalidDocument_ShouldObserveSameInstanceOnce`、`DependencyInjection_DefaultLoader_ShouldUseRegisteredObserver` |
| 静态 `ExcelMappingConfigurationLoader` 纯解析 | Bing.Offices.Tests | `StaticLoader_InvalidDocument_ShouldThrowWithoutObservation` |
| `NpoiFailureWorkbookWriter.CopyOptionalRowMetadata` 降级诊断 | Bing.Offices.Tests | `FailureWorkbook_UnsupportedRowMetadata_ShouldReportStructuredDiagnostic` |
| `ExcelResourceLimits.MaxInputBytes` XLS/XLSX DOM 前限制 | Bing.Offices.Tests | `Import_InputResourceLimit_ShouldCoverBothWorkbookProviders`、`ResourceLimits_DefaultInputBudget_ShouldBeBoundedAndOptional` |
| `ExportMappingBuilder<T>.HasConverter/Map` 方向化导出 | Bing.Offices.Tests / Integration | `StreamPipeline_FluentMappingConfiguration_ShouldApplyToSingleRequest`、`AddBingOfficesNpoi_FluentConfigurationAndXmlConfiguration_ShouldRoundTripRealWorkbook` |
| v2-only JSON/XML Loader | Bing.Offices.Tests / Docs.Tests | `MappingConfigurationLoader_JsonAndXml_ShouldLoadAndRejectDtd`、`MappingDocuments_ExternalConsumer_ShouldLoadV2AndPreserveStreams` |
| `NpoiRelationBinder` 泛型委托缓存 | Bing.Offices.Tests | `RelationBinder_SameTypeCombination_ShouldCacheInvoker`、`Import_RelationDelegateFailure_ShouldPreserveOriginalExceptionType` |
| `SheetExtensions.MovePictures` 未知 Provider | Bing.Offices.Tests | `SheetExtensions_MovePicturesUnknownSheet_ShouldThrowUnsupportedFeature` |
| `BingOfficesException` 抽象基类 | Bing.Offices.Tests | `BingOfficesException_BaseType_ShouldBeAbstract` |
| 三个 nupkg 与七个 NPOI 扩展容器 | PackageConsumer | `artifacts/package-consumer/runtime/Program.cs`，三运行 TFM 输出 `package-consumer-ok` |

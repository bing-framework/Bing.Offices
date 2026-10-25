# 生产符号与职责测试映射

任务：BO-API-USABILITY-20261025。下表覆盖本任务新增 API 和行为变化，保留旧任务的映射。Unit 指 `tests/Bing.Offices.Tests`（net6.0、net8.0），Docs 指 `tests/Bing.Offices.Docs.Tests`（net8.0）。完整规范签名见 [api-changes.md](api-changes.md)。

| 最终生产符号 | 关键行为 | 测试项目与方法 |
| --- | --- | --- |
| BingOfficesServiceCollectionExtensions / AddBingOfficesCore | Core-only CSV 完整输出、导入和校验；同步及异步 | Unit：CoreServiceRegistrationTest.AddBingOfficesCore_ShouldSupportCsvWithoutExcelProvider |
| AddBingOfficesCore / ExcelNpoiServiceCollectionExtensions.AddBingOfficesNpoi | 重复注册幂等、保留自定义服务、两种注册顺序 | Unit：CoreServiceRegistrationTest.AddBingOfficesCore_ShouldBeIdempotentAndPreserveExistingServices |
| AddBingOfficesCore | null 服务集合异常 | Unit：CoreServiceRegistrationTest.AddBingOfficesCore_NullServices_ShouldThrow |
| ExcelSingleSheetImportExtensions.ImportSheet / ImportSheetAsync / ImportSheetFromFile / ImportSheetFromFileAsync / ImportSheetFromBytes / ImportSheetFromBytesAsync | 六方法均直接调用，XLS/XLSX 成功和完整结构化失败、源行、流生命周期 | Unit：ExcelSingleSheetImportTest.ImportSheet_AllEntrypoints_ShouldReturnCompleteFirstSheet；ImportSheet_AllEntrypoints_ShouldPreserveInvalidHeader |
| 上述六方法 | 取消原样传播、Provider IOException 类型及实例不变、异步委托 | Unit：ExcelSingleSheetImportTest.ImportSheet_AllEntrypoints_ShouldPreserveCancellationAndProviderException |
| ExcelSingleSheetImportExtensions.CreateRequest / ToResult | 名称选择、映射、MaxErrors、生效配置、缺 Sheet 的全局错误不丢失、失败后再次调用 | Unit：ExcelSingleSheetImportTest.ImportSheet_OptionsAndMissingSheet_ShouldPreserveWorkbookDiagnostics |
| ExcelSingleSheetImportResult<T> 类型及构造器 | 保留集合引用与旧截断信息，适配器创建真实结果 | Unit：ImportResultStatusTest.ResultConstructors_ShouldSummarizeMultipleReasonsAndUnspecifiedTruncation；ExcelSingleSheetImportTest.ImportSheet_AllEntrypoints_ShouldReturnCompleteFirstSheet |
| ExcelSingleSheetImportResult<T>.Items / Errors / SourceRows / ErrorsTruncated / MaxErrors | 成功实体、完整错误、对应源行、截断和上限 | Unit：ExcelSingleSheetImportTest.ImportSheet_AllEntrypoints_ShouldPreserveInvalidHeader；ImportSheet_OptionsAndMissingSheet_ShouldPreserveWorkbookDiagnostics |
| ExcelSingleSheetImportResult<T>.IsSuccess / IsComplete / StopReason | 普通行错误、不完整结果和成功状态 | Unit：ExcelSingleSheetImportTest.ImportSheet_AllEntrypoints_ShouldReturnCompleteFirstSheet；ImportSheet_AllEntrypoints_ShouldPreserveInvalidHeader；ImportResultStatusTest.ResultConstructors_ShouldSummarizeMultipleReasonsAndUnspecifiedTruncation |
| ExcelWorkbookImportResult<T>.IsComplete / StopReason；ImportResultState.GetReason(Excel) | 正常、普通行失败、表头、资源、错误上限、显式范围、旧标志保留；XLS/XLSX × Sync/Async | Unit：ImportResultStatusTest.Excel_Import_ShouldReportCompletionWithoutChangingLegacyFlags |
| CsvImportResult<T>.IsSuccess / IsComplete / StopReason；ImportResultState.GetReason(CSV) | CSV 正常、行错误、表头、行/字节限制、错误上限、部分成功数据；Sync/Async | Unit：ImportResultStatusTest.Csv_Import_ShouldReportCompletionAndPreservePartialItems |
| ImportStopReason.None / InvalidHeader / ResourceLimit / ErrorLimit / Unspecified | 原因优先级、无原因旧截断、兼容手工构造结果 | Unit：ImportResultStatusTest.ResultConstructors_ShouldSummarizeMultipleReasonsAndUnspecifiedTruncation；上述 Excel/CSV 状态测试 |
| ImportStopReason 类型及编译器 value__ 字段；两种新增静态扩展类型 | 公开类型分类、完整成员和双 TFM 快照 | Unit：PublicApiContractTest.PublicApi_ExportedTypes_ShouldHaveGovernedClassification；PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot；ApiSnapshot CLI |
| NpoiExcelImporter 导入循环 / ExcelImportRuntime 行限额 | Sheet 边界耗尽时明确 ResourceLimit；最后一张 Sheet 恰好额度不误报 | Unit：ImportResultStatusTest.Excel_Import_ShouldReportCompletionWithoutChangingLegacyFlags（sheet-boundary、exact） |
| ExcelSheetImportBuilder 默认校验 / MaxReadColumns；ValidateMode | 默认每行首错、后续行继续；100 列成功、101 列拒绝、显式扩容成功 | Unit：UsabilityDefaultsTest.Import_DefaultValidation_ShouldContinueAfterInvalidRow；Import_DefaultColumnLimit_ShouldAllow100Reject101AndHonorExpansion |
| ExcelExport.UseTemplate（仅文档澄清） | 成功、失败、读取中取消的 leaveOpen 约定，预取消不接管模板 | Unit：TemplateAsyncBoundaryTest.ExportAsync_TemplateSuccess_ShouldHonorLeaveOpen；ExportAsync_TemplateReadFailure_ShouldHonorLeaveOpen；ExportAsync_TemplateReadCancellation_ShouldHonorLeaveOpen；ExportAsync_PreCancelled_ShouldNotReadOrDisposeTemplate（复用既有直接测试） |
| Regex 校验文档模型 | 从 Markdown 原文编译模型，ORD-123 成功、BAD 完整错误 | Docs：DocsConsumerTest.DocumentationFences_FromMarkdown_ShouldCompileAndExecuteIndividually |
| Core 注册、CSV 导出/导入、Excel Workbook 导出和单表导入 | 两份完整 ConsoleApplication 原文编译运行，没有隐藏 using、变量或 DTO | Docs：QuickStartConsumerTest.QuickStartPrograms_FromMarkdown_ShouldCompileAndExecuteIndividually |

扩展方法 IL 追溯门禁：Unit `PublicExtensionCoverageTest.PublicExtensions_ShouldHaveDirectBehaviorTestForEverySignature`；最终 Core 22、NPOI 66，共 88 个签名。逐签名到直接行为测试的完整报告保存在 `TestResults/usability/final/extension-coverage-net6.md`、`extension-coverage-net8.md`。未删除门禁，也未以方法指针或未执行 lambda 冒充直接调用。

最终测试结果见 [execution.md](execution.md)。新入口没有改变文件提交、缓存键、DOM 并发或外围异步 IO 实现；这些既有职责测试随双 TFM 全套执行。

## Review 修复 Round 1 补充

| 最终生产符号 | 关键行为 | Unit 直接测试（net6.0 / net8.0） |
| --- | --- | --- |
| NpoiExcelImporter.ImportAsync / ImportCore；单表 ImportSheet / ImportSheetAsync 委托路径 | 省略配置采用实际 128 MiB 预算、恰好额度可继续、超出一字节拒绝；显式 null 允许超出；真实读取和原始异常/流所有权 | DefaultInputLimitTest.Import_DefaultBudget_ShouldEnforceBoundaryAndHonorExplicitNull（12 组，串行） |
| 同上 | XLS/XLSX × Sync/Async × Workbook/单表；真实文件 IO；默认、显式精确限额、显式关闭均成功；少一字节限额返回完整资源异常 | DefaultInputLimitTest.Import_RealFile_ShouldSupportDefaultExactDisabledAndRejectOverflow（8 组） |
| NpoiExcelImporter.ImportBufferedCore 行限额诊断 | 索引、不同大小写名称使用真实表名；缺失名称回退；完整错误与失败摘要、部分成功数据、只报一次和流所有权 | SheetBoundaryDiagnosticTest.Import_RowLimitAtSheetBoundary_ShouldReportResolvedNameAndCompleteFailureSummary（12 组） |

本轮没有公共签名增删，API 成员追溯继续沿用上表；双 TFM canonical 契约比较验证不变。原始 review.md 保留为独立审查证据，修复结果单独写入 execution.md。

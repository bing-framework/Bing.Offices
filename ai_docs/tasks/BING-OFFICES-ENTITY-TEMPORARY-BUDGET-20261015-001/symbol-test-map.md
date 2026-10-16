# 生产符号与直接测试映射

| 生产符号／行为 | 测试项目 | 测试方法 |
| --- | --- | --- |
| `ExcelEntityBatchImportOptions` 五参数构造及默认兼容 | NPOI.Tests、ClosedXml.Tests，双 TFM | `Options_ShouldRejectNonPositiveBudgetAndPreserveDefaults` |
| `ExcelTemporaryStorageBudget.CreateFile` 多文件总额、精确边界与释放 | 同上，共享测试 | `CreateFile_WhenAggregateReachesLimit_ShouldRejectOnlyFurtherGrowth` |
| 预算流覆盖写、Seek、SetLength、归还额度 | 同上 | `WriteAndSetLength_ShouldAccountForGrowthOverwriteShrinkAndRelease` |
| 预算流真实异步写、预取消、异步释放 | 同上 | `WriteAsync_ShouldAccountCancelAndReleaseAcrossFiles` |
| Seek 超限后的空写不占额度 | 同上 | `ZeroLengthWrite_WhenPositionExceedsLimit_ShouldNotReserveOrGrow` |
| 调用独立预算、跨文件并发预留 | 同上 | `CreateFile_WhenUsedConcurrently_ShouldShareOnlyItsOwnBudget` |
| `NpoiEntityBatchImporter`、`ClosedXmlEntityBatchImporter` 与输入、共享字符串的同次预算传递 | 同上，真实 DI 入口 | `Xlsx_ShouldShareInputAndStringStorageBudget`、`SeekableInput_ShouldChargeOnlySharedStrings` |
| Buffered SummaryOnly 选项快照保留预算 | 同上 | `Xlsx_ShouldShareInputAndStringStorageBudget(mode: 2)` |
| 不受预算控制的组合预检 | 同上 | `UnsupportedCombination_ShouldRejectBeforeReading` |
| 回调异常／取消、输入 IO 取消清理与所有权 | 同上 | `CallbackFailure_ShouldCleanup`、`InputCancellation_ShouldCleanupBudgetedFiles` |
| `NpoiHssfForwardXlsxConverter` XLSX 投影及逐 Sheet XML 预算、异常释放 | NPOI.Tests，双 TFM | `Hssf_ShouldChargeConvertedPackageAndSheetFiles` |
| 输入中途 IO 异常清理与原样传播 | 同上，共享合同 | `InputFailure_ShouldCleanupAndPreserveException` |
| `ExcelForwardXlsxSharedStrings.ReadDiskBacked` 清理异常不覆盖解析异常，继续释放另一文件 | NPOI.Tests、ClosedXml.Tests，双 TFM | `ReadDiskBacked_WhenCleanupThrows_ShouldReleaseBothFilesAndPreserveParseError` |

共享测试文件：`tests/provider-shared/{ExcelTemporaryStorageBudgetTest,EntityTemporaryStorageBudgetContractTest}.cs`；每个 Provider 有独立具体测试类，均调用真实注册服务，不仅测试 helper。

## 续轮测试基础设施

| 修改符号 | 直接验证 | 结果 |
| --- | --- | --- |
| SourceLayoutTests.GetRepositoryDirectory | AbstractionsSources_ShouldUseCapabilityDirectories、AbstractionsPublicTypes_ShouldKeepCompatibleNamespaces、NpoiEntityExecutors_ShouldUseEntityNamespace，在仓库外输出启动 | net8.0 9/9 |
| PublicApiContractTest.FindRepositoryRoot | PublicApi_ProductionAssemblies_ShouldNotExposeProductionFriendAssemblies | net6.0/net8.0 各 1/1 |
| 同上，正式 API 输入校验不被绕过 | PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot | 双 TFM 找到真实 baseline 后明确失败于统一 Release 缺失输入；保持门禁未关闭 |

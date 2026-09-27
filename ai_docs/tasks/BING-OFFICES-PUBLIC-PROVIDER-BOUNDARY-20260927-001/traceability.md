# 生产符号 → 测试方法

所有行均验证 net6.0 和 net8.0；Docs 项目仅 net8.0。新增成员的完整规范化签名见 api-diff.json；两个快照记录最终公共表面。

| 生产符号 / 边界 | 关键行为 | 测试项目与方法 |
| --- | --- | --- |
| Core AssemblyInfo；八个生产程序集 IVT | 精确测试白名单，无生产友元；无友元的程序集也受控 | Bing.Offices.Tests.Integration / PublicApiContractTest.PublicApi_ProductionAssemblies_ShouldNotExposeProductionFriendAssemblies |
| NpoiExcelImporter 五参数 / 六参数构造；ClosedXmlExcelImporter 五参数 / 六参数构造 | 旧参数类型、顺序、默认值不变；新增提交参数，全部必传 | Bing.Offices.Tests.Integration / PublicApiContractTest.PublicApi_ImporterConstructors_ShouldPreserveCompatibility |
| NpoiExcelImporter.Import / ImportAsync；NpoiFailureWorkbookWriter.Write | XLS/HSSF 与 XLSX/XSSF 单次提交、路径、格式、令牌与完整工作簿 | Bing.Offices.Npoi.Tests / NpoiFailureWorkbookCommitterTest.InjectedCommitter_ShouldCommitCompleteWorkbookExactlyOnce |
| NpoiExcelImporter 构造链 | 旧入口、新入口 null 回退默认实现 | Bing.Offices.Npoi.Tests / NpoiFailureWorkbookCommitterTest.NullCommitter_ShouldUseDefaultWithoutChangingOldConstructor |
| ExcelNpoiServiceCollectionExtensions.AddBingOfficesNpoi | 保留宿主提交器、重复注册幂等 | Bing.Offices.Npoi.Tests / NpoiFailureWorkbookCommitterTest.Registration_ShouldPreserveHostCommitter |
| NpoiExcelImporter.Import / ImportAsync | 有效输入、禁用失败产物、输出流不调用提交器 | Bing.Offices.Npoi.Tests / NpoiFailureWorkbookCommitterTest.NonFileOrNoFailure_ShouldNotCallCommitter |
| NpoiExcelImporter；NpoiFailureWorkbookWriter | 失败/取消不损坏旧文件、清理临时文件、错误只观察一次 | Bing.Offices.Npoi.Tests / NpoiFailureWorkbookCommitterTest.CommitFailure_ShouldPreserveExceptionAndObserveOnce；WriteFailure_ShouldPreserveTargetAndObserveOnce；CancellationDuringWrite_ShouldPreserveTargetAndCleanup；PreCanceled_ShouldNotCallCommitter |
| ClosedXmlExcelImporter.Import / ImportAsync | XLSX 单次提交、路径、格式、令牌与完整工作簿 | Bing.Offices.ClosedXml.Tests / ClosedXmlFailureWorkbookCommitterTest.InjectedCommitter_ShouldCommitCompleteWorkbookExactlyOnce |
| ClosedXmlExcelImporter 构造链 | 旧入口、新入口 null 回退默认实现 | Bing.Offices.ClosedXml.Tests / ClosedXmlFailureWorkbookCommitterTest.NullCommitter_ShouldUseDefaultWithoutChangingOldConstructor |
| ExcelClosedXmlServiceCollectionExtensions.AddBingOfficesClosedXml | 宿主提交器不被覆盖；仍传入 DOM 准入实例 | Bing.Offices.ClosedXml.Tests / ClosedXmlFailureWorkbookCommitterTest.Registration_ShouldPreserveHostCommitter；ClosedXmlProviderTest.WorkbookAdmission_ShouldRejectWhenQueueCapacityIsExceeded；WorkbookAdmissionAsync_ShouldWaitCancelAndRelease |
| ClosedXmlExcelImporter.Import / ImportAsync | 非文件、无错误、禁用路径均不提交 | Bing.Offices.ClosedXml.Tests / ClosedXmlFailureWorkbookCommitterTest.NonFileOrNoFailure_ShouldNotCallCommitter |
| ClosedXmlExcelImporter.Import / ImportAsync | 失败/取消保留旧文件及原异常、清理临时文件、单次观察 | Bing.Offices.ClosedXml.Tests / ClosedXmlFailureWorkbookCommitterTest.CommitFailure_ShouldPreserveExceptionAndObserveOnce；WriteFailure_ShouldPreserveTargetAndObserveOnce；CancellationDuringWrite_ShouldPreserveTargetAndCleanup；PreCanceled_ShouldNotCallCommitter |
| DefaultFileExportCommitter.Commit / CommitAsync（实现未改） | 默认真实文件提交、替换、失败保护、取消 | Bing.Offices.Tests / DefaultFileExportCommitterTest（整类）；ReviewFixCoreRegressionTest.AtomicFileCommitter_* |
| NPOI / ClosedXML 默认失败产物路径 | AnnotatedOriginal / ErrorRowsOnly、真实锁定目标拒绝提交、序列化失败与流边界 | Bing.Offices.ProviderContract.Tests / FailureWorkbookFailureBoundaryContractTest.PathSuccess_ShouldCreateAndReplaceCompleteWorkbook；CommitFailure_ShouldPreserveTargetAndCleanTemporaryFiles；SerializationFailure_ShouldPreserveTargetAndCleanTemporaryFiles；StreamPartialWriteFailure_ShouldPreserveOwnership |
| IFileExportCommitter / DefaultFileExportCommitter / 两个新增构造重载 | 仅 NuGet 引用的外部消费者可编译、运行、读回失败产物 | Bing.Offices.ThirdPartyProvider.Consumer / Program.VerifyPublicFileCommitters；成功标记 third-party-public-only-provider-ok |
| 全部公开成员 | 双 TFM 只新增批准的两个构造，无删除 | Bing.Offices.Tests.Integration / PublicApiContractTest.PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot；build/ApiSnapshot 比较 |

## 迁移与审批

- 新增构造重载归类 Provider User API。用户实施指令已明确批准成员级变化，见 api-approval.md。
- 公开五参数签名与默认值保持不变；旧 Provider 依赖 Core 内部入口的二进制需要与 Core 配套升级。
- 文件提交服务、原子算法、资源预算均复用现有实现；未新增公共接口或公开内部工具。

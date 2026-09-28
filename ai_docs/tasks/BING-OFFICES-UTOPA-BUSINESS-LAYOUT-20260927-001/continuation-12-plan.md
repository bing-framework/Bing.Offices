# Continuation 12：可复用 Entity Layout Provider 合同源码模板

## 范围与交付

在 `tests/Bing.Offices.Testing` 增加单文件、无测试框架依赖的 Entity Layout 合同模板，输入仅为公开的 `IExcelEntityExporter` 与 `IExcelEntityImporter`。它分别验证动态列分组和最终 Footer、分页小计、连续分组小计、异步往返及预取消与调用方流所有权。失败携带场景上下文；既有能力档案继续独立维护，模板不根据 Provider 自述能力自动降低断言。

仓库 Provider Contract 使用模板对 NPOI/ClosedXML 执行；第三方包消费者通过源码链接同一文件，在 net6.0/net8.0 只引用 NuGet 包进行验证。文档给出第三方接入方式和 XLSX 前提。测试模板是源码模板，不增加生产 API、不发布测试包、不更改版本号；NPOI XLS/HSSF 的专属格式测试仍留在 Provider 职责测试。

## 影响与门禁

- ChangedProjects：Testing、ProviderContract.Tests、ThirdPartyProvider.Consumer；Docs。生产源码、API 快照与 Provider 能力声明不变。
- ChangedRuntimePaths：仅测试运行，不触及生产执行路径。ChangedBuildPackaging：Consumer 链接模板源码并从当前 NuGet 包运行。
- L1：新合同模板在 NPOI/ClosedXML 双 TFM 直接运行，并验证失败诊断。L2：Provider Contract 常规测试。L3：Consumer 还原、双 TFM 构建运行；Docs 测试与 Release build。对未变的生产路径复用上一阶段 Provider、API 与全量测试证据，不运行性能矩阵。
- 交付阶段执行记录、生产符号到测试方法追溯、严格 UTF-8/BOM/EOL 和 `git diff --check`。不自动 commit、push、PR 或发布。

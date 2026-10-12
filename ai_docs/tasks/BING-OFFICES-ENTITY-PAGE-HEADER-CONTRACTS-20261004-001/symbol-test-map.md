# 生产符号到测试方法

本阶段没有修改生产符号。新增的可复用测试入口验证下列既有生产能力；NPOI 的 XLS/HSSF 专项场景仍由其 Provider 测试负责。

| 最终生产符号 | 关键行为 | 测试项目与方法 |
|---|---|---|
| `ExcelEntityListRegionBuilder<TItem>.PageHeader`、`ExcelEntityPageHeaderContext<TItem>` | 空、单、多页布局，页码和复制页表头 | `Bing.Offices.ProviderContract.Tests.EntityLayoutReusableContractTest.PageHeaders_ShouldSatisfyReusableContract`（NPOI、ClosedXML） |
| `NpoiEntityExportExecutor`、`ClosedXmlEntityLayoutExecutor` | 占用冲突在写入前拒绝，调用方流保持原状 | `Bing.Offices.ProviderContract.Tests.EntityLayoutReusableContractTest.PageHeaders_ShouldSatisfyReusableContract`（NPOI、ClosedXML） |
| `NpoiEntityImportExecutor`、`ClosedXmlEntityLayoutExecutor` | 导入跳过复制页眉；标记不匹配返回原始坐标 | `Bing.Offices.ProviderContract.Tests.EntityLayoutReusableContractTest.PageHeaders_ShouldSatisfyReusableContract`（NPOI、ClosedXML） |
| `NpoiExcelImporter.ImportEntity`、`ClosedXmlExcelImporter.ImportEntity` | 输入字节限制返回资源错误并保留调用方流 | `Bing.Offices.ProviderContract.Tests.EntityLayoutReusableContractTest.PageHeaderInputLimit_ShouldRejectBeforeRead`（NPOI、ClosedXML） |

可复用入口为 `Bing.Offices.Testing.Entities.EntityLayoutProviderContractSuite.VerifyPageHeaders`；`Bing.Offices.ThirdPartyProvider.Consumer` 从 NuGet 包在 net6.0 和 net8.0 调用该入口。引擎特有的 XLS/XLSX、模板样式、合并和格式边界继续由 `Bing.Offices.Npoi.Tests.EntityRepeatedPageHeaderTest` 与 `Bing.Offices.ClosedXml.Tests.EntityRepeatedPageHeaderTest` 覆盖。

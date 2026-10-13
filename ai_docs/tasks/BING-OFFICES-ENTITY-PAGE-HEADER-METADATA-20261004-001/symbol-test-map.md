# 生产符号到测试方法

| 最终生产符号 | 关键行为 | 测试项目与方法 |
|---|---|---|
| `NpoiEntityExportExecutor.ValidatePageHeaderSourceMetadata` | XLS/XLSX 多页源元数据拒绝、单页保留 | `Bing.Offices.Npoi.Tests.EntityRepeatedPageHeaderTest.PageHeader_SourceMetadata_ShouldRejectMultiPageBeforeOutput`、`PageHeader_SourceMetadata_ShouldRemainForSinglePage` |
| `NpoiEntityExportExecutor.ValidateConfiguredPageHeaderValidations` | 配置式数据校验与页眉源区域相交时预检 | `Bing.Offices.Npoi.Tests.EntityRepeatedPageHeaderTest.PageHeader_ConfiguredValidationOverlap_ShouldRejectBeforeOutput` |
| `NpoiEntityExportExecutor.WritePageHeader` | 隐藏行及单元格锁定样式复制 | `Bing.Offices.Npoi.Tests.EntityRepeatedPageHeaderTest.PageHeader_HiddenRowAndUnlockedStyle_ShouldCopyAcrossPages` |
| `ClosedXmlEntityLayoutExecutor.ValidatePageHeaderSourceMetadata` | 多页源元数据拒绝、单页保留 | `Bing.Offices.ClosedXml.Tests.EntityRepeatedPageHeaderTest.PageHeader_SourceMetadata_ShouldRejectBeforeOutput`、`PageHeader_SinglePage_ShouldPreserveTemplateMetadata` |
| `ClosedXmlEntityLayoutExecutor.ValidateConfiguredPageHeaderValidations` | 配置式数据校验与页眉源区域相交时预检 | `Bing.Offices.ClosedXml.Tests.EntityRepeatedPageHeaderTest.PageHeader_ConfiguredValidationOnSource_ShouldRejectMultiPage` |
| `ClosedXmlEntityLayoutExecutor.WritePageHeader` | 隐藏行及单元格锁定样式复制 | `Bing.Offices.ClosedXml.Tests.EntityRepeatedPageHeaderTest.PageHeader_HiddenRowAndUnlockedCell_ShouldCopyToNextPage` |

上述用例均在 net6.0 和 net8.0 运行；NPOI 的参数化用例分别覆盖 HSSF 与 XSSF。跨 Provider 常规合同由 `Category!=Large` 解决方案测试验证。

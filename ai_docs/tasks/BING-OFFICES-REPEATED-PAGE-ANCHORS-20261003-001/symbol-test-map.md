# 最终生产符号到测试方法追溯

| 生产符号及关键行为 | 测试项目 | 测试方法 |
|---|---|---|
| `ExcelEntityListRegionBuilder<TItem>.PageSubtotalNamed`：连续分页名称、导入往返 | Bing.Offices.Npoi.Tests；Bing.Offices.ClosedXml.Tests | `EntityRepeatedPageAnchorTest.PageSubtotalNamed_MultiplePages_ShouldCreateSequentialNamesAndRoundTrip` |
| `ExcelEntityListRegion<TItem>.PageSubtotalAnchorPrefix`：Provider 读取布局描述；空或单页不创建名称 | Bing.Offices.Npoi.Tests；Bing.Offices.ClosedXml.Tests | `EntityRepeatedPageAnchorTest.PageSubtotalNamed_EmptyOrSinglePage_ShouldNotCreateName` |
| `ExcelEntityLayoutBuilder<T>.Build`：非法前缀在布局阶段失败 | Bing.Offices.Npoi.Tests；Bing.Offices.ClosedXml.Tests | `EntityRepeatedPageAnchorTest.PageSubtotalNamed_InvalidPrefix_ShouldRejectLayout` |
| `NpoiEntityExportExecutor`：XLS/XLSX 逐页名称与分组同边界布局 | Bing.Offices.Npoi.Tests | `EntityRepeatedPageAnchorTest.PageSubtotalNamed_WithGroupSubtotal_ShouldFollowGroupMarker` |
| `ClosedXmlEntityLayoutExecutor`：XLSX 逐页名称与分组同边界布局 | Bing.Offices.ClosedXml.Tests | `EntityRepeatedPageAnchorTest.PageSubtotalNamed_WithGroupSubtotal_ShouldFollowGroupMarker` |
| 两个 Provider：模板名称迁移、局部作用域和非法/过期名称拒绝 | Bing.Offices.Npoi.Tests；Bing.Offices.ClosedXml.Tests | `EntityRepeatedPageAnchorTest.PageSubtotalNamed_TemplateNames_ShouldRelocateAndPreserveScope`；`PageSubtotalNamed_StaleTemplateName_ShouldPreserveOutput`；`PageSubtotalNamed_InvalidTemplateName_ShouldPreserveOutput` |
| 两个 Provider：异步写入一致性 | Bing.Offices.Npoi.Tests；Bing.Offices.ClosedXml.Tests | `EntityRepeatedPageAnchorTest.PageSubtotalNamed_AsyncExport_ShouldCreateSequentialNames` |
| 打包公共 API：第三方消费者调用并回读 | Bing.Offices.Consumer.Net6；Bing.Offices.Consumer.Net8 | `Program` 的 `PageSubtotalNamed` 包消费场景（进程输出 `package-consumer-ok`） |

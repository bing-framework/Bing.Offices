# 生产符号与测试追溯

| 生产符号 | 关键行为 | 测试项目与方法 |
|---|---|---|
| `ExcelEntityListRegionBuilder<TItem>.GroupSubtotal`、`PageBreak`、`PageSubtotal` | 三项配置可组合，页容量只计算明细 | Npoi.Tests、ClosedXml.Tests：`EntityGroupPaginationCombinationTest.ExportAndImport_GroupPagination_ShouldPreserveDetailsAndPhysicalRows` |
| `ExcelEntityLayoutBuilder<TEntity>.Build` | 组合中间小计拒绝连续行求和公式 | Npoi.Tests、ClosedXml.Tests：`EntityGroupPaginationCombinationTest.Layout_GroupPagination_ShouldRejectContiguousIntermediateFormula` |
| `ExcelEntityRowPlan.Create` | 组跨页、同边界顺序、分页符、空明细及多行小计 | Npoi.Tests、ClosedXml.Tests：`ExportAndImport_GroupPagination_ShouldPreserveDetailsAndPhysicalRows`、`Export_GroupAndPageSubtotalAtSameBoundary_ShouldWriteGroupFirst`、`ExportAndImport_EmptyDetails_ShouldWriteOnlyFinalFooter`、`ExportAndImport_GroupPaginationWithMultiRowFooters_ShouldKeepCoordinates` |
| `NpoiEntityExportExecutor`、`ClosedXmlEntityLayoutExecutor` | XLS/XLSX 导出、边界预检、无页小计分页 | Npoi.Tests、ClosedXml.Tests：`ExportAndImport_GroupPaginationWithoutPageSubtotal_ShouldKeepDetailRows`、`Export_GroupPaginationBeyondRegionEnd_ShouldFailBeforeWriting` |
| `ExcelEntitySubtotalImportScanner.FindSpans` | 跳过组合小计，保留原始行号，拒绝缺失或倒序 marker | Npoi.Tests、ClosedXml.Tests：`ExportAndImport_GroupPaginationWithMultiRowFooters_ShouldKeepCoordinates`、`Import_GroupPaginationWithMissingPageMarker_ShouldRejectWorkbook`、`Import_GroupPaginationWithReversedMarkers_ShouldRejectWorkbook` |
| 公开程序集 API | 两个 TFM 的已审批成员集合保持一致 | Tests.Integration：`PublicApiContractTest.PublicApi_ReleaseAssemblies_ShouldMatchApprovedBaseline` |
| 文档 C# 示例 | 新组合示例可编译和执行 | Docs.Tests：`DocsConsumerTest.DocumentationFences_FromMarkdown_ShouldCompileAndExecuteIndividually` |

NPOI 组合测试使用 XLS/HSSF、XLSX/XSSF 两种格式；NPOI 与 ClosedXML 测试均在 net6.0 和 net8.0 运行。

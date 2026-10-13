# 生产符号到测试方法

| 最终生产符号 | 关键行为 | 测试项目与方法 |
|---|---|---|
| `NpoiEntityImportExecutor.FindLastPageHeaderRegionRow` | 稀疏表和大 `End` 下定位自身列表末行 | `Bing.Offices.Npoi.Tests.EntityRepeatedPageHeaderTest.PageHeader_SparseLargeEnd_ShouldRoundTrip` |
| `NpoiEntityLayoutSupport.FindFooterSkipSpans`、`ExcelEntitySubtotalImportScanner.FindSpans` | 页眉 marker 扫描和取消传递 | `Bing.Offices.Npoi.Tests.EntityRepeatedPageHeaderTest.PageHeader_SparseLargeEnd_CancelledImport_ShouldThrow`、`PageHeader_SparseLargeEnd_ShouldRoundTrip`；`Bing.Offices.ClosedXml.Tests.EntityRepeatedPageHeaderTest.PageHeader_LongMarkerScan_ShouldObserveMidScanCancellation` |
| `ClosedXmlEntityLayoutExecutor.FindLastPageHeaderRegionRow` | 稀疏表和大 `End` 下定位自身列表末行 | `Bing.Offices.ClosedXml.Tests.EntityRepeatedPageHeaderTest.PageHeader_SparseLargeEnd_ShouldIgnoreUnrelatedDistantCell` |
| `ClosedXmlEntityLayoutExecutor.FindFooterSkipSpans`、`ExcelEntitySubtotalImportScanner.FindSpans` | 页眉 marker 扫描和取消传递 | `Bing.Offices.ClosedXml.Tests.EntityRepeatedPageHeaderTest.PageHeader_SparseLargeEnd_ShouldIgnoreUnrelatedDistantCell`、`PageHeader_LongMarkerScan_ShouldObserveMidScanCancellation` |
| `ExcelEntityPageHeaderImportValidation.ValidateCopiedHeaders` | 复制页表头校验响应取消 | `Bing.Offices.ClosedXml.Tests.EntityRepeatedPageHeaderTest.PageHeader_CopiedTitleCheck_ShouldObserveMidCheckCancellation` |
| `ExcelImportRuntime`、`ClosedXmlEntityLayoutExecutor.ReadRegion` | 输入字节和明细行资源限制 | `Bing.Offices.Npoi.Tests.EntityRepeatedPageHeaderTest.PageHeader_SparseLargeEnd_InputLimit_ShouldReturnResourceError`、`Bing.Offices.ClosedXml.Tests.EntityRepeatedPageHeaderTest.PageHeader_SparseLargeEnd_ShouldRespectRowLimit` |

NPOI 参数化测试覆盖 XLS 与 XLSX；以上职责测试均在 net6.0、net8.0 执行。

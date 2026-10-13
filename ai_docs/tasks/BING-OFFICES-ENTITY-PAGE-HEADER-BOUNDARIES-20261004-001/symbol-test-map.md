# 生产符号到测试方法

| 最终生产符号或行为 | 测试项目 | 测试方法 |
|---|---|---|
| `ExcelEntityPageHeaderImportValidation.ValidateMarkerColumn` | Npoi.Tests；ClosedXml.Tests | `PageHeader_HeaderlessMarkerInsideDetailColumns_ShouldRejectBeforeOutput`；`PageHeader_WithoutListHeader_ShouldRequireDedicatedMarkerColumn` |
| `ExcelEntityPageHeaderImportValidation.ValidateCopiedHeaders` | Npoi.Tests；ClosedXml.Tests | `PageHeader_CopiedListHeaderMismatch_ShouldReportSourceCoordinate`；`PageHeader_CopiedTitleTampered_ShouldReportOriginalCell` |
| `NpoiEntityExportExecutor.ValidateWrittenDetailMarker` | Npoi.Tests | `PageHeader_DetailValueMatchingMarker_ShouldRejectBeforeOutput`（XLS、XLSX） |
| `ClosedXmlEntityLayoutExecutor.ValidateWrittenDetailMarker` | ClosedXml.Tests | `PageHeader_DetailMarkerCollision_ShouldPreserveTarget`；`PageHeader_DynamicOrCalculatedMarkerCollision_ShouldReject` |
| NPOI 页眉标题与合并边界 | Npoi.Tests | `PageHeader_FluentCellWithDifferentListTitle_ShouldReject`；`PageHeader_FluentCellWithMatchingListTitle_ShouldRoundTrip`；`PageHeader_MergeOverlappingListHeader_ShouldReject`；`PageHeader_MergeOutsideListHeader_ShouldRoundTrip` |
| ClosedXML 页眉标题与合并边界 | ClosedXml.Tests | `PageHeader_FluentTitleOverride_ShouldRejectButMatchingTitleRoundTrips`；`PageHeader_FluentMerge_ShouldRejectHeaderAndAllowDocumentTitle` |

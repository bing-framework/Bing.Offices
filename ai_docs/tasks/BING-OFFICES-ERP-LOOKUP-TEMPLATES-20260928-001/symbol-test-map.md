# 生产符号与测试追溯

| 生产符号 / 行为 | 测试项目 | 测试方法 |
|---|---|---|
| `ExcelWorkbookExportBuilder.LookupList`、`ExcelLookupListDefinition` 快照与边界 | Bing.Offices.Npoi.Tests | `LookupListModelTest.LookupList_ShouldNormalizeAndSnapshotValues`; `LookupListModelTest.LookupList_ShouldRejectInvalidDefinitions` |
| `ExcelDataValidationDefinition.ListSourceName` 互斥约束 | Bing.Offices.Npoi.Tests | `LookupListModelTest.NamedValidation_ShouldRequireExactlyOneSource` |
| NPOI XLS/XLSX 字典 Sheet、名称和下拉 | Bing.Offices.Npoi.Tests | `LookupListExportStructureTest.Export_LookupList_ShouldWriteHiddenNamedRangeAndSnapshot` |
| ClosedXML XLSX 字典 Sheet、名称和下拉 | Bing.Offices.ClosedXml.Tests | `LookupListExportStructureTest.Export_LookupList_ShouldWriteHiddenNamedRangeAndListValidation` |
| 混合内联/名称下拉、特殊字典文本 | Bing.Offices.Npoi.Tests; Bing.Offices.ClosedXml.Tests | `LookupListExportStructureTest.Export_MixedLists_ShouldKeepNamedFormulaAndLiteralText` |
| NPOI 名称作用域、区域形态、Unsupported 和公式字典 | Bing.Offices.Npoi.Tests | `NamedListValidationTest.NamedList_ShouldResolveScopesAndBoundedRanges`; `NamedListValidationTest.NamedList_InvalidLocalName_ShouldNotFallBackToGlobal`; `NamedListValidationTest.NamedList_UnsupportedDefinitions_ShouldReturnStructuredError`; `NamedListValidationTest.NamedList_FormulaCell_ShouldReturnStructuredError` |
| Entity 原生规则模式、默认兼容、固定字段和明细坐标 | Bing.Offices.Npoi.Tests; Bing.Offices.ClosedXml.Tests | `EntityNativeValidationTest.EntityImport_NamedList_ShouldValidateCellsAndKeepDefaultMode` |
| Entity 字典模板真实往返 | Bing.Offices.Npoi.Tests; Bing.Offices.ClosedXml.Tests | `EntityNativeValidationTest.EntityExport_LookupList_ShouldRoundTripTemplate` |
| 打包后的公开 API 调用 | Bing.Offices.Consumer.Net6; Bing.Offices.Consumer.Net8 | `Program.cs` 中 `LookupList`、`DataValidation` 与 `ExcelEntityImportOptions` 往返检查 |
| NPOI XLSX 内联规则 XML 修正不覆盖名称规则 | Bing.Offices.Npoi.Tests | `LookupListExportStructureTest.Export_OverlappingLists_ShouldPreserveNamedFormula` |
| 请求名称前向引用与一维预检 | Bing.Offices.Npoi.Tests; Bing.Offices.ClosedXml.Tests | `LookupListExportStructureTest.Export_NamedList_ShouldResolveLaterSheet`; `LookupListExportStructureTest.Export_TwoDimensionalNamedList_ShouldLeaveOutputEmpty` |
| 合法特殊工作表名解析与单工作簿名称缓存 | Bing.Offices.Npoi.Tests; Bing.Offices.ClosedXml.Tests | `NamedListValidationTest.NamedList_QuotedSheetNameWithSpecialCharacters_ShouldImport`; `NamedListValidationTest.NamedList_QuotedSheetNameWithSpecialCharacters_ShouldValidate`; `NamedListValidationTest.NamedList_ShouldCacheSourcePerWorkbook`; `NamedListValidationTest.NamedList_ShouldCacheValuesPerWorkbook` |
| Entity 原生校验行优先和分页小计原始坐标 | Bing.Offices.Npoi.Tests; Bing.Offices.ClosedXml.Tests | `EntityNativeValidationTest.EntityImport_InvalidNativeValue_ShouldSkipLaterConversionInSameRow`; `EntityNativeValidationTest.EntityImport_PageSubtotalNamedList_ShouldReportOriginalRow` |
| 采购类别、仓库、状态字典与业务编码映射 | Bing.Offices.Npoi.Tests; Bing.Offices.ClosedXml.Tests | `EntityNativeValidationTest.PurchaseHeader_LookupLists_ShouldRoundTripValueMappings` |
| 模板名称冲突、辅助 Sheet 顺延和长文本字典 | Bing.Offices.Npoi.Tests; Bing.Offices.ClosedXml.Tests | `LookupListExportStructureTest.Export_LookupList_TemplateNameCollision_ShouldLeaveOutputEmpty`; `LookupListExportStructureTest.Export_LookupList_ShouldSkipExistingHelperSheet`; `LookupListExportStructureTest.Export_LookupList_LongText_ShouldPreserveFullValue` |
| 字典文件导出预取消与提交失败清理 | Bing.Offices.Npoi.Tests; Bing.Offices.ClosedXml.Tests | `LookupListExportStructureTest.Export_LookupList_PreCanceled_ShouldPreserveTarget`; `LookupListExportStructureTest.Export_LookupList_CommitFailure_ShouldCleanTemporaryFile` |
| 不支持 Provider 对字典和名称规则明确拒绝 | Bing.Offices.MiniExcel.Tests; Bing.Offices.SpreadCheetah.Tests | `MiniExcelProviderTest.LookupListRequests_ShouldFailBeforeWritingToStream`; `MiniExcelProviderTest.LookupListRequests_ShouldPreserveExistingFile`; `SpreadCheetahStreamingExcelExporterTest.LookupListRequests_ShouldFailBeforeWritingToStream`; `SpreadCheetahStreamingExcelExporterTest.LookupListRequests_ShouldPreserveExistingFile` |
| NPOI HSSF/XSSF 显式列表导入及 Workbook/Entity 原始错误坐标 | Bing.Offices.Npoi.Tests | `InlineListImportTest.WorkbookImport_ExplicitList_ShouldAcceptAndRejectValues`; `InlineListImportTest.EntityImport_ExplicitList_ShouldValidateFixedAndDetailCells` |
| NPOI 名称列表扫描取消令牌 | Bing.Offices.Npoi.Tests | `NamedListCancellationTest.ValidateCell_NamedListCanceledDuringScan_ShouldThrow` |
| 字典异步导出完整结构与中途取消资源保护 | Bing.Offices.Npoi.Tests; Bing.Offices.ClosedXml.Tests | `LookupListExportStructureTest.ExportAsync_LookupList_ShouldWriteReadableWorkbook`; `LookupListExportStructureTest.ExportAsync_LookupList_MidFlightCancellation_ShouldPreserveTargetAndCleanTemporaryFile` |
| NPOI 联合区域及外部引用名称拒绝 | Bing.Offices.Npoi.Tests | `NamedListValidationTest.NamedList_UnsupportedDefinitions_ShouldReturnStructuredError`; `NamedListValidationTest.NamedList_ExternalReference_ShouldReturnStructuredError` |

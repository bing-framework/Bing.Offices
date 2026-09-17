# Symbol Test Map

- Task ID：`BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001`

| 最终生产符号/职责 | 直接测试方法 | TFM 证据 |
| --- | --- | --- |
| `MiniExcelExcelExporter.Export` / `ExportAsync` | `RoundTrip_ShouldSupportMultipleSheetsAndMapping`、`AsyncRoundTrip_ShouldUseMiniExcelAsyncApi` | MiniExcel 专项 25/25，net6/net8 |
| `MiniExcelExcelExporter.ExportToFileAsync` | `AsyncFileExport_PreCanceled_ShouldPreserveExistingTargetAndCleanTempFile`、`AsyncFileExport_MidFlightCancellation_ShouldPreserveExistingTargetAndCleanTempFile`；集成 `ExportToFileAsync_ThenImportFromFileAsync_ShouldRoundTripRealXlsx` | 专项/Integration 双 TFM |
| `MiniExcelExcelExporter.EnumerateRows` | `FixedColumnConverter_ShouldReceiveOneBasedDataRows`、`DynamicColumns_ShouldRoundTripWithNamedConverter` | MiniExcel 专项双 TFM |
| `MiniExcelExcelImporter.Import` / `ImportAsync` | `RoundTrip_ShouldSupportMultipleSheetsAndMapping`、`AsyncRoundTrip_ShouldUseMiniExcelAsyncApi`、文件集成往返 | 专项/Integration 双 TFM |
| `MiniExcelExcelImporter.ImportTypedSheetRows` | `DataRowStartIndex_ShouldAlignRawDateSerialsAcrossProviders`、`DataRowStartIndexAsync_ShouldAlignRawDateSerials` | 跳过一行/多行、非零表头、固定/动态日期、Converter/Validator RowIndex、SourceRows；专项双 TFM |
| `MiniExcelValueAdapter.ConvertRaw` / `ConvertFrom` | `DateSerial_ShouldUseCoreDateContractAnd1904Flag`、`DateValues_ShouldNormalizeCrossTargetTypesThroughCoreContract`、`RealXlsxDateSystems_ShouldUseCoreSerialContract` | MiniExcel 专项双 TFM；真实 1900/1904 fixture 同时覆盖 DateTime/固定 offset DateTimeOffset 与 NPOI 对比 |
| `MiniExcelRawDateSerialReader.Read` / `CreateKey` | `DataRowStartIndex_ShouldAlignRawDateSerialsAcrossProviders`、`DataRowStartIndexAsync_ShouldAlignRawDateSerials`、`RealXlsxDateSystems_ShouldUseCoreSerialContract` | 真实 worksheet XML 数值 serial 按物理行列对齐，双 TFM |
| `ExcelDateParser.TryParse` | `TryParse_NativeAndSerialValues_ShouldUseRawValueAndDateSystem`、`TryParseValidation_Time_ShouldUseStableDateAndSerialValue` | Core unit 双 TFM |
| `MiniExcelXlsxPreflight` / shared `ExcelXlsxZipPreflight` | `XlsxDate1904Preflight_ShouldReadWorkbookPropertyAndRestorePosition`、`ResourceLimit_ShouldFailBeforeMiniExcelParser`；NPOI preflight 27/27 | 专项和 preflight 双 TFM |
| unsupported capability gate | `Xls_ShouldBeRejectedAsUnsupported`、`UnsupportedPresentationOptions_ShouldFailBeforeWriting`、`FailureWorkbook_ShouldBeRejectedBeforeParser` | MiniExcel 专项双 TFM |
| async cancellation / stream ownership | `PreCanceledAsync_ShouldNotWriteOrRead`、`AsyncExport_MidWriteCancellation_ShouldKeepDestinationOpen`、`AsyncImport_MidReadCancellation_ShouldKeepSourceOpen`、`AsyncFileExport_MidFlightCancellation_ShouldPreserveExistingTargetAndCleanTempFile` | MiniExcel 专项双 TFM |
| DI registration | `AddBingOfficesMiniExcel_ShouldRegisterProviderServices` | MiniExcel 专项双 TFM |
| relation comparer | `Relations_ShouldHonorCustomKeyComparer` | MiniExcel 专项双 TFM |
| `ExcelImportValidationMode` | `ValidationMode_ShouldDisableConfiguredRulesAndRejectWorkbookRules` | MiniExcel 专项双 TFM |
| MiniExcel direct ValueMap/Validation/Unique/1904 | `ValueMap_ShouldConvertConfiguredTextToTargetValue`、`ValidationMode_ShouldDisableConfiguredRulesAndRejectWorkbookRules`、`Unique_ShouldRejectDuplicateRowsAndPreserveFirstRow`、`RealXlsx1904Fixture_ShouldPassMiniExcelImportPreflight` | MiniExcel 专项双 TFM |
| NPOI/MiniExcel common contract | `CommonScalarAndDynamicContract_ShouldMatchAcrossProviders`、`ValueMapConverterAndValidationContract_ShouldMatchAcrossProviders`、`StructuredErrorContract_ShouldMatchAcrossProviders`、`RelationContract_ShouldMatchAcrossProviders`、`UnsupportedFeatureContract_ShouldFailFastWithoutOutput` | cross-provider 5/5 双 TFM |

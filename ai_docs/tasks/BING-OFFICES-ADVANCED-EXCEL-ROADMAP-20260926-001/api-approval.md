# Public API / Provider Package Approval

- `approvedBy`: user (implementation-plan approval)
- `approvedAt`: `2026-09-26`
- `taskId`: `BING-OFFICES-ADVANCED-EXCEL-ROADMAP-20260926-001`
- `scope`: additive dual-TFM reporting, formula, rendering, conversion and optional Provider APIs

## Approved Additions

- `ExcelFormat.Xlsm` and `ExcelFormat.Ods` are appended after the existing values; `Xls`, `Xlsx` and `Xlsb` numeric values remain unchanged.
- `ExcelProviderFeatures` and `IExcelProviderFeatureDescriptor` provide optional operation-level capability declarations without changing existing capability interfaces.
- Common report definitions cover tables, auto filters, freeze panes, conditional formats, named ranges and print layout.
- `IExcelStreamingExporter` and `ExcelStreamingExportOptions` expose forward-only XLSX creation without introducing `IAsyncEnumerable` to `netstandard2.0`.
- `IExcelFormulaProcessor`, `IExcelDocumentRenderer`, `IExcelPageRenderer` and `IExcelDocumentConverter` are independent optional contracts.
- `ExcelPdfCompliance` and `ExcelRenderRequest.PdfCompliance` make PDF/A output selection explicit.
- `ExcelWorkbookOpenOptions`, `ExcelWorkbookSaveOptions` and `ExcelMacroPolicy` make password and macro handling explicit; passwords are not diagnostic data.
- `Bing.Offices.SpreadCheetah` is an optional MIT streaming writer; `Bing.Offices.AsposeCells` is an optional commercial Provider with host-configured license and fonts.

## Compatibility

本轮用户批准开源业务能力计划，授权以下增量，不删除或修改既有签名：

- `ExcelSheetImageDefinition`：Content、Row、Column、Width、Height、OffsetX、OffsetY、Validate，PNG/JPEG、零基位置和 96 DPI 像素尺寸。
- `ExcelDataValidationType`：Integer、Decimal、Date、TextLength、List。
- `ExcelDataValidationDefinition`：Range、Type、Operator、Value1、Value2、Date1、Date2、Values、IgnoreBlanks、ShowErrorMessage、InputTitle、InputMessage、ErrorTitle、ErrorMessage、Validate。
- `ExcelSheetExportRequest.Images`、`DataValidations`。
- `ExcelSheetExportBuilder<T>.Image(ExcelSheetImageDefinition)`、`DataValidation(ExcelDataValidationDefinition)`。
- 上述类型的默认构造器及 init 访问器纳入成员快照追溯。

商业 Provider 不属于本轮验收范围；快照记录其未变更表面不等于商业能力验收。

- No existing public member is removed or changes signature.
- Existing enum numeric values are preserved; unsupported formats and capabilities fail during Provider preflight without automatic fallback.
- Core and Abstractions do not reference Aspose.Cells; the commercial package is independently selectable.
- Both new Provider packages target `net6.0` and `net8.0`; Abstractions remains compatible with `netstandard2.0`.

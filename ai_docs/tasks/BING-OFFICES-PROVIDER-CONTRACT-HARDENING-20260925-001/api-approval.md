# Public API / Entity API Approval

- `approvedBy`: user (chat approval)
- `approvedAt`: `2026-09-25T08:57:30.4092699+08:00`
- `taskId`: `BING-OFFICES-PROVIDER-CONTRACT-HARDENING-20260925-001`
- `baseCommit`: `159e6a9fd1f2bf2da9e679f3c9b6fb3a87edddc1`
- `scope`: additive Public API baseline update for `net6.0` and `net8.0`

## Approved Existing Additions

The following already-implemented additive members are approved for the dual-TFM baseline:

- `Bing.Offices.Exports.ExcelRowHeightOptions`
- `ExcelRowHeightOptions.ExcelRowHeightOptions(double, double)`
- `ExcelRowHeightOptions.HeaderHeight`
- `ExcelRowHeightOptions.BodyHeight`
- `ExcelRowHeightOptions.Validate()`
- `Bing.Offices.Exports.ExcelSheetExportBuilder<T>.RowHeight(ExcelRowHeightOptions)`
- `Bing.Offices.Exports.ExcelSheetExportRequest.RowHeight`
- `Bing.Offices.Entities.ExcelEntityLayoutBuilder<TEntity>.HasMany(...)`
- `Bing.Offices.Entities.ExcelEntityLayout<TEntity>.Relations`
- `Bing.Offices.Imports.ExcelResourceLimits.MaxSheets`
- `Bing.Offices.Imports.ExcelResourceLimits.MaxColumnsPerSheet`
- `Bing.Offices.Imports.ExcelResourceLimits.MaxCells`

## Approved Entity Resource API

The selected design is candidate 1 from `api-diff.md`: a provider-neutral options object plus a separate opt-in provider SPI. The existing `IExcelEntityImporter` is intentionally unchanged.

Approved new public members:

- `Bing.Offices.Entities.ExcelEntityImportOptions`
- `ExcelEntityImportOptions.ExcelEntityImportOptions(ExcelResourceLimits)`
- `ExcelEntityImportOptions.ResourceLimits`
- `Bing.Offices.Entities.IExcelEntityResourceImporter`
- `IExcelEntityResourceImporter.ImportEntity<TEntity>(Stream, ExcelEntityLayout<TEntity>, ExcelEntityImportOptions, CancellationToken)`
- `IExcelEntityResourceImporter.ImportEntityAsync<TEntity>(Stream, ExcelEntityLayout<TEntity>, ExcelEntityImportOptions, CancellationToken)`
- `IExcelEntityResourceImporter.ImportForTemplate<TEntity>(Stream, ExcelEntityLayout<TEntity>, ExcelEntityTemplateOptions, ExcelEntityImportOptions, CancellationToken)`
- `IExcelEntityResourceImporter.ImportForTemplateAsync<TEntity>(Stream, ExcelEntityLayout<TEntity>, ExcelEntityTemplateOptions, ExcelEntityImportOptions, CancellationToken)`
- `ExcelEntityExtensions.ImportEntity<TEntity>(IExcelImporter, Stream, ExcelEntityLayout<TEntity>, ExcelEntityImportOptions, CancellationToken)`
- `ExcelEntityExtensions.ImportEntityAsync<TEntity>(IExcelImporter, Stream, ExcelEntityLayout<TEntity>, ExcelEntityImportOptions, CancellationToken)`
- `ExcelEntityExtensions.ImportForTemplate<TEntity>(IExcelImporter, Stream, ExcelEntityLayout<TEntity>, ExcelEntityTemplateOptions, ExcelEntityImportOptions, CancellationToken)`
- `ExcelEntityExtensions.ImportForTemplateAsync<TEntity>(IExcelImporter, Stream, ExcelEntityLayout<TEntity>, ExcelEntityTemplateOptions, ExcelEntityImportOptions, CancellationToken)`
- `NpoiExcelImporter` implementation of the four `IExcelEntityResourceImporter` members
- `ClosedXmlExcelImporter` implementation of the four `IExcelEntityResourceImporter` members

The new options-aware paths apply `ExcelResourceLimits` to source buffering, workbook preflight, entity row/error collection, synchronous import, asynchronous import, and template import. Providers that only implement the legacy SPI receive a structured `UnsupportedFeature` preflight exception from the Core extension overloads.

## Compatibility And Baseline Rules

## Approved Failure Workbook Path API

- `approvedBy`: user (implementation-plan approval)
- `approvedAt`: `2026-09-25T08:57:30.4092699+08:00`
- `Bing.Offices.Imports.ExcelImportFailureOptions.DestinationPath`

`DestinationPath` is additive and mutually exclusive with the existing caller-owned `Destination` stream.
It is the supported atomic file-output route for Failure Workbook artifacts; Stream output remains compatible
and best-effort after complete serialization.

- No public member is removed or changed in signature.
- Existing `IExcelEntityImporter` implementations remain source and binary compatible.
- No package version, dependency, or target framework is changed by this approval.
- The baseline must be captured from the current Release assemblies and freshly packed product packages for both `net6.0` and `net8.0`.

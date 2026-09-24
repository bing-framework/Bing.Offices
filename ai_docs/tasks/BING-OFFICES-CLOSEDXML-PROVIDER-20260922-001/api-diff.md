# API Diff

## Added

- `Bing.Offices.ClosedXml` provider project and package version property `ClosedXmlPackageVersion` (`[0.105.1]`)。
- Provider public members, all in `Bing.Offices.ClosedXml.dll` only:
  - `Bing.Offices.ClosedXml.Exports.ClosedXmlExcelExporter`
  - `Bing.Offices.ClosedXml.Imports.ClosedXmlExcelImporter`
  - `Bing.Offices.ClosedXml.ClosedXmlProviderOptions`
  - `Bing.Offices.ClosedXml.Extensions.ExcelClosedXmlServiceCollectionExtensions.AddBingOfficesClosedXml(IServiceCollection)`
- Additive Abstractions members required by the provider contract:
  - `ExcelRowHeightOptions.HeaderHeight` / `BodyHeight` and `ExcelSheetExportBuilder<T>.RowHeight(...)` / `ExcelSheetExportRequest.RowHeight`
  - `ExcelResourceLimits.MaxSheets` / `MaxColumnsPerSheet` / `MaxCells`
  - `ExcelEntityLayoutBuilder<TEntity>.HasMany<TParent,TChild,TKey>(...)` and immutable `ExcelEntityLayout<TEntity>.Relations`
- ClosedXML Unit/Integration test projects and Package Consumer references。

### Member-level Provider Diff

The current `net6.0` and `net8.0` snapshots contain the same 29 public members. The original 25 provider members remain below; the four newly captured provider members are listed first. They are `Provider-only` additions in `Bing.Offices.ClosedXml.dll`; they are not additions to the third-party provider SPI.

| Status | Member |
| --- | --- |
| Added / Provider-only | `ClosedXmlProviderOptions()` |
| Added / Provider-only | `ClosedXmlProviderOptions.MaxConcurrentWorkbooks` |
| Added / Provider-only | `ClosedXmlProviderOptions.MaxQueuedOperations` |
| Added / Provider-only | `ExcelClosedXmlServiceCollectionExtensions.AddBingOfficesClosedXml(IServiceCollection, Action<ClosedXmlProviderOptions>)` |
| Added / Provider-only | `ClosedXmlExcelExporter(IEnumerable<IExcelValueConverter>?, IExcelMappingPlanFactory?, IEnumerable<IBingOfficesExceptionObserver>?, IFileExportCommitter?)` |
| Added / Provider-only | `ClosedXmlExcelImporter(IEnumerable<IExcelValidationRule>?, IEnumerable<IExcelValueConverter>?, IEnumerable<INamedExcelValidationRule>?, IExcelMappingPlanFactory?, IEnumerable<IBingOfficesExceptionObserver>?)` |
| Added / Provider-only | `ClosedXmlExcelExporter.Capabilities` |
| Added / Provider-only | `ClosedXmlExcelExporter.ProviderName` |
| Added / Provider-only | `ClosedXmlExcelImporter.Capabilities` |
| Added / Provider-only | `ClosedXmlExcelImporter.ProviderName` |
| Added / Provider-only | `ClosedXmlExcelExporter.Export(ExcelWorkbookExportRequest, Stream, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelExporter.ExportAsync(ExcelWorkbookExportRequest, Stream, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelExporter.ExportToFile(ExcelWorkbookExportRequest, string, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelExporter.ExportToFileAsync(ExcelWorkbookExportRequest, string, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelExporter.Supports(ExcelProviderCapabilities)` |
| Added / Provider-only | `ClosedXmlExcelExporter.ExportEntity<TEntity>(TEntity, ExcelEntityLayout<TEntity>, Stream, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelExporter.ExportEntityAsync<TEntity>(TEntity, ExcelEntityLayout<TEntity>, Stream, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelExporter.ExportEntityToFile<TEntity>(TEntity, ExcelEntityLayout<TEntity>, string, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelExporter.ExportEntityToFileAsync<TEntity>(TEntity, ExcelEntityLayout<TEntity>, string, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelExporter.ExportForTemplate<TEntity>(TEntity, ExcelEntityLayout<TEntity>, ExcelEntityTemplateOptions, Stream, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelExporter.ExportForTemplateAsync<TEntity>(TEntity, ExcelEntityLayout<TEntity>, ExcelEntityTemplateOptions, Stream, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelImporter.Import(Stream, ExcelWorkbookImportRequest<TWorkbook>, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelImporter.ImportAsync(Stream, ExcelWorkbookImportRequest<TWorkbook>, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelImporter.Supports(ExcelProviderCapabilities)` |
| Added / Provider-only | `ClosedXmlExcelImporter.ImportEntity<TEntity>(Stream, ExcelEntityLayout<TEntity>, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelImporter.ImportEntityAsync<TEntity>(Stream, ExcelEntityLayout<TEntity>, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelImporter.ImportForTemplate<TEntity>(Stream, ExcelEntityLayout<TEntity>, ExcelEntityTemplateOptions, CancellationToken)` |
| Added / Provider-only | `ClosedXmlExcelImporter.ImportForTemplateAsync<TEntity>(Stream, ExcelEntityLayout<TEntity>, ExcelEntityTemplateOptions, CancellationToken)` |
| Added / Provider-only | `ExcelClosedXmlServiceCollectionExtensions.AddBingOfficesClosedXml(IServiceCollection)` |

Snapshot evidence: `artifacts/api-current/api-snapshot-net6.0.json` and `api-snapshot-net8.0.json`, each with `memberCount=29` for the ClosedXML assembly. The member list is intentionally recorded separately from the approved baseline comparison so that missing approval does not get mistaken for a source-level breaking change.

## Changed

- Solution project graph、benchmark provider comparison/runtime dependency、package consumer PackageReference。
- API snapshot tooling now includes ClosedXML net6/net8 assembly paths, candidate identity and five-package asset set。
- Public API contract classification/release input now includes the four ClosedXML public provider types, including `ClosedXmlProviderOptions`。

## Removed / Breaking

- 无 Abstractions/Core 公共成员删除或既有 capability bit 重排。
- ClosedXML 类型没有进入公共 Request/Result/SPI；Provider-native types remain assembly-local to the ClosedXML provider package。
- No `Changed`, `Removed`, or `Breaking` member was identified in the current provider-local diff。

## Approval status

- `build/api-snapshot-baseline.json` 未修改。Current Public API contract compare reports the new `Bing.Offices.ClosedXml` assembly is missing from the approved baseline on both `net6.0` and `net8.0`；the additive Abstractions members (`ExcelRowHeightOptions`、resource limits、Entity Relations) also produce an expected member/hash mismatch until maintainer approval. These are `BLOCKED_APPROVAL`, not implementation failures。
- Candidate Identity self-test and API snapshot build pass; baseline update requires member-level maintainer approval。

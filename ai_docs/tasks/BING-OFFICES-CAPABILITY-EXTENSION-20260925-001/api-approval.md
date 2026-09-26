# Public API / Provider Package Approval

- `approvedBy`: user (implementation-plan approval)
- `approvedAt`: `2026-09-25`
- `taskId`: `BING-OFFICES-CAPABILITY-EXTENSION-20260925-001`
- `scope`: additive dual-TFM API and a new optional read-only Provider package

## Approved Additions

- `ExcelFormat.Xlsb` is appended after the existing `Xls` and `Xlsx` values; existing numeric values remain unchanged.
- `IExcelBatchImporter`, `ExcelBatchImportRequest<TItem>`, `ExcelImportBatch<TItem>` and `ExcelBatchImportSummary` provide callback-based single-Sheet batch import without adding an `IAsyncEnumerable` dependency.
- `IExcelProviderCapabilityDescriptor` is an optional fine-grained capability description. Existing `IExcelProviderCapabilities` implementations remain valid.
- The `Bing.Offices.ExcelDataReader` package targets `net6.0` and `net8.0`, exposes a read-only `ExcelDataReaderExcelImporter`, and does not register or simulate an exporter.
- `ExcelDataReaderServiceCollectionExtensions.AddBingOfficesExcelDataReader(IServiceCollection)` registers only the importer and batch importer services.

## Compatibility

- No existing public member is removed or changes signature.
- Existing `ExcelFormat.Xls` and `ExcelFormat.Xlsx` numeric values are preserved.
- The new package is optional and does not add a dependency from Core or the existing providers.
- XLSB is explicitly unsupported by providers that do not declare it; no automatic provider fallback is introduced.

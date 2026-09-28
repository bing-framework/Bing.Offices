# Public API approval: workbook export strategy

- approvedBy: user (continued implementation of the prioritized TODO list)
- approvedAt: 2026-09-28 (Asia/Shanghai)
- taskId: BING-OFFICES-WORKBOOK-EXPORT-STRATEGY-20260928-001
- Scope: additive report-level selection over existing public export interfaces.

## Approved additions

- `Bing.Offices.Exports.ExcelWorkbookExportMode`: `CompleteWorkbook`, `ForwardStreaming`.
- `Bing.Offices.Exports.ExcelWorkbookExportStrategy`: constructor accepting optional `IExcelExporter` and `IExcelStreamingExporter`; stream and file `Export`/`ExportAsync` operations with explicit mode, optional streaming options and cancellation token.

Both types are User API. Existing exporters, interfaces and signatures remain unchanged. `CompleteWorkbook` describes the public complete-workbook contract, not a physical DOM guarantee: NPOI may use SXSSF internally and MiniExcel uses its own forward writer. `ForwardStreaming` never falls back to a complete exporter and rejects templates before output. Consumers migrate only when they need explicit per-report selection.

The net6.0 and net8.0 snapshot must show only these additions and zero removals before baseline approval is recorded.

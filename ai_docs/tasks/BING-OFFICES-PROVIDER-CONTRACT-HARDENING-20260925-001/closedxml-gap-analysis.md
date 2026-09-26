# ClosedXML Gap Analysis

## Failure Workbook

Execution evidence: `ClosedXmlExcelImporter` now routes supported failure modes through the internal writer. ClosedXML unit coverage is `84/84` on both TFMs, including annotation, preservation, comment conflict, ErrorRowsOnly, resource caps, atomic destination, async ownership, and the approved Entity resource option path. Rich OOXML parts remain an explicit preflight boundary.

Target: reuse `ExcelImportFailureOptions`; annotate the correct row/column; preserve unaffected workbook content; honor comment conflict policy, stream ownership, cancellation, maximum serialized/copy resources, and atomic destination behavior. Unsupported chart/image/macro/pivot preservation remains preflight fail-fast.

Implementation order: common scenario/expected snapshot -> confirm NPOI/Mini expected profiles -> ClosedXML internal writer -> ClosedXML unit tests -> real XLSX/file integration. No ClosedXML public option.

## Workbook Data Validation import

Execution evidence: `WorkbookRules` and `ConfiguredAndWorkbook` now enter the ClosedXML workbook-validation pipeline. Frozen Golden inputs cover list, whole-number, decimal, date, time, text-length, and custom rules. `WorkbookValidationContractTest` passes 16/16 direct NPOI/ClosedXML cases per TFM, including explicit unsupported Report/Fail policy; configured validation remains supported.

Target: parse ClosedXML worksheet validations into provider-neutral executable rules. Support is gated per validation type and operator; unsupported formulas/references must fail or follow the existing explicit policy, never disappear. The detailed planning matrix is in `validation-support-matrix.md`.

Risk: list ranges, relative formulas, named ranges, cross-sheet references, locale-specific formulas, full operator/blank combinations, and custom formulas may not map to the public validation model. A `PARTIAL` outcome is valid if each unsupported cell/rule has deterministic behavior.

Golden governance evidence: the offline generator freezes 21 manifest entries. Formula/date/template XML content passes `GoldenFixtureContentContractTest` 7/7 per TFM, the provider-specific formula profile passes `FormulaGoldenContractTest` 12/12 per TFM, and `GoldenFixtureManifestTest` verifies hashes and exact resource membership 1/1 per TFM.

## Entity List Region Dynamic Columns

Current evidence: `ClosedXmlEntityLayoutExecutor` reuses the shared physical-column planner for entity fixed + dynamic columns; the former dynamic-column rejection path is removed.

Execution evidence: `ClosedXmlEntityLayoutExecutor` supports fixed + dynamic columns, `Order`, `PlacementKey`, and `ColumnIndex`; import creates a case-insensitive dictionary when needed. The direct round-trip test plus converter/raw-and-converted validation test and Entity-filtered tests pass on both TFMs. Collision/overflow and relation-specific dynamic cases remain provider-specific follow-up coverage.

Required tests: shared entity dynamic scenario, collision/overflow/unknown-key failure, ClosedXML planner/executor unit tests, and real XLSX round trip. NPOI current behavior must be established before asserting cross-provider equivalence; MiniExcel is `NOT_APPLICABLE` while Entity is unsupported.

## Entity Resource Options

The user approved the provider-neutral option design. `ExcelEntityImportOptions` carries `ExcelResourceLimits`, and the opt-in `IExcelEntityResourceImporter` adds sync/async ordinary and template imports without changing the legacy `IExcelEntityImporter`. NPOI and ClosedXML enforce the limits; direct MaxInputBytes tests, legacy-provider compatibility tests, dual-TFM API comparison, and package consumers pass. This gap is `CLOSED`.

## Explicit exclusions

Chart creation, PivotTable, XLSM macro support, full formula engine, and smart routing are out of scope. Image/Table remain Spike-only without a mature public contract and golden XLSX evidence.

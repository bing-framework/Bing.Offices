# Resource Report

## Evidence

- Follow-up adds exact-boundary Workbook `MaxRows` success for all three providers in both entry modes, while single-Sheet overflow now explicitly asserts an empty root. `ResourceLimitContractTest`: 27/27 per TFM.
- Failure Workbook file/stream boundary matrix: 48/48 per TFM, including NPOI HSSF/XSSF and ClosedXML XSSF; both output modes and both entry modes. Core real-file cancellation-after-write covers existing and new targets, 12/12 default committer tests per TFM. This strengthens failure-path evidence without rerunning unchanged capacity workloads.

- ClosedXML unit coverage includes input bytes, rows, sheets/columns/cells, ZIP/XML admission, pictures, tracked unique values, error count, Failure Workbook serialized bytes, cancellation, temporary cleanup, atomic destination checks, dynamic collision/overflow, relation binding, and the approved Entity resource option path; the current suite passed `101/101` on both TFMs.
- NPOI direct Entity provider coverage includes the approved `IExcelEntityResourceImporter` MaxInputBytes gate; the current suite passed `571/571` on both TFMs.
- Shared `ResourceLimitContractTest.MaxInputBytes_ShouldRejectBeforeMaterialization` passed for NPOI, MiniExcel, and ClosedXML, `3/3` per TFM.
- The post-gap shared resource matrix passed `15/15` per provider-contract TFM, covering `MaxInputBytes`, `MaxRows`, `MaxSheets`, `MaxColumnsPerSheet`, `MaxCells`, `MaxErrors`, and `MaxTrackedUniqueValues` with structured errors and bounded materialization.
- MiniExcel now enforces `MaxRows` during row enumeration and maps `UniqueTracker` overflow to `ExcelImportErrorCode.ResourceLimit` instead of `ValueConversion`; the full MiniExcel unit suite remained `43/43` per TFM.
- Existing `artifacts/resource-probe/resource-matrix.jsonl` and related resource artifacts are reused because the benchmark harness and workload identity were unchanged. No new L5 probe was required by the change-impact rules.
- Golden resource governance is independently verified by a 21-entry SHA-256 manifest; malformed ZIP and oversized-entry inputs are frozen for future preflight/resource contracts.

## Gate Classification

The resource implementation evidence is `PASS` for the covered limits. Large-capacity and production-machine evidence remains an external/release gate and is not fabricated as a fresh measurement.

## Provider Contract 后续缺口

- MiniExcel `MaxRows` is now a workbook-shared row budget. The direct two-sheet contract verifies that aggregate overflow produces `ResourceLimit` and returns an empty root rather than previously materialized entities.
- NPOI and ClosedXML count `MaxCandidateErrorRows` only for `ErrorRowsOnly`, deduplicate by `(SheetName, RowIndex)`, and leave `AnnotatedOriginal` unrestricted by that candidate-row budget. NPOI direct preflight evidence is `15/15` per TFM; ClosedXML direct preflight evidence is `3/3` per TFM, including the same-row/different-Sheet distinction and the AnnotatedOriginal exemption.
- No L5 matrix rerun was needed: this change modifies admission semantics and output atomicity, not the resource harness, workload, or capacity threshold.

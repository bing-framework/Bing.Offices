# Provider Contract Report

## Current Candidate

ProviderContract passed 192/192 on each TFM in the final solution test (exit 0). The follow-up adds 48 failure-workbook file/stream cases and six exact MaxRows boundary cases per TFM to the prior 138-case checkpoint. See `gap-followup-plan.md` and `production-symbol-test-map.md` for scope and methods. Complete raw results are under `artifacts/provider-followup-final/trx/`; `solution-test.log` identifies the project/TFM for each result file.

## Historical Result

`Bing.Offices.ProviderContract.Tests` passed `112/112` on `net6.0` and `112/112` on `net8.0` in the post-gap checkpoint.

Covered scenarios:

- Scalar, dynamic columns, mapping, named converter, validation, and relations across NPOI, MiniExcel, and ClosedXML.
- Independent expected snapshots and independent capability profiles.
- Workbook native list/numeric/decimal/date/time/text-length validation plus unsupported custom-rule `Report`/`Fail` behavior (16 direct provider cases per TFM).
- Frozen Golden manifest/hash governance (1/1 per TFM) and independent formula/date/template content checks (7/7 per TFM).
- Formula Golden profile matrix: 12/12 per TFM; ClosedXML preserves formula text, while NPOI/MiniExcel cached-value behavior is explicitly `PARTIAL`.
- Rich layout profile matrix: 12/12 per TFM across style, merge, template, and entity scenarios; MiniExcel unsupported outcomes are asserted as structured preflight failures and entity is explicitly `NOT_APPLICABLE`.
- Sync/async equivalence, caller stream ownership, and input-byte resource rejection.
- Shared resource-limit matrix for rows, sheets, columns, cells, errors, and tracked unique values across all three providers.
- Real-file Failure Workbook output and pre-cancelled destination-preservation cases across all three providers.
- Runner diagnostics capture for scenario, format, expected profile, declared capabilities, and actual outcome.

The expected profiles are test-owned. Provider capability declarations are checked for consistency but do not decide whether a case executes. The shared suite replaces the removed common methods in both legacy integration contract classes; native DI, resource, and public-surface boundaries remain in their original project.

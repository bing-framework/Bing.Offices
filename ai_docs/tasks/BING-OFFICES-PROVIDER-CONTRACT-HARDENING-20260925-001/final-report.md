# Final Report

## Current Candidate — follow-up completed

The analysis in `gap-followup-plan.md` identified four remaining verification/delivery gaps. All four are implemented without changing production code, public APIs, dependency versions or resource thresholds:

1. Added 48 file/stream failure-boundary cases per TFM: NPOI XLS/XLSX and ClosedXML XLSX, synchronous/asynchronous entry, AnnotatedOriginal/ErrorRowsOnly, new/replaced file, commit/serialization failure, partial stream write and caller ownership.
2. Added six cross-Sheet exact-MaxRows boundary cases per TFM and strengthened overflow to require an empty root.
3. Added three Core cases per TFM for cancellation after an actual temporary-file write, with new/existing targets preserved correctly. Provider copy-phase cancellation is evidenced at the shared committer rather than falsely attributed to input cancellation tests.
4. Re-ran current-package consumers with an isolated cache and collected a complete final solution-test log, exit code and TRX set.

Final Release build: exit 0, zero warnings/errors. Final Release solution test: exit 0, **2378 passed / 0 failed / 0 skipped across 19 TRX runs**. Command: `dotnet test Bing.Offices.sln --no-restore -c Release -v:minimal -m:1 --logger trx --results-directory artifacts/provider-followup-final/trx`. Evidence: `artifacts/provider-followup-final/solution-test.log`, `solution-test.exitcode` and `trx/`.

| Project | net6.0 | net8.0 |
| --- | ---: | ---: |
| Core unit | 204/204 | 204/204 |
| NPOI unit | 571/571 | 571/571 |
| MiniExcel unit | 43/43 | 43/43 |
| ClosedXML unit | 101/101 | 101/101 |
| ProviderContract | 192/192 | 192/192 |
| Shared integration | 30/30 | 30/30 |
| NPOI integration | 30/30 | 30/30 |
| MiniExcel integration | 9/9 | 9/9 |
| ClosedXML integration | 4/4 | 4/4 |
| Docs | N/A | 10/10 |

API snapshot comparison passed for both TFMs. Four consumer build/run combinations passed from current `artifacts/api-candidate` packages; all five product packages in the isolated cache match that feed byte-for-byte. See `package-consumer-report.md`. Independent code review found no actionable issue; final evidence review is recorded in `review.md`. External CI/other OS/font/production-capacity evidence remains BLOCKED_EXTERNAL.

Below is historical implementation context; its old counts do not describe the current candidate.

## Implementation

The provider-neutral Testing/ProviderContract graph is in place and drives all three providers. ClosedXML now has contract-first Failure Workbook output, an explicit native workbook-validation adapter, and Entity ListRegion dynamic-column support reusing the existing physical-layout planner. The user-approved additive Entity resource API is implemented through `ExcelEntityImportOptions` and the opt-in `IExcelEntityResourceImporter`; the legacy `IExcelEntityImporter` remains unchanged.

## Historical Verification (superseded by current candidate evidence)

- ProviderContract historical candidate: `77/77` per TFM. Post-gap targeted/full contract checkpoint: `112/112` per TFM, adding the shared resource matrix, real-file Failure Workbook boundaries, snapshot comparer edges, and runner diagnostics.
- Provider units: NPOI `568/568`, MiniExcel `43/43`, ClosedXML `84/84` per TFM; the additional direct tests cover the approved entity resource option path.
- Provider integrations: NPOI `30/30`, MiniExcel `9/9`, ClosedXML `4/4` per TFM; shared `Bing.Offices.Tests.Integration` passed `30/30` per TFM.
- Core tests: `201/201` per TFM.
- Consumers and ThirdParty public-only checks passed against the fresh approved package feed at identity `2.0.0`; the five consumer-cache nupkg hashes match the approved feed.
- Public API snapshot compare and `PublicApiContractTest` passed for both `net6.0` and `net8.0`; the baseline now includes the approved ClosedXML assembly/package and Entity resource API members.
- Full solution Release build passed with 0 warnings and 0 errors after excluding the independent GoldenFixtures project tree from the parent BuildScript source glob.
- Golden governance: 21 frozen manifest entries; manifest `1/1`, formula/date/template content `7/7`, workbook validation `16/16`, and formula profile `12/12` per TFM.

The post-gap full solution test recheck passed with exit code 0 on both target frameworks: Core `201/201`, NPOI `568/568`, MiniExcel `43/43`, ClosedXML `86/86`, NPOI/MiniExcel/ClosedXML integrations `30/9/4`, shared integration `30/30`, and ProviderContract `112/112`.

## Gates

Workbook validation's matrix now has explicit `SUPPORTED` or `UNSUPPORTED_EXPECTED` entries; the historical `PARTIAL` label is obsolete. External CI/OS/font/production-capacity evidence remains unavailable in this environment. Golden XLSX governance is frozen and directly verified. External gates are separate from implementation/test status.

Phase 10 rich baseline and Phase 16 method-level common-contract migration are complete. Public API and Entity resource approvals are closed with member-level evidence. The remaining release status is external-only: CI/OS/font/capacity evidence. Phase 15 Golden governance remains complete with frozen resources and dual-TFM self-tests.

## Post-Gap Implementation Checkpoint

- MiniExcel resource-limit behavior is now aligned with the shared contract for row budgets and unique-value budgets; no public API or baseline change was made.
- Shared resource, Failure Workbook, cancellation, exception, snapshot, entity-dynamic, and real-file tests passed on both target frameworks.
- Golden manifest entries require and retain a non-empty `UpdateReason`; the existing 21-entry manifest satisfies the strengthened governance check.

## Provider Contract 后续缺口

All eight planned groups are closed locally. `ExcelImportFailureOptions.DestinationPath` is an approved additive public API, file output uses the existing Core atomic committer, and caller-owned stream output remains explicitly best-effort. MiniExcel now applies `MaxRows` across the workbook; NPOI/ClosedXML apply the ErrorRowsOnly candidate-row budget once per physical row; ClosedXML workbook validation has explicit support or structured Unsupported results for every inspected rule shape.

Previous review checkpoint: provider builds `0` warnings/errors; NPOI `571/571`, MiniExcel `43/43`, and ClosedXML `101/101` unit tests per TFM; ProviderContract `138/138` and shared integration `30/30` per TFM; API snapshot comparison passed for both TFMs. The later follow-up obtains a single complete solution-test log, exit code and TRX set instead of relying on the earlier partial console transcript. No commit, push, release, dependency upgrade, or version change was performed.

The final review-fix checkpoint adds ProviderContract `138/138` per TFM, NPOI candidate-row preflight `15/15`, and ClosedXML candidate-row preflight `3/3`. It verifies real-I/O mid-read cancellation for all three providers in the shared cancellation contract and preserves the prior target with no commit temporary file in the NPOI/ClosedXML file-target contract.

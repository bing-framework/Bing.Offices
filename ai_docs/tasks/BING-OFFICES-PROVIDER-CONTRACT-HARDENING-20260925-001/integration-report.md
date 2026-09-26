# Integration Report

## Current Candidate Gate

The follow-up final Release solution test exited 0: 2378 passed / 0 failed / 0 skipped in 19 TRX runs. Integration counts below remain current; ProviderContract is now 192/192 per TFM. Evidence is `artifacts/provider-followup-final/solution-test.log`, `solution-test.exitcode` and `trx/`. Package consumer verification now uses `artifacts/api-candidate` with an isolated matching cache (four successful build/run combinations). Older checkpoint sections below are historical only.

## Provider Integration Suites

| Project | net6.0 | net8.0 |
| --- | ---: | ---: |
| `Bing.Offices.Npoi.Tests.Integration` | 30/30 | 30/30 |
| `Bing.Offices.MiniExcel.Tests.Integration` | 9/9 | 9/9 |
| `Bing.Offices.ClosedXml.Tests.Integration` | 4/4 | 4/4 |
| `Bing.Offices.Tests.Integration` | 30/30 | 30/30 |

The `Bing.Offices.Tests.Integration` result includes the retained MiniExcel/ClosedXML native boundary methods, the options-aware Entity extension compatibility test, and the approved Public API snapshot test. The dual-TFM API baseline now includes the approved Abstractions additions and ClosedXML assembly/package identity.

## Provider Contract Checkpoint

The historical ProviderContract checkpoint passed `77/77`; the post-gap checkpoint passed `112/112` on both target frameworks. The result includes the new Style, Merge, Template, and Entity profile-driven cases, all Golden governance/content cases, workbook validation cases, formula profile cases, the shared resource matrix, and real-file Failure Workbook boundaries.

## Historical Candidate Gate

The historical solution checkpoint recorded Core 201/201 per TFM; NPOI Unit 568/568; MiniExcel Unit 43/43; ClosedXML Unit 84/84; NPOI Integration 30/30; MiniExcel Integration 9/9; ClosedXML Integration 4/4; ProviderContract 77/77; shared Integration 30/30. It is not the follow-up final candidate evidence.

## Post-Gap Targeted Integration Evidence

- `FailureWorkbookFileContractTest` passed `6/6` per TFM, covering real XLSX file input, complete failure output, unsupported MiniExcel preflight, cancellation, destination preservation, and caller-owned input streams.
- The MiniExcel provider integration filter passed `4/4` per TFM after the resource-limit implementation change.

## Post-Gap Full Solution Recheck

The complete solution test passed with exit code 0 on both target frameworks: NPOI integration `30/30`, MiniExcel integration `9/9`, ClosedXML integration `4/4`, shared integration `30/30`, and ProviderContract `112/112`.

## Provider Contract 后续缺口最终结果

- `DestinationPath` is covered as a public additive API, with exactly-one-target validation and atomic replacement for NPOI/ClosedXML. MiniExcel rejects both output target forms before writing any bytes.
- The current shared integration suite passed `30/30` on each TFM after the API snapshot ordering and Core IVT approval list were updated.
- Previous review ProviderContract checkpoint passed `138/138` on each TFM, including all three providers in the real-I/O mid-read cancellation contract and the shared assertion-helper diagnostics self-test.
- The final solution Release test completed successfully after these additions. The prior integration/consumer evidence is reused only where its production source scope is unchanged.

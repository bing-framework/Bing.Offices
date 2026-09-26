# Unit Test Report

## Current Candidate

Final solution test exit 0, 19 TRX files and 2378 passing cases, no failures/skips. Current per-TFM unit counts: Core 204/204; NPOI 571/571; MiniExcel 43/43; ClosedXML 101/101; ProviderContract 192/192. Docs net8.0 10/10. Raw evidence is `artifacts/provider-followup-final/{solution-test.log,solution-test.exitcode,trx/}`; the final-report table includes all integration counts.

## Historical Result

All affected provider unit suites passed on both target frameworks:

| Project | net6.0 | net8.0 |
| --- | ---: | ---: |
| `Bing.Offices.Tests` | 201/201 | 201/201 |
| `Bing.Offices.Npoi.Tests` | 568/568 | 568/568 |
| `Bing.Offices.MiniExcel.Tests` | 43/43 | 43/43 |
| `Bing.Offices.ClosedXml.Tests` | 86/86 | 86/86 |

Commands used the project-level `dotnet test -c Release -f <TFM> --no-restore` form. ClosedXML's 86-case result includes Failure Workbook, workbook validation, Entity Dynamic Columns, the new collision/overflow and relation cases, and the approved Entity resource option coverage; NPOI's 568-case result includes its corresponding direct resource option test.

## Notes

- No provider unit regression was observed after the new ClosedXML paths. The final ClosedXML count includes the entity dynamic converter/raw-and-converted validation regression and the MaxInputBytes entity gate.
- The historical ProviderContract checkpoint passed `77/77`; the post-gap checkpoint passed `112/112` on each TFM and additionally covers the shared resource matrix, real-file Failure Workbook boundaries, snapshot comparer edges, and runner diagnostics.
- The SDK emitted the expected `NETSDK1138` warning only when package consumers targeted net6.0; provider unit projects completed without test failures.

## Post-Gap Full Solution Recheck

The complete solution test passed with exit code 0 on both target frameworks. The current counts are Core `201/201`, NPOI `568/568`, MiniExcel `43/43`, ClosedXML `86/86`, and ProviderContract `112/112`.

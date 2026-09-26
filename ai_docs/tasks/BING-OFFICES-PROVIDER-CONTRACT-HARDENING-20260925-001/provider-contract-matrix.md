# Provider Expected Capability Matrix

Legend: `S` supported/success, `U` expected structured unsupported, `P` partial policy requires matrix, `N/A` not applicable. Expected values are test-owned and must not be derived from production capability flags.

| Contract feature | NPOI XLSX | NPOI XLS | MiniExcel XLSX | ClosedXML XLSX | Declared capability check | Required actual outcome |
| --- | --- | --- | --- | --- | --- | --- |
| List / Workbook | S | S | S | S | List, Workbook, format | expected snapshot |
| Sync / Async outer IO | S | S | S | S | Async | equivalent result, real cancellation/ownership |
| Mapping / ValueMap / Converter | S | S | S | S | base format capability | expected snapshot |
| Dynamic columns | S | S | S | S | base workbook capability | expected physical/result snapshot |
| Configured validation | S | S | S | S | base workbook capability | stable error snapshot |
| Workbook native validation | P | P | U | P | no fine-grained public bit | explicit rule outcome, never silent ignore |
| Relations | S | S | S | S | Workbook/Entity as applicable | relation snapshot |
| Entity layout | S | S | U | S | Entity | success or fail-fast unsupported |
| Entity dynamic columns | S/P to confirm | S/P to confirm | N/A | S | Entity + test profile | shared placement/mapping snapshot |
| Template | S | S | U | S with preflight limits | Template | preservation snapshot or structured unsupported |
| Merge | S | S | U | S/basic | Merge | public merged-range result |
| Row height | S | S | U | S | no fine-grained bit | success or no-output unsupported |
| Formula text/cache | P | P | P | P | no formula bit | per-cell-type profile |
| Failure Workbook | S | S | U | S (documented preservation limits) | existing import option | output/error/atomicity snapshot |
| Public resource limits | S | S | S | S | base import capability | same error category/no partial data |
| ZIP/XML protection | S | N/A | provider preflight policy | S | provider-specific | provider unit tests |
| XLS format | S | S | U | U | Xls | structured unsupported and no output |

## Three-layer assertion

Follow-up file/stream boundary evidence: `FailureWorkbookFailureBoundaryContractTest` covers NPOI XLS/XLSX and ClosedXML XLSX, both modes and both entry paths (48 cases per TFM). Exact aggregate MaxRows is covered for all three providers in both entry modes. Copy-phase cancellation is directly exercised at the shared Core committer for existing/new files; provider input mid-read cancellation is separate. Native validation's `P` means limited formula semantics with explicit Unsupported; the supported subset is complete in `validation-support-matrix.md`.

Each case records `Expected`, `Declared`, and `Actual`. A case passes only when declared capability is consistent with the coarse public bit and actual behavior matches the independent expected profile. Rich features without a public bit use a test-only `ContractFeature`; this task must not add product capability bits solely for testing.

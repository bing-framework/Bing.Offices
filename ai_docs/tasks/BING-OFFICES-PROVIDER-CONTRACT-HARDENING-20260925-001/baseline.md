# Baseline

## Candidate identity

- Task: `BING-OFFICES-PROVIDER-CONTRACT-HARDENING-20260925-001`
- Baseline commit: `159e6a9fd1f2bf2da9e679f3c9b6fb3a87edddc1`
- Branch: `feat/miniexcel-provider`
- SDK: `.NET SDK 10.0.401`
- Working tree at planning time: `.github/workflows/ci.yml` has a pre-existing user modification; this task must not overwrite it.
- Text contract: Markdown is UTF-8 without BOM + LF; C# is UTF-8 BOM + LF; `.csproj`/`.sln` are CRLF.

## Repository inventory

At planning time, the solution contained five production projects (`Abstractions`, `Core`, `Npoi`, `MiniExcel`, and the new `ClosedXml` provider), seven principal unit/integration test projects, API snapshot tooling, Docs tests, ProfileFixtures, ResourceProbe, package consumers, and the third-party provider consumer. `Bing.Offices.Testing` and `Bing.Offices.ProviderContract.Tests` did not yet exist.

The primary test-source scale observed during planning is:

| Project | C# files | Approx. lines | Responsibility today |
| --- | ---: | ---: | --- |
| `Bing.Offices.Tests` | 46 | 7,778 | Core/Abstractions and legacy tests |
| `Bing.Offices.Tests.Integration` | 12 | 2,851 | Public integration, API snapshot, two hand-written provider contract classes |
| `Bing.Offices.Npoi.Tests` | 31 | 15,882 | NPOI unit, provider contract, HSSF/XSSF, staging |
| `Bing.Offices.Npoi.Tests.Integration` | 10 | 1,479 | Real XLS/XLSX/file/failure workbook |
| `Bing.Offices.MiniExcel.Tests` | 10 | 2,248 | MiniExcel unit and adapter boundaries |
| `Bing.Offices.MiniExcel.Tests.Integration` | 9 | 574 | Real XLSX/file/cancellation |
| `Bing.Offices.ClosedXml.Tests` | 9 | 2,677 | ClosedXML unit and native DOM boundaries |
| `Bing.Offices.ClosedXml.Tests.Integration` | 9 | 239 | Real XLSX/file/cancellation |

`rg` found approximately 792 public test methods in the principal test projects. This is an inventory metric, not a success metric.

## Existing evidence to reuse

The immediately preceding ClosedXML task reports:

- ClosedXML Unit: `74 x 2 TFM` passing.
- ClosedXML Integration: `4 x 2 TFM` passing.
- Existing cross-provider checks: `10 x 2 TFM` passing.
- Full solution: 1,944 tests, 1,942 passing and two API snapshot approval failures.
- Reused pre-task review state: `OPEN_ACTIONABLE=0`; API snapshot was `BLOCKED_APPROVAL`, while external OS/font validation was `BLOCKED_EXTERNAL`.

Source: `ai_docs/tasks/BING-OFFICES-CLOSEDXML-PROVIDER-20260922-001/{final-report.md,review.md,test-matrix.md,symbol-test-map.md}`. These results may be reused while their relevant source scope remains unchanged. The execution phase must produce fresh evidence only at the phase checkpoints defined in `plan.md`.

## Current contract-test baseline

- Baseline `MiniExcelProviderContractTest` had 7 NPOI/MiniExcel contracts and duplicated private models/request setup/string snapshots; the current candidate retains only MiniExcel unsupported-preflight and DI registration boundaries after method-level migration.
- `ClosedXmlProviderContractTest` has 10 two/three-provider contracts and repeats the same categories of infrastructure.
- Several comparisons derive expected results from NPOI actual output. NPOI therefore acts as an implicit golden provider.
- Provider constructors are repeated throughout tests; there is no driver, independent expected capability profile, or common case source.
- Existing XLSX resources are legacy examples/bug fixtures under `Bing.Offices.Tests/Resources`; they have no provider-neutral manifest/hash governance.

## Known capability boundaries

- NPOI: XLS/XLSX, list/workbook/entity/template/merge/async; Failure Workbook implemented.
- MiniExcel: XLSX list/workbook/async; entity, template, merge, XLS and Failure Workbook are explicit unsupported paths.
- ClosedXML: XLSX list/workbook/entity/template/basic merge/async; Failure Workbook and workbook-native validation are preflight unsupported; Entity List Region dynamic columns are plan-stage unsupported.
- Entity APIs expose no resource options input, while `ExcelEntityImportResult` already exposes `ErrorsTruncated`/`MaxErrors`. Adding resource input is a public Abstractions decision.

## Baseline policy

Do not regenerate or overwrite API baselines, benchmark baselines, or golden resources during ordinary migration. Baseline identity changes require an explicit reason in `decisions.md`. No package/version/dependency update belongs to this task.

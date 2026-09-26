# Member API Diff

## Snapshot Scope

- Task: `BING-OFFICES-ADVANCED-EXCEL-ROADMAP-20260926-001`
- Previous baseline: `artifacts/open-provider-20260926/api-snapshot-baseline.before-advanced-20260926.json`
- Current capture: `artifacts/open-provider-20260926/api-final/capture`
- Final comparison: `artifacts/open-provider-20260926/api-final/compare`
- Capture assemblies: Abstractions, Core, NPOI, MiniExcel, ClosedXML, ExcelDataReader, SpreadCheetah and AsposeCells.
- Target frameworks: `net6.0` and `net8.0`.

The final capture was produced from the frozen Release output and the final eight-package feed. Its candidate identity is recorded in `api-final/capture/api-candidate.json`; the tracked baseline contains the same identity metadata. No public API snapshot shape or approved member list was changed by the identity refresh.

## Result

| TFM | Assembly | Before | After | Added snapshot lines | Removed |
|---|---|---:|---:|---:|---:|
| `net6.0` | `Bing.Offices.Abstractions` | 1128 | 1163 | 38 | 0 |
| `net8.0` | `Bing.Offices.Abstractions` | 1128 | 1163 | 38 | 0 |
| `net6.0` | `Bing.Offices.Core` | unchanged | unchanged | 0 | 0 |
| `net8.0` | `Bing.Offices.Core` | unchanged | unchanged | 0 | 0 |
| `net6.0/net8.0` | NPOI, MiniExcel, ClosedXML, ExcelDataReader, SpreadCheetah, AsposeCells | unchanged | unchanged | 0 | 0 |

`Removed=0` for every captured TFM and assembly. No existing public member or enum value was changed.
The snapshot adds 38 canonical lines: 35 countable public members plus 3 public type declarations. The API tool's `memberCount` excludes `type|` lines, which is why the count changes from 1128 to 1163 rather than 1166.

## Final Candidate Identity

The tracked `build/api-snapshot-baseline.json` received an identity-only refresh after the final candidate was frozen. The `assemblies` member lines, public API hashes, and approved 38-line additive boundary are unchanged. The following are canonical candidate identity hashes, not raw file hashes:

| Identity field | Previous candidate | Final candidate |
|---|---|---|
| Candidate source manifest | `FF0D2FBF...4A0700` | `2F0BDDFE...1CFA03` |
| Approval file | `B576AB8C...1EFE5` | `84AD079E...56138` |
| `netstandard2.0/Bing.Offices.Abstractions.dll` | `2DB52608...4B701` | `669FF413...29836` |
| `Bing.Offices.Abstractions.2.0.0.nupkg` | `99BD3754...261F0` | `C2FF0FF9...52B34` |
| `Bing.Offices.Core.2.0.0.nupkg` | `D257B482...7F3AA` | `4FB3E1AD...C5F9D` |
| `Bing.Offices.Npoi.2.0.0.nupkg` | `0677BF3C...E2DE9` | `282211EB...56E73` |
| `Bing.Offices.MiniExcel.2.0.0.nupkg` | `0FC15BED...97AB1` | `D2E67829...4C06` |
| `Bing.Offices.ClosedXml.2.0.0.nupkg` | `8D41A264...BEAEBE` | `CB9C5E67...249EED` |
| `Bing.Offices.ExcelDataReader.2.0.0.nupkg` | `04493B8D...A979` | `95389CEF...B4396` |
| `Bing.Offices.SpreadCheetah.2.0.0.nupkg` | `63734CA4...9442E` | `605A30BF...518B72` |
| `Bing.Offices.AsposeCells.2.0.0.nupkg` | `31F1E32E...68D0B` | `11510849...DC686A` |

The final comparison revalidated the refreshed identity and reported empty `net6.0` and `net8.0` API diffs. The historical `api-snapshot-baseline.before-advanced-20260926.json` remains unchanged.

## Added Members

The following 38 canonical snapshot lines are identical in both TFMs. Type declarations are called out separately because they are not included in `memberCount`.

### `ExcelDataValidationDefinition` (17 members)

- public type declaration
- public parameterless constructor
- `Date1`
- `Date2`
- `ErrorMessage`
- `ErrorTitle`
- `IgnoreBlanks`
- `InputMessage`
- `InputTitle`
- `Operator`
- `Range`
- `ShowErrorMessage`
- `Type`
- `Value1`
- `Value2`
- `Values`
- `Validate()`

### `ExcelDataValidationType` (7 members)

- public enum type
- `Integer`
- `Decimal`
- `Date`
- `TextLength`
- `List`
- underlying `value__` field

### `ExcelSheetImageDefinition` (10 members)

- public type declaration
- public parameterless constructor
- `Content`
- `Row`
- `Column`
- `Width`
- `Height`
- `OffsetX`
- `OffsetY`
- `Validate()`

### Existing request and builder types (4 members)

- `ExcelSheetExportRequest.Images`
- `ExcelSheetExportRequest.DataValidations`
- `ExcelSheetExportBuilder<T>.Image(ExcelSheetImageDefinition)`
- `ExcelSheetExportBuilder<T>.DataValidation(ExcelDataValidationDefinition)`

## Approval Boundary

The additions are limited to the image and native data-validation surface explicitly approved in `api-approval.md` by the user on `2026-09-26`. This file does not approve commercial Aspose behavior, package version changes, or any future public API.

## Reproduction

```text
dotnet build Bing.Offices.sln --no-restore -c Release -v:minimal -m:1
dotnet build build/ApiSnapshot/ApiSnapshot.csproj --no-restore -c Release -v:minimal -m:1
dotnet run --project build/ApiSnapshot/ApiSnapshot.csproj -c Release --no-build --no-restore -- --root output/release --packages artifacts/open-provider-20260926/packages-final --repository . --approval ai_docs/tasks/BING-OFFICES-ADVANCED-EXCEL-ROADMAP-20260926-001/api-approval.md --capture true --output artifacts/open-provider-20260926/api-final/capture --commit 159e6a9fd1f2bf2da9e679f3c9b6fb3a87edddc1
dotnet run --project build/ApiSnapshot/ApiSnapshot.csproj -c Release --no-build --no-restore -- --root output/release --packages artifacts/open-provider-20260926/packages-final --repository . --baseline build/api-snapshot-baseline.json --approval ai_docs/tasks/BING-OFFICES-ADVANCED-EXCEL-ROADMAP-20260926-001/api-approval.md --output artifacts/open-provider-20260926/api-final/compare --commit 159e6a9fd1f2bf2da9e679f3c9b6fb3a87edddc1
```

The member comparison uses the canonical `lines` arrays from the old baseline and the final capture; both TFM comparisons report 38 added lines and zero removals. Of those additions, 35 contribute to `memberCount` and 3 are type declarations. The final compare's `api-diff.json` is `{ "net6.0": {}, "net8.0": {} }`.

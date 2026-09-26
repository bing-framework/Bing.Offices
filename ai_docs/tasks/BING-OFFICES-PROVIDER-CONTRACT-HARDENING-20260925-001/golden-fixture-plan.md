# Golden XLSX Fixture Plan

## Baseline

The repository has 18 source-controlled XLSX files under `Bing.Offices.Tests/Resources` (legacy Bugs, Purchase, national-code import templates, and `CarImport.xlsx`). They have no provider-neutral manifest/hash governance. Do not move or reinterpret them without auditing their existing consumers.

Current cross-provider formula/date fixtures are generated at runtime with NPOI. ClosedXML also generates many provider-specific templates with `XLWorkbook`; those native unit fixtures should remain local. Only provider-neutral inputs belong in the new golden set.

## Frozen source-controlled set

| Area | Initial files | Contract evidence |
| --- | --- | --- |
| Formula | `cached-number.xlsx`, `cached-bool.xlsx`, `cached-string.xlsx`, `no-cache.xlsx` | formula text and cached result policy |
| Date | `date1900.xlsx`, `date1904.xlsx` | serial/date target normalization |
| Validation | `list-invalid.xlsx`, `whole-between-valid.xlsx`, `whole-less-than-invalid.xlsx`, `decimal-between-valid.xlsx`, `date-between-valid.xlsx`, `time-between-valid.xlsx`, `text-length-invalid.xlsx`, `custom-unsupported.xlsx` | validation rule/operator Spike and ClosedXML direct matrix |
| Relations | `parent-child.xlsx` | deterministic relation binding |
| Template | `styles.xlsx`, `comments.xlsx`, `conditional-formatting.xlsx`, `merged-cells.xlsx` | basic preservation boundaries |
| Security | `malformed.zip`, `oversized-entry.xlsx` | controlled resource/preflight rejection |

The offline `build/GoldenFixtures` generator currently freezes 21 manifest entries. The exact
set is enforced by `GoldenFixtureManifestTest.GoldenManifest_ShouldMatchFrozenResources`, which
passed 1/1 on both target frameworks. Formula/date/template XML content is independently checked
by `GoldenFixtureContentContractTest` (7/7 per TFM), and validation inputs are consumed by the
NPOI/ClosedXML matrix (16/16 per TFM).

## Manifest contract

Each resource entry records path, purpose, expected sheets/cells/formulas/cached values/features, provenance/generator, SHA-256, creation tool/version, and contract version. A manifest self-test recalculates hashes from source-controlled files while excluding `bin/obj` copies.

Fixture updates require a documented reason, contract/snapshot impact, regenerated hash, and review. Tests consume frozen files and never invoke NPOI/MiniExcel/ClosedXML to generate their expected input at runtime. The generator is an offline net8 tool and is not part of normal test execution.

## Generator boundary

A deterministic offline fixture generator may be added under test tooling when manual authoring is impractical. Its output is reviewed and committed; the generator is not run as part of normal contract tests. If a library generates a file, a different inspection path plus manifest assertions must validate it before adoption.

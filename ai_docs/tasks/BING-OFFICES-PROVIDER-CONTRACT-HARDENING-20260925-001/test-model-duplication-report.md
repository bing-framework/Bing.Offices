# Test Model And Fixture Duplication Report

This is a semantic audit, not a duplicate-line count. Paths and line numbers reflect the planning baseline.

| Cluster | Locations / examples | Semantic difference | Action |
| --- | --- | --- | --- |
| Contract workbook/row | Removed from `MiniExcelProviderContractTest` and `ClosedXmlProviderContractTest`; shared models under `tests/Bing.Offices.Testing/Models` | Legacy models represented common scalar/dynamic/validation/relation intent in two classes. | Replaced by explicit `ScalarContract*`, `MappingContract*`, `DynamicContract*`, `ValidationContract*`, `RelationContract*`, and `EntityContract*` models. |
| Required validation model | Removed from both legacy contract classes | Equivalent required/quantity semantics were duplicated. | Shared `ValidationContractWorkbook/Row` is used by ProviderContract. |
| Relation model | Removed from both legacy contract classes; provider unit variants remain | Contract pair was semantically equivalent; provider variants add internal concerns. | Shared relation model remains in `Bing.Offices.Testing`; native variants retained. |
| Round trip runner | Removed `MiniExcelProviderContractTest.RoundTripAsync` and `ClosedXmlProviderContractTest.RoundTripAsync` | Equivalent export/import orchestration. | `ProviderContractRunner.RoundTripAsync` is the single common runner. |
| Snapshot helpers | Removed `ToContractRows`, `ToRuleRows`, `ToErrorContract`, `ToRelationRows`, `Snapshot`, `DynamicSnapshot`, `FormulaSnapshot`, and `RelationSnapshot` from legacy classes | Pipe-delimited strings and implicit culture rules were not independent typed contracts. | Typed `ContractSnapshots` and comparers are the common implementation. |
| Provider construction | Baseline contract classes instantiated providers repeatedly. | No semantic need for repetition. | Current common contracts use `ProviderDrivers`; retained integration classes instantiate only their native boundary providers. |
| Date/range fixtures | Mini provider test `:1281-1436,2164-2301`; NPOI boundary regression `:165-238,333-426` | Cross-provider date serial expectations are shared; raw reader/preflight assertions are native. | Freeze cross-provider XLSX/expected data; retain native parser assertions. |
| Stream doubles | NPOI `AsyncPipelineTest`, `TemplateAsyncBoundaryTest`, `StreamPipelineTest`; Core CSV tests; ClosedXML and MiniExcel provider tests | Most differ only by seekability/tracking/blocking/throw point. | Share behavior-based streams; retain filesystem/provider staging doubles. |
| Recording observer | At least five copies across Core/NPOI/ClosedXML. | Basic exception recording is identical; observer-failure variants are specialized. | Share recording observer; retain throwing/specialized variants. |
| Committer/temp directory | ClosedXML and NPOI throwing committers; NPOI/CSV temp directory helpers | Common failure injection and cleanup. | Share parameterized committer and disposable temp directory fixture where no provider internals leak. |
| Workbook builders | Multiple NPOI request/regression helpers and integration helpers | Simple workbook/request setup is common; HSSF/XSSF/rich features are NPOI-specific. | Extract only public-contract request factories; retain rich native builders. |
| Generic small models | `DynamicRow`, `DateRow`, `CsvRow`, `UniqueWorkbook/Row`, RowHeight `Row` | Some are same shape but different contract intent. | Rename/share only verified semantic matches; Docs/consumer/benchmark models stay local. |

## Provider-specific retention

- NPOI: HSSF/XSSF success and failure branches, extensions, style/font/picture/chart/rich text, failure-workbook internals, ZIP preflight, streaming/staging.
- MiniExcel: lazy/streaming behavior, raw date reader, parser preflight, provider adapter cancellation, native-not-started unsupported assertions.
- ClosedXML: `XLWorkbook` lifecycle/admission, template preflight, style adapter/DOM, formula policy, native template preservation and unsupported parts.
- Integration: real path locks, atomic replace, temporary-file cleanup, and large-file execution.
- ProfileFixtures, Docs, package consumers, ThirdPartyProvider.Consumer, benchmark and ResourceProbe models remain independent.

## Before/after metrics

The following counts are semantic-family counts, not duplicate-line counts. They use the baseline inventory above and the current shared-contract source tree; provider-native/rich-feature helpers remain intentionally separate.

| Metric | Before Phase 2 | Current candidate | Evidence / interpretation |
| --- | ---: | ---: | --- |
| Audited semantic model clusters | 12 | 12 | The 12 clusters in this report remain tracked; five common contract families now live under `Bing.Offices.Testing`, while rich/provider-native families remain local. |
| Duplicate common round-trip helper families | 2 | 1 | Legacy MiniExcel/ClosedXML helpers were deleted; `ProviderContractRunner.RoundTripAsync` is the single common family. |
| Duplicate common snapshot helper families | 2 | 1 | Legacy provider string helpers were deleted; typed `ContractSnapshots` is the common family. |
| Common stream fixture families | 4 | 1 | Provider-specific stream doubles remain local; common contract cases use `Bing.Offices.Testing.Streams.ContractStreams`. |
| Repeated common request-setup clusters | 2 | 1 | Common scalar/mapping/dynamic/validation/relation setup is centralized in `ContractRequests`; rich/native setup remains intentionally local. |

These metrics close the Phase 16 common-contract migration while retaining provider-native and rich-feature evidence whose semantics differ. Phase 10/16 has no remaining local actionable duplication item.

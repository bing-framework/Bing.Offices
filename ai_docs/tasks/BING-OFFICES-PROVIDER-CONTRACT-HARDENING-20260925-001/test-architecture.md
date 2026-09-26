# Testing Architecture

## Target layers

```text
Product public contract
        |
        v
Bing.Offices.Testing
  models + deterministic data + requests + expected snapshots
  comparers + streams + fixtures + golden manifest reader
        |
        v
Bing.Offices.ProviderContract.Tests
  scenario runner + independent expected profile + provider drivers
        |
        +-------------+----------------+
        v             v                v
      NPOI        MiniExcel        ClosedXML
        |             |                |
        v             v                v
 provider unit tests and provider integration tests
```

## Project boundaries

### `Bing.Offices.Testing`

- Target `net6.0;net8.0`, `IsPackable=false`.
- Reference only `Bing.Offices.Abstractions` and `Bing.Offices.Core`.
- Do not import xUnit, NPOI, MiniExcel, ClosedXML, or provider implementation projects.
- Own provider-neutral models, deterministic data, request factories, typed expected snapshots, normalization/comparison primitives, stream fixtures, temp-directory fixtures, and frozen XLSX resource metadata.
- Do not contain provider selection branches.

### `Bing.Offices.ProviderContract.Tests`

- Target `net6.0;net8.0`, `IsPackable=false`, xUnit test project.
- Reference `Bing.Offices.Testing` plus the three provider projects.
- Own `IExcelProviderTestDriver`, provider drivers, `ExcelProviderCases`, `ExcelProviderContractProfile`, scenario runners, and contract assertions.
- Use public provider APIs only; receive no `InternalsVisibleTo`.

### Existing test projects

- `Bing.Offices.Tests`: Core/Abstractions/mapping/validation/converter tests.
- Provider unit projects: native adapters, internals, lifecycle, provider-specific unsupported boundaries.
- Provider integration projects: real file/stream/library, atomic commit, cancellation, cleanup.
- `ProfileFixtures`: remain an independent external assembly scanning fixture.
- Consumers and third-party provider consumer: remain independent and must not reference Testing.
- Benchmarks/ResourceProbe: keep performance-specific models where shared attributes/reflection would change measurements.

## Contract primitives

- Scenario: immutable input, request factory, typed expected result, side-effect expectations, applicable execution modes and formats.
- Driver: provider identity and public exporter/importer/service-provider creation only.
- Expected profile: repository-owned expectation independent of `IExcelProviderCapabilities`.
- Runner: executes `Sync` or `Async`, captures result/exception/stream state, and normalizes to a snapshot.
- Assertion: compares expected to actual, then separately checks expected capability vs declared capability vs runtime behavior.

Do not skip unsupported cases. Their expected result is a structured `BingOfficesUnsupportedFeatureException` plus no partial output, preserved target/source ownership, and no provider-native operation where observable.

## Initial directory shape

```text
tests/Bing.Offices.Testing/
  Models/{Scalar,Mapping,Dynamic,Validation,Relations,Formula,Entity,Resources}/
  Data/
  Requests/
  Snapshots/
  Comparers/
  Streams/
  Fixtures/
  Resources/{Date,Formula,Validation,Relations,Template,Security}/

tests/Bing.Offices.ProviderContract.Tests/
  Drivers/
  Profiles/
  Cases/
  Contracts/
  Assertions/
```

Avoid one-type-per-file fragmentation. Group tightly related contract records and factories by scenario until a file has multiple independent responsibilities.

## Determinism and normalization

- Fixed inputs cover null, empty, whitespace, blank/missing cells, Unicode/Chinese, numeric types, enum, nullable, `DateTime`, and `DateTimeOffset`.
- Culture cases explicitly include invariant, `zh-CN`, and `en-US`.
- Snapshots are typed records; no pipe-delimited string snapshots.
- Error snapshots contractually compare code, sheet, row, column, and column key. Message text is diagnostic unless a scenario declares a stable message fragment.
- No `Random.Shared`, business-data `Guid.NewGuid()`, or `DateTime.Now`. Fuzz tests use a fixed documented seed.

## Fourth-provider acceptance

A new Provider D must enter common contracts by adding only a public driver, an expected capability profile, and provider-specific tests. Adding copied scalar/mapping/validation/relation/dynamic/stream/resource contract classes fails the architecture gate.

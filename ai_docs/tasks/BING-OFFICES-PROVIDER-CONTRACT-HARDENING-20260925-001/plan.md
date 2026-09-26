# Bing.Offices Provider Contract Testing And ClosedXML Hardening Plan

## Task

- Task ID: `BING-OFFICES-PROVIDER-CONTRACT-HARDENING-20260925-001`
- Stage: `create-plan`
- Execution entry: `/execute-plan`
- Baseline: `159e6a9fd1f2bf2da9e679f3c9b6fb3a87edddc1`
- Target frameworks: `net6.0;net8.0`
- Risk: HIGH. The work changes the solution test graph, shared provider contracts, real XLSX fixtures, and three ClosedXML runtime paths.

## Objectives

1. Establish `Bing.Offices.Testing` as a provider-neutral, xUnit-free test infrastructure library.
2. Establish `Bing.Offices.ProviderContract.Tests` so NPOI, MiniExcel, and ClosedXML execute the same inputs, requests, expected snapshots, and side-effect contracts.
3. Replace implicit NPOI golden behavior with independently authored expected snapshots and independent expected capability profiles.
4. Migrate semantic duplication without weakening provider unit or real integration evidence.
5. Use the common contracts to drive ClosedXML Failure Workbook, Workbook Data Validation, and Entity Dynamic Columns.

## Non-goals and safety boundaries

- Do not implement Chart creation, PivotTable, XLSM macro support, a formula engine, smart routing, or dependency upgrades.
- Do not change package/assembly versions, commit, push, tag, publish, or update the approved API baseline.
- Do not give Testing/ProviderContract assemblies production internals access.
- Do not make consumers or `ThirdPartyProvider.Consumer` reference Testing.
- Do not delete a legacy test before method-level replacement evidence exists.
- `Entity Resource Options` is `BLOCKED_APPROVAL`; no provider-private substitute is allowed.
- Preserve the user's existing `.github/workflows/ci.yml` modification.

## Change impact analysis

```text
ChangedProjects: Bing.Offices.Testing, Bing.Offices.ProviderContract.Tests,
  existing provider unit/integration tests, solution; conditionally ClosedXml/Core/Abstractions
ChangedPublicContracts: expected no for Phases 0-12; possible Entity resource proposal is BLOCKED_APPROVAL
ChangedRuntimePaths: ClosedXML import failure output, workbook validation, entity list-region layout
ChangedProviders: all three in contract tests; ClosedXML production in Phases 11-13
ChangedTFMs: net6.0 and net8.0 test graph; product TFMs unchanged
ChangedBuildPackaging: solution graph only; packages unchanged
ChangedBenchmarkHarness: no, unless a later reviewed test fixture reuse is proven measurement-neutral
ChangedDocsOnly: false
AffectedDependents: provider tests, integration tests, consumers/API snapshot at checkpoints
RiskLevel: HIGH
```

## Phase plan

### Phase 0 - Freeze baseline and duplication audit

- Goal: freeze candidate identity, test inventory, inherited evidence, duplicate clusters, and existing user changes.
- Files: task `baseline.md`, duplication/coverage/migration matrices; read-only scan of `tests/**` and prior ClosedXML task.
- Why: establishes evidence keys and prevents mechanical merging of same-named, different-semantic models.
- Dependencies: none.
- Excludes: code changes, restore, broad tests, baseline regeneration.
- Verify: `git status --short`; `git rev-parse HEAD`; `rg --files tests`; `git ls-files --eol .editorconfig .gitattributes AGENTS.md`.
- Success: every test project and duplicate family is classified; dirty-worktree ownership is recorded.

### Phase 1 - Finalize testing architecture and contracts

- Goal: approve project boundaries, scenario/driver/profile/snapshot contracts, naming, determinism, and tracing schema.
- Files: `test-architecture.md`, `provider-contract-matrix.md`, `decisions.md`, planned namespaces under both new projects.
- Why: implementation must not mix provider construction, expected behavior, and assertions.
- Dependencies: Phase 0.
- Excludes: product capability enum expansion and provider runtime work.
- Verify: architecture checklist; dependency graph review; ensure every proposed rich feature has explicit Supported/Unsupported/Partial/NotApplicable outcome.
- Success: Provider D can conceptually join using only driver + profile + provider-specific tests.

### Phase 2 - Create `Bing.Offices.Testing` skeleton

- Goal: add the dual-TFM non-packable class library and solution reference.
- Files: `tests/Bing.Offices.Testing/Bing.Offices.Testing.csproj`, solution, initial namespace guards.
- Why: common artifacts need a provider-neutral home without xUnit/native packages.
- Dependencies: Phase 1.
- Excludes: mass migration and provider references.
- Verify: `dotnet restore tests/Bing.Offices.Testing/Bing.Offices.Testing.csproj`; `dotnet build tests/Bing.Offices.Testing/Bing.Offices.Testing.csproj -c Release --no-restore`; inspect assets/references for forbidden packages.
- Success: net6/net8 build and an automated dependency-isolation check pass.

### Phase 3 - Add shared models, deterministic data, requests, and fixtures

- Goal: add explicitly named contract models/data, request factories, stream families, observer/committer/temp fixtures.
- Files: Testing `Models`, `Data`, `Requests`, `Streams`, `Fixtures`.
- Why: common inputs and requests are prerequisites for meaningful cross-provider comparison.
- Dependencies: Phase 2.
- Excludes: provider-native types, xUnit asserts, indiscriminate migration of Docs/consumer/benchmark models.
- Verify: Testing build; targeted unit tests in ProviderContract.Tests once available; static forbidden-namespace scan.
- Success: scalar/mapping/dynamic/validation/relation/resource scenarios can be built without provider branches.

### Phase 4 - Add typed expected snapshots and comparers

- Goal: define typed row/workbook/error/relation/style snapshots and stable normalization.
- Files: Testing `Snapshots`, `Comparers`; contract assertion adapters in ProviderContract.Tests.
- Why: string concatenation and NPOI-derived expected data are not independent contracts.
- Dependencies: Phase 3.
- Excludes: provider native style IDs and unstable full exception messages.
- Verify: comparer unit cases for ordering, invariant/zh-CN/en-US, null/empty/blank, stable error fields.
- Success: expected instances are authored directly and no expected result reads another provider's actual result.

### Phase 5 - Add drivers and independent capability profiles

- Goal: implement three public-API drivers, provider case source, execution modes, and test-owned expected profiles.
- Files: ProviderContract `Drivers`, `Profiles`, `Cases`, `Assertions`.
- Why: prevents capability self-declaration from deciding which tests run.
- Dependencies: Phases 3-4.
- Excludes: public capability bit additions.
- Verify: profile completeness test; Expected/Declared/Actual mismatch self-tests; DI first-registration tests.
- Success: all providers are enumerated and unsupported cases execute rather than skip.

### Phase 6 - Build Provider Contract framework

- Goal: add generic scenario runner for sync/async, result/exception/side-effect capture, and trace IDs.
- Files: ProviderContract runner/assertions/base contracts; production-symbol map template.
- Why: each contract needs identical orchestration and observable-state assertions.
- Dependencies: Phase 5.
- Excludes: provider-specific internals and real-file replacement of integration suites.
- Verify: run framework self-tests and one supported/unsupported synthetic case per provider.
- Success: failures identify scenario, provider, mode, format, expected profile, declared capabilities, and actual outcome.

### Phase 7 - Migrate scalar, mapping, converter, value-map, and dynamic contracts

- Goal: migrate the first stable batch from both legacy contract classes.
- Files: `Contracts/{Scalar,Mapping,Converter,ValueMap,DynamicColumn}ContractTest.cs`; update migration/coverage maps.
- Why: proves the architecture on mature common features.
- Dependencies: Phase 6.
- Excludes: deletion of provider-native planner/cache tests.
- Verify: `dotnet test tests/Bing.Offices.ProviderContract.Tests/Bing.Offices.ProviderContract.Tests.csproj -c Release --no-restore --filter "FullyQualifiedName~ScalarContractTest|FullyQualifiedName~MappingContractTest|FullyQualifiedName~DynamicColumnContractTest"`.
- Success: all three XLSX providers match independent expected snapshots; exact legacy replacements are recorded before deletion.

### Phase 8 - Migrate validation, relations, resources, and structured errors

- Goal: unify configured validation, relations, public resource limits, error snapshots, and expected unsupported outcomes.
- Files: corresponding Contract tests; provider coverage/migration matrices.
- Why: these are public correctness and safety contracts with current duplicated assertions.
- Dependencies: Phase 7.
- Excludes: provider ZIP/XML/DOM admission internals and workbook-native validation implementation.
- Verify: targeted contract filters plus NPOI HSSF/XSSF success/failure tests for touched common behavior.
- Success: MaxInputBytes/Rows/Sheets/Columns/Cells/Errors/Unique cases assert complete result or structured error and no partial data.

### Phase 9 - Migrate async, cancellation, streams, exceptions, and atomicity

- Goal: run common scenarios in sync/async modes with shared stream fixtures and observable ownership/target state.
- Files: `AsyncContractTest`, `CancellationContractTest`, `StreamOwnershipContractTest`, `ExceptionContractTest`; provider integration tests.
- Why: async and IO contracts cannot be inferred from value-only round trips.
- Dependencies: Phase 8.
- Excludes: `Task.Run`, `.Result`, `.Wait()`, or replacing real IO with mocks.
- Verify: ProviderContract targeted tests; each provider integration project; cancellation must leave existing targets intact and cleanup temporaries.
- Success: true outer async IO, cancellation, ownership, exception metadata, and atomicity pass for all applicable profiles.

### Phase 10 - Migrate formula, style, merge, template, and entity baseline

- Goal: move rich public outcomes to profile-driven contracts while retaining native assertions.
- Files: Formula/Style/Layout/Template/Entity contracts and provider unit tests.
- Why: these areas have intentional capability differences that must be documented rather than flattened.
- Dependencies: Phases 8-9 and initial Golden fixtures where required.
- Excludes: formula calculation engine, chart/pivot/macro support.
- Verify: targeted contract tests; full NPOI/MiniExcel/ClosedXML unit suites at phase checkpoint.
- Success: every difference is `UNSUPPORTED_EXPECTED`, `PARTIAL`, or `NOT_APPLICABLE`, with no silent skip.

### Phase 11 - Implement ClosedXML Failure Workbook contract-first

- Goal: generate failure output via existing `ExcelImportFailureOptions` with correct annotations and preservation boundaries.
- Files: common Failure Workbook scenario; `ClosedXmlExcelImporter` and new internal writer/helper; ClosedXML unit/integration tests.
- Why: current code rejects all non-None modes in preflight.
- Dependencies: Phases 6, 8, 9, 10.
- Excludes: new public options and unsafe preservation of chart/image/macro/pivot.
- Verify: targeted common contract; ClosedXML unit tests for error coordinates/comment policy/resource limits; real file/stream integration for cancellation/atomic commit/ownership.
- Success: supported modes match common snapshots; unsupported rich templates fail before output; destination is never partially committed.

### Phase 12 - Spike and implement ClosedXML workbook validation

- Goal: map safe workbook-native rules into provider-neutral validation, with explicit policy for every rule/operator.
- Files: `validation-support-matrix.md`; validation Golden XLSX; ClosedXML importer/internal adapter; common and provider tests.
- Why: current `WorkbookRules`/`ConfiguredAndWorkbook` paths are preflight unsupported.
- Dependencies: Golden validation fixtures and Phase 8.
- Excludes: pretending ClosedXML is a full formula engine or silently ignoring custom/range rules.
- Verify: targeted matrix tests for WholeNumber, Decimal, Date, Time, TextLength, List, Custom and all comparison operators; configured+workbook ordering; both TFMs.
- Success: each matrix cell is Supported/Partial/Unsupported with named evidence; errors contain stable coordinates and unsupported rules follow explicit policy.

### Phase 13 - Implement ClosedXML Entity Dynamic Columns; document resource API decision

- Goal: reuse mapping/physical-layout planning for entity list regions and produce a provider-neutral resource-options proposal.
- Files: Entity dynamic contract; `ClosedXmlEntityLayoutExecutor` plus reused internal planner; ClosedXML tests; `api-diff.md` proposal.
- Why: current executor rejects any dynamic column at plan stage.
- Dependencies: Phase 7 dynamic contracts and Phase 10 entity baseline.
- Excludes: a second planner or `ClosedXmlEntityResourceOptions`.
- Verify: fixed+dynamic placement, explicit index, collision/overflow, converter/validation, import dictionary, relations, sync/async and real XLSX tests.
- Success: ClosedXML matches the approved shared scenario; Entity resource API remains `BLOCKED_APPROVAL` unless maintainers explicitly approve a member-level design.

### Phase 14 - Cross-provider regression checkpoint

- Goal: run all provider contracts and affected unit/integration suites, then update evidence by scope.
- Files: reports only unless failures expose an `OPEN_ACTIONABLE` defect.
- Why: shared scenarios and ClosedXML changes cross project/provider boundaries.
- Dependencies: Phases 7-13.
- Excludes: benchmark/resource reruns unrelated to changed hot paths.
- Verify: full ProviderContract.Tests; all three provider Unit and Integration projects; Core tests when Abstractions/Core changed.
- Success: no unclassified differences; any blocked/accepted limitation is separated from implementation/test status.

### Phase 15 - Freeze Golden XLSX and governance

- Goal: commit only independently reviewed provider-neutral resources and manifest hashes.
- Files: Testing `Resources/**`, manifest/README, manifest self-test, optional offline generator.
- Why: runtime provider-generated fixtures compromise independence.
- Dependencies: feature Spikes identify exact fixture needs.
- Excludes: moving legacy Bugs/Purchase resources without consumer audit and checking `bin/obj` copies.
- Verify: hash manifest self-test; inspect package/copy-to-output behavior; run Formula/Date/Validation/Template contracts.
- Success: every Golden file has purpose, expected cells/features, provenance, hash, and update reason.

### Phase 16 - API, consumers, documentation, and migration cleanup

- Goal: finish method-level migration, remove only proven duplicate tests, capture API diff, and protect consumer boundaries.
- Files: migration/coverage/duplication maps, README/docs as needed, solution/tests; dual-TFM API artifacts.
- Why: architecture is incomplete while old duplicate contracts or dependency leaks remain.
- Dependencies: Phases 14-15.
- Excludes: automatic API baseline update, package/version changes.
- Verify: dependency scans; `PublicApiContractTest`; build/run package consumers and third-party consumer from packages; recalculate duplication metrics.
- Success: consumers have no Testing reference; third-party remains public-only; no unapproved public member; each removed test has replacement evidence.

### Phase 17 - Final verification and independent review

- Goal: produce execution/test/integration/resource/final reports and independent review against all gates.
- Files: required execution artifacts, final symbol-to-test map, no business changes unless review finds `OPEN_ACTIONABLE`.
- Why: implementation, tests, release approval, performance/resource evidence, and external gates must remain separate.
- Dependencies: all prior phases.
- Excludes: automatic fix loops for `BLOCKED_APPROVAL`, `BLOCKED_EXTERNAL`, `PARTIAL`, or `NOT_APPLICABLE`.
- Verify: `dotnet test Bing.Offices.sln --no-restore -c Release -v:minimal -m:1`; `git diff --check`; strict UTF-8/BOM/EOL checks; `git diff --stat`; `git ls-files --eol` for modified files; API/consumer gates.
- Success: `OPEN_ACTIONABLE=0`; all 20 requested completion gates have evidence or a correctly classified non-actionable state; independent review reports implementation/test/release statuses separately.

## Verification strategy

- L0 after each edit group: targeted build/static checks, `git diff --check`, encoding/BOM/EOL verification.
- L1 after each scenario/runtime change: exact filtered contract and provider unit tests.
- L2 at migration batch boundaries: affected complete unit projects.
- L3 for stream/file/template/provider changes: affected integration and consumer tests.
- L4 only at Phase 10, Phase 14, and final candidate: affected cross-project/full solution.
- L5 only if a real hot path, harness, or approved resource gate changes. Reuse unchanged ClosedXML 1K/10K/100K evidence; 500K/1M remains approval-gated.

Before every test batch, execution.md must record ChangedFiles, ChangedProjects, public contracts, runtime paths, providers, TFMs, packaging, benchmark harness, dependents, and risk level.

## Completion gates

1. `Bing.Offices.Testing` exists, is provider-neutral, and builds on net6/net8.
2. Testing has no NPOI/MiniExcel/ClosedXML/xUnit dependency.
3. One Provider Contract suite drives all three providers.
4. Expected profiles are independent of production declarations.
5. Expected snapshots are independent of NPOI actual output.
6. Legacy MiniExcel/ClosedXML common contracts are migrated method by method.
7. Shared models/fixtures are no longer semantically duplicated in common contracts.
8. Provider-specific unit-test responsibilities remain covered.
9. Provider integrations retain real IO/library evidence.
10. ClosedXML Failure Workbook passes its public contract or has a classified non-actionable boundary.
11. Workbook Validation has a completed per-rule matrix and passing/partial explicit behavior.
12. Entity Dynamic Columns passes the shared contract or is correctly blocked.
13. NPOI/MiniExcel existing tests have no unexplained regression.
14. ClosedXML existing tests have no unexplained regression.
15. Cross-provider outcomes match expected profiles, including unsupported cases.
16. Package consumers have no Testing dependency.
17. ThirdPartyProvider.Consumer remains public-only.
18. Dual-TFM API snapshots contain no unapproved product API change.
19. Full solution build/test passes except explicitly classified approval/external gates.
20. No automatic commit, push, version bump, dependency upgrade, baseline update, or package publish occurred.

## Required execution artifacts

The executor must create/update `execution.md`, `unit-test-report.md`, `provider-contract-report.md`, `integration-report.md`, `package-consumer-report.md`, `resource-report.md`, `review.md`, `final-report.md`, and a final production-symbol-to-test-method map. The planning matrices in this directory are living execution checklists but may not be replaced by untraceable summaries.

## Stop conditions

Only findings classified `OPEN_ACTIONABLE` permit automatic fixes. Stop when none remain, even if release is blocked by API approval, external CI/OS/font evidence, approved limitations, or resource budgets. Apply the no-progress and maximum-review-loop rules from `GOAL_RULES.md`.

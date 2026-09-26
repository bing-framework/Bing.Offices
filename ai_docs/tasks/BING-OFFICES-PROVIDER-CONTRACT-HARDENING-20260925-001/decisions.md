# Decisions

| ID | Decision | Rationale / status |
| --- | --- | --- |
| D-001 | Build two projects: a provider-neutral library and a provider contract test project. | Separates reusable infrastructure from xUnit/provider construction. |
| D-002 | `Bing.Offices.Testing` references only Abstractions/Core and does not reference xUnit. | Keeps models, requests, snapshots, and fixtures reusable and provider-neutral. |
| D-003 | Expected snapshots are authored from product contracts, never copied from NPOI output. | Removes NPOI as implicit golden provider. |
| D-004 | Expected capability profiles are independent of production capability declarations. | Prevents a missing capability declaration from silently skipping tests. |
| D-005 | Unsupported is an executed contract outcome, not a skipped test. | Verifies exception metadata, no partial output, ownership, and atomicity. |
| D-006 | Migrate with `Add -> Verify -> Migrate -> Delete`; retain old tests until method-level equivalence is recorded. | Avoids evidence gaps during restructuring. |
| D-007 | Provider-native DOM assertions and real IO remain in provider unit/integration projects. | Common contracts assert outcomes, not implementation details. |
| D-008 | Golden XLSX files are frozen, manifested, and hashed; runtime generation by a provider is forbidden. | Keeps fixture production independent from the implementation under test. |
| D-009 | Keep ProfileFixtures, consumers, third-party consumer, benchmarks, and resource probes independent. | Preserves external assembly/public-only/performance semantics. |
| D-010 | ClosedXML Failure Workbook uses existing public options and is contract-first. | No provider-specific API expansion. |
| D-011 | Workbook validation starts with a rule/operator support Spike and explicit fail/partial policy. | ClosedXML object coverage does not imply all Excel validation formulas can be mapped safely. |
| D-012 | Entity Dynamic Columns must reuse the existing mapping/physical-layout planner; no second planner. | Prevents divergent placement and validation semantics. |
| D-013 | Entity resource options were initially `BLOCKED_APPROVAL`. | At planning time the public entity API had no resource input; this initial boundary was superseded by D-017 after explicit user approval. |
| D-014 | Do not implement chart creation, pivot tables, macros, a formula engine, or smart routing. | Explicitly outside task scope. Image/Table remain research-only unless contract and golden evidence are approved. |
| D-015 | Reuse unchanged prior test/benchmark evidence by scope and run fresh broad gates only at phase/final checkpoints. | Applies GOAL_RULES evidence reuse and cost controls. |
| D-016 | Before explicit approval, do not commit, push, tag, publish, bump versions/dependencies, or update the API baseline. | This records the initial execution boundary; the API-baseline portion was superseded by D-018, while the remaining release prohibitions stayed in force. |
| D-017 | Use provider-neutral `ExcelEntityImportOptions` plus opt-in `IExcelEntityResourceImporter`; keep legacy `IExcelEntityImporter` unchanged. | The user approved candidate 1 on 2026-09-25; direct NPOI/ClosedXML and legacy-provider compatibility tests pass on both TFMs. |
| D-018 | Update the dual-TFM API baseline and candidate identity for the explicitly approved additive API and ClosedXML package. | `api-approval.md` records the approval; API compare, `PublicApiContractTest`, and fresh package consumers pass. No version/dependency change or package publication occurred. |

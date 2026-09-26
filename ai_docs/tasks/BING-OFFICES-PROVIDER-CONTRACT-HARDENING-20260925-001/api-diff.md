# API Diff Plan

## Planned non-public additions

- `tests/Bing.Offices.Testing` and its public test-support types are non-packable and are not Bing.Offices product API.
- `tests/Bing.Offices.ProviderContract.Tests` and all drivers/profiles/scenarios are test-only.
- Solution/project graph changes are expected, but package consumers must not reference either new project.

These additions do not require product API baseline changes.

## ClosedXML features

- Failure Workbook should reuse the existing `ExcelImportFailureOptions`; no ClosedXML-specific public option is planned.
- Workbook Data Validation should reuse `ExcelImportValidationMode`, `ExcelUnsupportedFeaturePolicy`, and provider-neutral validation bindings. A Spike must prove whether an internal adapter is sufficient.
- Entity Dynamic Columns should first reuse existing mapping-plan and `ExcelDynamicColumnDefinition` concepts. If the current `ExcelEntityListRegion` snapshot already carries adequate mapping configuration, implementation can remain provider-internal.

## Approval boundary: Entity resources

Current `IExcelEntityImporter`, entity extension methods, and entity import operations accept layout/template/cancellation but no `ExcelResourceLimits` or entity resource options. A provider-private `ClosedXmlEntityResourceOptions` is prohibited.

Candidate provider-neutral designs to document during Phase 13:

1. Add an `ExcelEntityImportOptions` parameter containing `ExcelResourceLimits` to entity import APIs.
2. Add immutable resource options to `ExcelEntityLayout<TEntity>`/builder.

Both change public Abstractions and require dual-TFM snapshot, consumer compilation, member-level approval, migration guidance, and direct NPOI/ClosedXML tests. The user approved candidate 1 on 2026-09-25. Candidate 2 remains rejected for this task because it would couple reusable layouts to per-import resource policy.

### Approved candidate 1

- Add `ExcelEntityImportOptions` with `ExcelResourceLimits ResourceLimits`.
- Add the opt-in `IExcelEntityResourceImporter` SPI with sync/async ordinary and template import methods.
- Add four options-aware `ExcelEntityExtensions` overloads that dispatch only to the new SPI and otherwise fail fast with structured `UnsupportedFeature` context.
- Implement the SPI in NPOI and ClosedXML while preserving the existing `IExcelEntityImporter` unchanged for third-party Providers.
- Apply limits to source buffering, workbook preflight, entity row/error processing, and template imports.
- Approval evidence is recorded in `api-approval.md`; direct NPOI, ClosedXML, and legacy-provider compatibility tests are required before the gate is closed.

## Snapshot gate

Execution must capture net6/net8 current API and compare against `build/api-snapshot-baseline.json`. After the explicit approval recorded in `api-approval.md`, the baseline may be updated once from fresh Release assemblies and packages. Additive, changed, removed, and breaking members must be listed individually. The ClosedXML assembly/package identity is included in the same approved dual-TFM capture.

## Expected product API result

- Preferred result: no public API changes for shared test architecture, Failure Workbook, Workbook Validation, or Entity Dynamic Columns.
- Approved result: provider-neutral entity resource options through a separate opt-in SPI, with legacy entity importer compatibility retained.
- Forbidden: package version bump, dependency upgrade, capability bit reordering, provider-native types in public contracts, or removed public members.

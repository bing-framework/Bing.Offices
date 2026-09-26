# API Migration Record

## Approval

This is an additive change approved by the user through the current advanced Excel implementation plan and recorded in `api-approval.md`:

- `approvedBy`: `user (implementation-plan approval)`
- `approvedAt`: `2026-09-26`
- `taskId`: `BING-OFFICES-ADVANCED-EXCEL-ROADMAP-20260926-001`

No existing public member was removed or had its signature changed. Existing consumers do not require source migration.

## New Opt-In API

Callers that need native worksheet images or data validation can opt in on a Sheet builder:

```csharp
sheet.Image(new ExcelSheetImageDefinition
{
    Content = pngBytes,
    Row = 0,
    Column = 0,
    Width = 120,
    Height = 80
});

sheet.DataValidation(new ExcelDataValidationDefinition
{
    Range = new ExcelRangeDefinition
    {
        StartRow = 1,
        StartColumn = 0,
        EndRow = 100,
        EndColumn = 0
    },
    Type = ExcelDataValidationType.Integer,
    Operator = ExcelConditionalComparisonOperator.GreaterThanOrEqual,
    Value1 = 0,
    IgnoreBlanks = true
});
```

`Image` validates PNG/JPEG bytes, zero-based coordinates, positive dimensions and non-negative offsets. `DataValidation` validates the bounded range, supported comparison values, date/list constraints and prompt lengths before the request is executed.

## Provider Behavior

- NPOI, ClosedXML and SpreadCheetah consume the approved common subset.
- MiniExcel rejects image/data-validation requests during provider preflight; it does not silently ignore them or fall back to another provider.
- Existing requests without these members retain their previous behavior.
- The generic `Stream` target remains caller-owned; file targets continue to use the existing atomic file commit path.

## Package and TFM Compatibility

The public Abstractions surface remains compatible with `netstandard2.0`. Provider packages continue to target `net6.0` and `net8.0`. The package version remains `2.0.0`; no dependency or product version upgrade is part of this migration.

## Verification

The dual-TFM snapshot in `member-api-diff.md` reports only the 38 approved snapshot additions (35 countable public members and 3 type declarations) and `Removed=0`. Package consumers are run against an isolated local feed containing the freshly packed eight Provider/Core packages, so an existing same-version package cannot mask the current public surface.

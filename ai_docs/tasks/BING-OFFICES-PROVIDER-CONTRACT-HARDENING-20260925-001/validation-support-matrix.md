 ClosedXML Workbook Validation Spike Matrix

Historical stage result: the later task `BING-OFFICES-NAMED-LIST-VALIDATION-20260926-001` adds bounded absolute one-dimensional named-list import validation. See that task's `validation-support-matrix.md` and `execution.md` for the new evidence; the unsupported-name row below describes this earlier stage only.

This is the execution result. Status is based on the ClosedXML adapter implementation and named dual-TFM direct/contract evidence. Supported entries use the same eight comparison operators as ClosedXML; formula-based forms outside the stated boundary are explicitly unsupported rather than interpreted as a different rule.

| Validation type | Operators/forms to inspect | Candidate mapping | Execution status | Evidence |
| --- | --- | --- | --- | --- |
| WholeNumber | between, not-between, >, <, =, !=, >=, <= | numeric comparison rule | SUPPORTED | `WorkbookValidationComparisonOperators_ShouldMatchClosedXmlRuleSemantics`; `ClosedXmlWorkbookValidationPipeline.ValidateNumber` |
| Decimal | between, not-between, >, <, =, !=, >=, <= | decimal comparison rule with invariant conversion | SUPPORTED | `WorkbookValidationComparisonOperators_ShouldMatchClosedXmlRuleSemantics`; Golden `DecimalBetweenRule_ShouldAcceptGoldenValue` (2 providers x 2 TFMs) |
| Date | between, not-between, >, <, =, !=, >=, <= | date comparison honoring 1900/1904 system | SUPPORTED | `WorkbookValidationComparisonOperators_ShouldMatchClosedXmlRuleSemantics` uses a 1904 serial boundary; `ValidateDate` delegates to `DateTimeExcelValidationRule` |
| Time | between, not-between, >, <, =, !=, >=, <= | time serial comparison | SUPPORTED | `WorkbookValidationComparisonOperators_ShouldMatchClosedXmlRuleSemantics`; `ValidateDate(..., timeOnly: true)` |
| TextLength | between, not-between, >, <, =, !=, >=, <= | string-length rule | SUPPORTED | `WorkbookValidationComparisonOperators_ShouldMatchClosedXmlRuleSemantics`; Golden `TextLengthRule_ShouldRejectGoldenValue` (2 providers x 2 TFMs) |
| List literal | delimited values | set membership | SUPPORTED | `ExplicitListRule_ShouldReturnWorkbookValidationError` |
| List range | same-sheet/cross-sheet bounded range | resolved set membership | SUPPORTED | `WorkbookValidationBlankAndListRules_ShouldUseExplicitPolicies`; `ResolveListValues` accepts bounded A1 ranges only |
| List name/external/unresolved | named range, external workbook, missing sheet | explicit unsupported | UNSUPPORTED_EXPECTED | `WorkbookValidationUnresolvableReferences_ShouldBeExplicitlyUnsupported` |
| Custom | simple single-cell comparison; complex formula/reference | safe comparison or explicit unsupported | SUPPORTED | `WorkbookValidationSimpleCustomComparison_ShouldValidateResolvedCellValue`; unsupported forms use `WorkbookValidationUnresolvableReferences_ShouldBeExplicitlyUnsupported` |

`WorkbookValidationBlankAndListRules_ShouldUseExplicitPolicies` covers `IgnoreBlanks=true/false`; `ClosedXmlConfiguredAndWorkbook_ShouldRunWorkbookRulesBeforeConfiguredRulesInBothModes` covers stable Workbook-before-configured ordering and sync/async equivalence. Unsupported formulas or references always produce structured errors: `ExcelUnsupportedFeaturePolicy.Report` keeps the row while recording the error, and `Fail` rejects it. The adapter has no current-value or literal fallback for an unresolved reference.

The matrix records `SUPPORTED` and `UNSUPPORTED_EXPECTED`, plus the exact ClosedXML API inspected and test method. The direct rule matrix is deterministic across both target frameworks.

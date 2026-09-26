# Provider Contract Coverage Baseline

Status values: `PASS`, `UNSUPPORTED_EXPECTED`, `PARTIAL`, `BLOCKED`, `NOT_APPLICABLE`, `MISSING`. `Actual` is based on the named dual-TFM execution evidence recorded in `execution.md`.

| Contract | NPOI | MiniExcel | ClosedXML | Shared scenario today | Expected | Actual |
| --- | --- | --- | --- | --- | --- | --- |
| Scalar | PASS | PASS | PASS | `ScalarContractTest.ScalarContract_ShouldMatchIndependentSnapshot` | common expected snapshot | PASS, 6 cases per TFM |
| Mapping/ValueMap | PASS | PASS | PASS | `MappingContractTest.MappingContract_ShouldMatchIndependentSnapshot`; converter contract | common expected snapshot | PASS, 6 cases per TFM |
| Dynamic columns | PASS | PASS | PASS | `DynamicColumnContractTest.DynamicColumnContract_ShouldMatchIndependentSnapshot` | fixed + dynamic placement | PASS, 6 cases per TFM |
| Configured validation | PASS | PASS | PASS | `ValidationContractTest.ValidationContract_ShouldReturnStableErrorSnapshot` | stable error snapshot | PASS, 6 cases per TFM |
| Relations | PASS | PASS | PASS | `RelationContractTest.RelationContract_ShouldMatchIndependentSnapshot` | relation snapshot | PASS, 6 cases per TFM |
| Style | PASS | UNSUPPORTED_EXPECTED | PASS | `RichLayoutContractTest.StyleContract_ShouldMatchProfileAndOoxml` | public style output and structured unsupported result | PASS, 3 provider cases per TFM |
| Merge | PASS | UNSUPPORTED_EXPECTED | PASS | `RichLayoutContractTest.MergeContract_ShouldMatchProfileAndOoxml` | public merge range and structured unsupported result | PASS, 3 provider cases per TFM |
| Template preservation | PASS | UNSUPPORTED_EXPECTED | PASS | `RichLayoutContractTest.TemplateContract_ShouldPreserveDeclaredStyles`; Golden template content tests | declared style preservation and unsupported policy | PASS for declared subset, 3 provider cases per TFM |
| Entity layout | PASS | NOT_APPLICABLE | PASS | `RichLayoutContractTest.EntityContract_ShouldMatchProfileAndRoundTrip` | fixed/list/merge public entity round trip | PASS NPOI/ClosedXML; MiniExcel explicit `NOT_APPLICABLE`, 3 cases per TFM |
| Row height | PASS | UNSUPPORTED_EXPECTED | PASS | `ClosedXmlProviderContractTest.RowHeightContract_ShouldMatchNpoiAndClosedXml` | public layout result | PASS native boundary, 1 case per TFM |
| Formula cache/text | PARTIAL | PARTIAL | PASS | `FormulaGoldenContractTest.FormulaGolden_ShouldMatchIndependentProviderProfile`; `GoldenFixtureContentContractTest.FormulaFixture_ShouldPreserveFormulaAndCachedValue` | NPOI/MiniExcel cached-value policy; ClosedXML formula-text policy | Explicit profile: NPOI/MiniExcel `PARTIAL`, ClosedXML `PASS`, 12 provider cases per TFM |
| Resource limits | PASS | PASS | PASS | `ResourceLimitContractTest` shared matrix plus direct Entity resource SPI tests | structured rejection/no partial data | PASS shared `15/15` matrix per TFM, including rows/sheets/columns/cells/errors/unique and input-byte gates |
| Async/cancellation | PASS | PASS | PASS | `AsyncContractTest.SyncAndAsync_ShouldMatchScalarSnapshot` | common execution modes + real IO | PASS sync/async result equivalence; cancellation remains provider-local |
| Stream ownership | PASS | PASS | PASS | `StreamOwnershipContractTest.AsyncRoundTrip_ShouldLeaveCallerStreamsOpen` | common source/destination state | PASS caller streams remain open |
| Unsupported side effects | PASS | PASS | PASS | Mini contract only | metadata + no output/native start | PARTIAL |
| Failure Workbook | PASS | UNSUPPORTED_EXPECTED | PASS | shared Failure Workbook contract plus real-file boundary tests and ClosedXML native `FailureWorkbook_*` tests | profile then Closed target | PASS supported NPOI/ClosedXML output; Mini expected unsupported; real-file `6/6` per TFM |
| Workbook validation | PASS/PARTIAL | UNSUPPORTED_EXPECTED | PARTIAL | `WorkbookValidationContractTest` with frozen list/whole/decimal/date/time/text-length/custom Golden inputs | rule/operator matrix | PASS 16 direct cases per TFM; PARTIAL full rule matrix |
| Entity dynamic columns | PARTIAL | NOT_APPLICABLE | PASS | ClosedXML entity dynamic round-trip + converter/validation regression | shared layout scenario | PASS ClosedXML; NPOI remains provider-specific |

Execution updates this table only from named test evidence; production capability declarations alone cannot produce `PASS`.

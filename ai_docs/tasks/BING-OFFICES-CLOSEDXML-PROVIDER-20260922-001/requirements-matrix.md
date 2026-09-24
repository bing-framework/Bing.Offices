# Requirements Matrix

| Requirement | Implementation | Evidence | Status |
| --- | --- | --- | --- |
| XLSX List/Workbook export | `ClosedXmlExcelExporter` | ClosedXmlProviderTest.Export | VERIFIED |
| XLSX List/Workbook import | `ClosedXmlExcelImporter` | ClosedXmlProviderTest.RoundTrip | VERIFIED |
| Multiple sheets and selector | request sheet loop; name/index resolution | Unit + real file integration | VERIFIED |
| Core mapping/conversion/dynamic columns | Core mapping plan + ClosedXml adapters + unified physical column planner | fixed/dynamic round-trip, mixed explicit/default fixed index, Before/After, physical index, sparse width and column style tests；Mapping Plan cache isolation/failure recovery | VERIFIED |
| Excel date serial/1904 date system | workbookPr date1904 preflight + date adapter | real modified XLSX serial read-back | VERIFIED_BASIC |
| Validation/unique/relations | validation bindings, unique tracker, relation binder/comparer; Entity `HasMany` coordinator | structured validation, unique limit, Workbook and Entity relation tests | VERIFIED |
| Style/number format/width/row height | font/fill/alignment/border/reset/format/width/row-height adapters | direct style, reset, dynamic style and row-height read-back | VERIFIED |
| Formula write/read contract | formula strings write to `FormulaA1`; string targets preserve formula text and other targets use cached value | formula write/read direct test；NPOI/ClosedXML numeric and boolean cached-value contract with explicit string-policy difference | VERIFIED_BASIC |
| Template stream/style preservation and unsupported-part preflight | `XLWorkbook(Stream)`, `LeaveTemplateOpen`, ZIP entry scan before DOM | template, style, custom header/comment, conflict/overwrite policy, pre-cancel disposal and Chart-part rejection tests | VERIFIED_BASIC |
| Async/cancellation/atomic commit | real stream copy + Core committer | unit/integration tests；sync/async observer, non-seekable input ownership, template leaveOpen/failure/cancel and file failure matrix | VERIFIED_BASIC |
| ZIP/XML/input resource preflight | shared `ExcelXlsxZipPreflight` before DOM；ClosedXML opt-in 图片扫描；行/Sheet/Column/Cell 预算 XML 预检在 `XLWorkbook` 前拒绝；异步 non-seekable bounded buffering | MaxInputBytes exact/over/non-seekable/cancel、ZIP/XML entries、compression ratio、SharedStrings/Styles/Worksheet、图片、MaxRows/MaxSheets/MaxColumns/MaxCells、Unique、MaxErrors、非法 ZIP | VERIFIED |
| Entity layout and DOM admission | `ClosedXmlEntityLayoutExecutor`；fixed cells、multiple list regions、merge、template、async/conversion/relations；shared configurable admission gate | Entity fixed/list/merge/template/async/file/converter/value-map/relations tests；准入并发/队列上限和取消测试 | VERIFIED_BASIC |
| Failure workbook/native validation/chart/pivot/image | fail-fast or template preservation only | unsupported policy code + DOM/output unchanged tests | ACCEPTED_LIMITATION |

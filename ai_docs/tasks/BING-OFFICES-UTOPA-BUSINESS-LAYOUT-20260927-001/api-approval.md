# Public API approval - Entity Layout business extensions

- approvedBy: user (explicit implementation-plan approval)
- approvedAt: 2026-09-27 (Asia/Shanghai)
- taskId: BING-OFFICES-UTOPA-BUSINESS-LAYOUT-20260927-001
- evidence: 本任务计划明确批准属性式固定单元格、动态列分组、未知值策略、Footer 及公开隐藏 SPI；旧 Entity Layout 成员保持兼容。

## Approved member additions (net6.0 / net8.0)

### Consumer-facing layout surface

1. `Bing.Offices.Entities.ExcelEntityCellAttribute(string sheetName, string address)`
2. `Bing.Offices.Entities.ExcelEntityCellAttribute.SheetName`
3. `Bing.Offices.Entities.ExcelEntityCellAttribute.Address`
4. `Bing.Offices.Entities.ExcelEntityCellAttribute.ConverterName`
5. `Bing.Offices.Entities.ExcelEntity.LayoutFromAttributes<TEntity>(Action<ExcelEntityLayoutBuilder<TEntity>> configure = null)`
6. `Bing.Offices.Entities.ExcelEntityLayoutBuilder<TEntity>.CellsFromAttributes()`
7. `Bing.Offices.Entities.ExcelEntityListRegionBuilder<TItem>.DynamicColumnGroup(string groupKey, Expression<Func<TItem, IDictionary<string, object>>> values, IReadOnlyList<ExcelDynamicColumnDefinition> definitions)`
8. `Bing.Offices.Entities.ExcelEntityListRegionBuilder<TItem>.UnknownDynamicValues(ExcelUnknownDynamicValuePolicy policy)`
9. `Bing.Offices.Entities.ExcelEntityListRegionBuilder<TItem>.Footer(string markerText, Action<ExcelEntityListFooterBuilder<TItem>> configure = null)`
10. `Bing.Offices.Entities.ExcelEntityListFooterBuilder<TItem>()`
11. `ExcelEntityListFooterBuilder<TItem>.GapRows(int count)`
12. `ExcelEntityListFooterBuilder<TItem>.Cell(string address, object value, ExcelCellStyle style = null, string numberFormat = null)`
13. `ExcelEntityListFooterBuilder<TItem>.Cell<TValue>(string address, Func<IReadOnlyList<TItem>, TValue> valueFactory, ExcelCellStyle style = null, string numberFormat = null)`
14. `ExcelEntityListFooterBuilder<TItem>.Merge(string range)`
15. `ExcelEntityListFooterBuilder<TItem>.MarkerStyle(ExcelCellStyle style)`

### Provider SPI descriptions

The following public types and members are intentionally marked `EditorBrowsable(Never)` so Provider implementations can consume the immutable layout description without friend access:

- `IExcelEntityDynamicColumnGroup` and `ExcelEntityDynamicColumnGroup<TItem>`: `GroupKey`, `Property`, `Getter`, `Setter`, `Definitions`.
- `IExcelEntityListFooter`, `ExcelEntityListFooter<TItem>`: `MarkerText`, `GapRows`, `Cells`, `Merges`, `MarkerStyle`.
- `IExcelEntityListFooterCell`, `ExcelEntityListFooterCell<TItem>`: `Reference`, `Evaluate`, `Style`, `NumberFormat`, `ValueFactory`.
- `ExcelEntityListRegion<TEntity>.DynamicColumnGroups`, `Footer`, `UnknownDynamicValues`.

The exact net6.0 and net8.0 signatures are recorded in `api-final/capture` and `api-final/compare`; both target frameworks have the same additive set and `removed` is empty.

## Compatibility and migration

All existing `ExcelEntity.Layout`, Fluent cell/list/merge calls, single-dictionary dynamic-column mapping, and Provider entry points remain source and binary compatible. Consumers can opt into attributes and explicit groups incrementally. Provider implementations must use the public SPI descriptions and must not reference internal production types or rely on `InternalsVisibleTo`.

No existing public member was removed or changed in this task. The repository baseline `build/api-snapshot-baseline.json` was updated under this explicit task approval to record the additive candidate identity; this document remains the member-level approval evidence.

## TODO continuation approval

- approvedBy: user (继续实现 TODO 列表能力)
- approvedAt: 2026-09-27 (Asia/Shanghai)
- scope: 为显式动态列组补充内置校验规则，并允许实体导入选择当前行校验失败后的处理策略。

### Additional member additions (net6.0 / net8.0)

1. `Bing.Offices.Entities.ExcelEntityImportOptions(ExcelResourceLimits resourceLimits, ExcelValidationFailureMode validationFailureMode)`
2. `Bing.Offices.Entities.ExcelEntityImportOptions.ValidationFailureMode`
3. `Bing.Offices.Exports.ExcelDynamicColumnDefinition.ValidationRules`

上述成员仅扩展已有实体导入与动态列契约；旧构造函数和默认行为保持不变，默认使用 `StopOnFirstFailure`。

## TODO continuation approval: named anchors

- approvedBy: user (继续实现 TODO 列表能力)
- approvedAt: 2026-09-27 (Asia/Shanghai)
- scope: 为模板实体布局补充命名锚点，减少业务 DTO 对固定 A1 地址的依赖；Provider 通过公开隐藏 SPI 解析运行时坐标。

### Additional member additions (net6.0 / net8.0)

1. `Bing.Offices.Entities.ExcelEntityLayoutBuilder<TEntity>.CellNamed<TValue>(string sheetName, string anchorName, Expression<Func<TEntity, TValue>> property, string converterName = null)`
2. `Bing.Offices.Entities.ExcelEntityLayoutBuilder<TEntity>.CellNamed<TValue>(string sheetName, string anchorName, Expression<Func<TEntity, TValue>> property, Action<ExcelEntityCellBuilder<TValue>> configure)`
3. `Bing.Offices.Entities.ExcelEntityLayoutBuilder<TEntity>.ListRegionNamed<TItem>(string sheetName, string anchorName, Expression<Func<TEntity, IEnumerable<TItem>>> items, Action<ExcelEntityListRegionBuilder<TItem>> configure = null)`
4. `Bing.Offices.Entities.ExcelEntityCellBinding<TEntity>.AnchorName`
5. `Bing.Offices.Entities.ExcelEntityListRegion<TEntity>.AnchorName`
6. `Bing.Offices.Entities.ExcelEntityCellBinding<TEntity>.WithResolvedReference(ExcelEntityCellReference reference)`
7. `Bing.Offices.Entities.ExcelEntityListRegion<TEntity>.WithResolvedStart(ExcelEntityCellReference start)`

`AnchorName`、`WithResolvedReference` 和 `WithResolvedStart` 标记为 `EditorBrowsable(Never)`，仅用于 Provider 消费不可变布局描述，不能替代面向业务的 Fluent API。命名范围必须解析为布局工作表中的单个单元格；局部名称优先于工作簿级名称。命名列表区域不接受绝对 `End`，以避免模板坐标和运行时明细边界产生歧义。

## TODO continuation approval: calculated columns

- approvedBy: user (继续实现 TODO 列表能力)
- approvedAt: 2026-09-27 (Asia/Shanghai)
- scope: 为实体列表布局补充行上下文和导出计算列，保持导入与既有布局兼容。

### Additional member additions (net6.0 / net8.0)

1. `Bing.Offices.Entities.ExcelEntityListRegionBuilder<TItem>.CalculatedColumn<TValue>(string key, string title, Func<ExcelEntityRowContext<TItem>, TValue> valueFactory, Action<ExcelEntityCalculatedColumnBuilder> configure = null)`
2. `Bing.Offices.Entities.ExcelEntityCalculatedColumnBuilder()`
3. `ExcelEntityCalculatedColumnBuilder.Order(int order)`
4. `ExcelEntityCalculatedColumnBuilder.Placement(ExcelColumnPlacement placement)`
5. `ExcelEntityCalculatedColumnBuilder.HeaderStyle(ExcelCellStyle style)`
6. `ExcelEntityCalculatedColumnBuilder.BodyStyle(ExcelCellStyle style)`
7. `ExcelEntityCalculatedColumnBuilder.NumberFormat(string numberFormat)`
8. `ExcelEntityRowContext<TItem>.Item`
9. `ExcelEntityRowContext<TItem>.Items`
10. `ExcelEntityRowContext<TItem>.Index`
11. `ExcelEntityRowContext<TItem>.RowIndex`
12. `ExcelEntityRowContext<TItem>.RowNumber`
13. `ExcelEntityRowContext<TItem>.ColumnIndex`
14. `ExcelEntityRowContext<TItem>.ColumnNumber`
15. `ExcelEntityRowContext<TItem>.SheetName`
16. `ExcelEntityRowContext<TItem>.ColumnKey`

`IExcelEntityCalculatedColumn`、`ExcelEntityCalculatedColumn<TItem>`、`ExcelEntityListRegion<TEntity>.CalculatedColumns` 及计算列定义的 Provider SPI 成员标记为 `EditorBrowsable(Never)`。计算列只参与导出；导入保留物理列位置并跳过实体绑定。计算委托以本次导出的只读明细快照为边界，每个单元格执行一次；取消异常原样透传，其他异常必须包含 Sheet、行、列和 Key 定位。

## TODO continuation: consecutive group subtotals

- approvedBy: user (继续实现 TODO 列表能力)
- approvedAt: 2026-09-27 (Asia/Shanghai)
- scope: 在既有列表 Footer 之后增加连续分组小计；不自动排序，不引入公式 DSL。

### Additional member additions (net6.0 / net8.0)

1. `Bing.Offices.Entities.ExcelEntityListRegionBuilder<TItem>.GroupSubtotal<TKey>(Expression<Func<TItem, TKey>> keySelector, string markerText, Action<ExcelEntityListFooterBuilder<TItem>> configure = null)`

`IExcelEntityGroupSubtotal`、`ExcelEntityGroupSubtotal<TItem, TKey>` 和 `ExcelEntityListRegion<TEntity>.GroupSubtotal` 标记为 `EditorBrowsable(Never)`，仅用于 Provider 消费不可变布局描述。分组按照输入明细的相邻键变化写入小计；导入跳过小计标记、间隔和尾部区域，最终 Footer 仍可同时配置且必须使用不同标记文本。

## TODO continuation approval: entity pagination page breaks

- approvedBy: user (继续实现 TODO 列表能力)
- approvedAt: 2026-09-28 (Asia/Shanghai)
- scope: 为实体列表导出增加每页明细行数的水平分页符；不插入空白行、不改变导入边界，并与连续分组小计保持互斥。

### Additional member additions (net6.0 / net8.0)

1. `Bing.Offices.Entities.ExcelEntityListRegionBuilder<TItem>.PageBreak(int rowsPerPage)`
2. `Bing.Offices.Entities.ExcelEntityListRegion<TEntity>.PageBreakRows`

`PageBreakRows` 标记为 `EditorBrowsable(Never)`，仅用于 Provider 消费不可变布局描述。NPOI 的 XLS/HSSF、XLSX/XSSF 与 ClosedXML XLSX 在导出时写入水平分页符；分页只影响打印分页元数据，不影响明细物化、Footer 或导入 marker 识别。

## TODO continuation approval: entity pagination subtotals

- approvedBy: user (继续实现 TODO 列表能力)
- approvedAt: 2026-09-28 (Asia/Shanghai)
- scope: 在实体列表的水平分页符之后增加按页明细快照计算的分页小计；只写中间页，最终页继续由 Footer 收尾，导入跳过分页小计区域。

### Additional member additions (net6.0 / net8.0)

1. `Bing.Offices.Entities.ExcelEntityListRegionBuilder<TItem>.PageSubtotal(string markerText, Action<ExcelEntityListFooterBuilder<TItem>> configure = null)`
2. `Bing.Offices.Entities.ExcelEntityListRegion<TEntity>.PageSubtotal`

`PageSubtotal` 的布局描述属性标记为 `EditorBrowsable(Never)`，仅供 Provider 消费。该能力必须与 `PageBreak` 同时配置，不能与 `GroupSubtotal` 组合；NPOI XLS/HSSF、XLSX/XSSF 和 ClosedXML XLSX 的导出、导入边界与 Footer 预检保持一致。

## TODO continuation approval: final footer named anchor

- approvedBy: user (继续实现后续阶段内容)
- approvedAt: 2026-09-28 (Asia/Shanghai)
- scope: 为最终 Footer 绑定工作簿命名锚点，使名称随可变长度明细移动；中间分页和分组小计不绑定单格名称。

### Additional member additions (net6.0 / net8.0)

1. `Bing.Offices.Entities.ExcelEntityListRegionBuilder<TItem>.FooterNamed(string anchorName, string markerText, Action<ExcelEntityListFooterBuilder<TItem>> configure = null)`
2. `Bing.Offices.Entities.ExcelEntityListRegion<TEntity>.FooterAnchorName`

`FooterAnchorName` 标记为 `EditorBrowsable(Never)`，仅供 Provider 消费不可变布局描述。两个 TFM 的候选差异均仅含上述两个新增成员，`removed=[]`；保留旧 `Footer` 和全部构造签名。模板中同名且同工作表的单格名称重新定位，其他工作表、非单格和歧义名称在 Plan 阶段拒绝；导入仍使用 marker 文本确定边界。

## TODO continuation approval: explicit footer formula

- approvedBy: user (继续实现 TODO 列表的能力，按照阶段建议实现)
- approvedAt: 2026-09-28 (Asia/Shanghai)
- scope: Entity Layout Footer 增加显式 A1 公式写入；保留既有 `Cell` 与 Provider SPI 接口签名。

### Additional member additions (net6.0 / net8.0)

1. `Bing.Offices.Entities.ExcelEntityListFooterBuilder<TItem>.Formula(string address, string formula, ExcelCellStyle style = null, string numberFormat = null)`
2. `Bing.Offices.Entities.ExcelEntityFooterFormulaValue`
3. `Bing.Offices.Entities.ExcelEntityFooterFormulaValue.FormulaA1`

公式值描述类型标记为 `EditorBrowsable(Never)`，仅供 Provider 从现有 Footer Cell `Evaluate` 结果识别原生公式。公式文本不由框架解析、改写或求值；无公共成员删除或签名变化。

## TODO continuation approval: contiguous footer sum formula

- approvedBy: user (继续实现 TODO 列表的能力，按照建议的实现方式推进进度)
- approvedAt: 2026-09-28 (Asia/Shanghai)
- scope: 为连续明细的 Footer 增加按实际行数生成的原生求和公式；非连续的最终总计仍使用聚合值或显式公式。

### Additional member additions (net6.0 / net8.0)

1. `Bing.Offices.Entities.ExcelEntityListFooterBuilder<TItem>.FormulaSumContiguousRowsAbove(string address, string sourceColumn, ExcelCellStyle style = null, string numberFormat = null)`

双 TFM 只新增此方法，无公共成员删除或签名变化。`sourceColumn` 仅接受 A 到 XFD 列字母；空明细生成 `=0`，最终 Footer 与分页/分组小计共存时在布局构建阶段拒绝该便捷方法。

## TODO continuation approval: final footer detail-row sum

- approvedBy: user (继续实现 TODO 列表的能力)
- approvedAt: 2026-09-28 (Asia/Shanghai)
- scope: 最终 Footer 在分页或分组小计后仍能生成只引用真实明细行的原生总计公式。

### Additional member additions (net6.0 / net8.0)

1. `Bing.Offices.Entities.ExcelEntityListFooterBuilder<TItem>.FormulaSumDetailRowsAbove(string address, string sourceColumn, ExcelCellStyle style = null, string numberFormat = null)`
2. `Bing.Offices.Entities.ExcelEntityFooterDetailSumValue`
3. `Bing.Offices.Entities.ExcelEntityFooterDetailSumValue.SourceColumn`
4. `Bing.Offices.Entities.ExcelEntityFooterDetailSumValue.ToFormulaA1(IReadOnlyList<(int FirstRow, int LastRow)> detailRows)`

值对象标记为 `EditorBrowsable(Never)`，供第三方 Provider 从 Footer Cell 的 `Evaluate` 结果中识别，并传入实际写出的明细行段。只引用明细闭区间，空明细写入 `=0`；多段使用原生 `SUM`，过长公式明确拒绝。保留全部既有公开接口和构造签名。

## TODO continuation approval: explicit workbook export strategy

- approvedBy: user (继续实现 TODO 列表的能力，按照建议的实现方式推进进度)
- approvedAt: 2026-09-28 (Asia/Shanghai)
- taskId: BING-OFFICES-WORKBOOK-EXPORT-STRATEGY-20260928-001
- scope: 对现有完整工作簿与前向流式接口提供逐报表显式选择；不改变 Provider 内部算法或旧接口。

### Additional member additions (net6.0 / net8.0)

1. `Bing.Offices.Exports.ExcelWorkbookExportMode`：`CompleteWorkbook`、`ForwardStreaming`。
2. `Bing.Offices.Exports.ExcelWorkbookExportStrategy(IExcelExporter completeExporter = null, IExcelStreamingExporter streamingExporter = null)`。
3. `ExcelWorkbookExportStrategy.Export`、`ExportAsync`、`ExportToFile`、`ExportToFileAsync`，均明确传入本次报表模式。

分类为 User API。完整模式沿用现有完整工作簿接口，Provider 可在内部优化；前向模式只调用流式接口并在模板、格式或能力不匹配时预检拒绝。仅新增类型与成员，无公共成员删除或旧签名变化。

using Bing.Offices.Exports;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Exports.PivotTables;

/// <summary>
/// 创建完整缓存的原生 XLSX 透视表。
/// </summary>
internal static class NpoiPivotTableWriter
{
    // NPOI 原生 AddColumnLabel 使用该哨兵表示普通度量没有基准字段或项目。
    private const int DefaultDataFieldBaseField = -1;
    private const uint DefaultDataFieldBaseItem = 1048832;

    /// <summary>
    /// 串行初始化原生透视创建所需的进程级内建样式。
    /// </summary>
    /// <remarks>
    /// 原生默认样式首次并发加载可能进入尚未完成的样式查找；仅初始化共享样式，不串行化工作簿写入。
    /// </remarks>
    private static readonly Lazy<ITableStyle> DefaultStyle = new(
        () => XSSFBuiltinTableStyle.GetStyle(XSSFBuiltinTableStyleEnum.PivotStyleLight16),
        LazyThreadSafetyMode.ExecutionAndPublication);

    /// <summary>
    /// 创建原生字段并完成全部来源缓存。
    /// </summary>
    /// <param name="workbook">来源所属的 XSSF 工作簿。</param>
    /// <param name="source">已经完成来源和容量预检的计划。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>待随工作簿序列化的完整缓存记录。</returns>
    internal static NpoiPivotCacheRecords Create(XSSFWorkbook workbook, NpoiPivotSourcePlan source,
        CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var definition = source.Definition;
        if (workbook.GetSheet(definition.TargetSheet) != null)
            throw NpoiPivotSourceValidator.Invalid(definition, definition.SourceRange.StartRow,
                definition.SourceRange.StartColumn, $"目标工作表已存在: {definition.TargetSheet}。");
        _ = DefaultStyle.Value;
        token.ThrowIfCancellationRequested();
        var target = (XSSFSheet)workbook.CreateSheet(definition.TargetSheet);
        var range = definition.SourceRange;
        var area = new AreaReference(new CellReference(range.StartRow, range.StartColumn),
            new CellReference(range.EndRow, range.EndColumn), SpreadsheetVersion.EXCEL2007);
        var pivot = target.CreatePivotTable(area,
            new CellReference(definition.TargetRow, definition.TargetColumn), source.Sheet);
        var table = pivot.GetCTPivotTableDefinition();
        table.name = NpoiPivotTextEncoding.Encode(definition.Name);
        table.rowGrandTotals = true;
        table.colGrandTotals = true;
        table.dataOnRows = false;
        if (definition.ValueFields.Count == 0)
        {
            ApplyAxis(table, source, new[] { new ExcelPivotAxisFieldDefinition { Field = definition.RowField } },
                ST_Axis.axisRow, token);
            ApplyValues(workbook, table, source, new[] { new ExcelPivotValueDefinition
            {
                Field = definition.ValueField, Caption = definition.ValueField,
                Aggregation = ExcelPivotAggregation.Sum
            } }, token);
        }
        else
        {
            table.compact = false;
            table.compactData = false;
            table.outline = false;
            table.outlineData = false;
            ApplyAxis(table, source, definition.RowFields, ST_Axis.axisRow, token);
            ApplyAxis(table, source, definition.ColumnFields, ST_Axis.axisCol, token);
            ApplyValues(workbook, table, source, definition.ValueFields, token);
        }
        ApplyFilterFields(table, source, token);
        var records = NpoiPivotCacheWriter.Complete(pivot, source.Headers, source.Rows, token);
        ApplyFilterSelections(table, source, records, token);
        // 缓存保存全部来源；显示与刷新由客户端原生透视完成，不预写静态聚合值。
        pivot.GetPivotCacheDefinition().GetCTPivotCacheDefinition().refreshOnLoad = true;
        token.ThrowIfCancellationRequested();
        return records;
    }

    /// <summary>
    /// 配置有序行轴或列轴字段。
    /// </summary>
    /// <param name="table">原生透视定义。</param>
    /// <param name="source">完整来源计划。</param>
    /// <param name="definitions">按层级顺序排列的轴字段。</param>
    /// <param name="axis">行轴或列轴。</param>
    /// <param name="token">取消令牌。</param>
    private static void ApplyAxis(CT_PivotTableDefinition table, NpoiPivotSourcePlan source,
        IReadOnlyList<ExcelPivotAxisFieldDefinition> definitions, ST_Axis axis, CancellationToken token)
    {
        var fields = new List<CT_Field>();
        foreach (var definition in definitions)
        {
            token.ThrowIfCancellationRequested();
            var index = source.Columns[definition.Field];
            var field = table.pivotFields.pivotField[index];
            field.axis = axis;
            field.axisSpecified = true;
            field.showAll = false;
            field.compact = false;
            field.outline = false;
            field.subtotalTop = false;
            field.defaultSubtotal = definition.SubtotalMode == ExcelPivotSubtotalMode.Automatic;
            field.sortType = definition.SortOrder switch
            {
                ExcelPivotSortOrder.Default => ST_FieldSortType.manual,
                ExcelPivotSortOrder.Ascending => ST_FieldSortType.ascending,
                ExcelPivotSortOrder.Descending => ST_FieldSortType.descending,
                _ => throw new ArgumentOutOfRangeException(nameof(definition.SortOrder))
            };
            fields.Add(new CT_Field { x = index });
        }
        if (axis == ST_Axis.axisRow)
            table.rowFields = fields.Count == 0 ? null : new CT_RowFields
            { field = fields, count = checked((uint)fields.Count) };
        else
            table.colFields = fields.Count == 0 ? null : new CT_ColFields
            { field = fields, count = checked((uint)fields.Count) };
    }

    /// <summary>
    /// 配置按显示顺序排列的度量与合成 Values 轴。
    /// </summary>
    /// <param name="workbook">负责保存数字格式的工作簿。</param>
    /// <param name="table">原生透视定义。</param>
    /// <param name="source">完整来源计划。</param>
    /// <param name="definitions">有序度量定义。</param>
    /// <param name="token">取消令牌。</param>
    private static void ApplyValues(XSSFWorkbook workbook, CT_PivotTableDefinition table,
        NpoiPivotSourcePlan source, IReadOnlyList<ExcelPivotValueDefinition> definitions, CancellationToken token)
    {
        var fields = new List<CT_DataField>();
        foreach (var definition in definitions)
        {
            token.ThrowIfCancellationRequested();
            var index = source.Columns[definition.Field];
            var pivotField = table.pivotFields.pivotField[index];
            pivotField.dataField = true;
            pivotField.defaultSubtotal = false;
            pivotField.subtotalTop = false;
            pivotField.showAll = false;
            var field = new CT_DataField
            {
                fld = checked((uint)index), name = NpoiPivotTextEncoding.Encode(definition.Caption),
                showDataAs = ST_ShowDataAs.normal,
                baseField = DefaultDataFieldBaseField,
                baseItem = DefaultDataFieldBaseItem,
                subtotal = definition.Aggregation switch
                {
                    ExcelPivotAggregation.Sum => ST_DataConsolidateFunction.sum,
                    ExcelPivotAggregation.Count => ST_DataConsolidateFunction.count,
                    ExcelPivotAggregation.Average => ST_DataConsolidateFunction.average,
                    ExcelPivotAggregation.Minimum => ST_DataConsolidateFunction.min,
                    ExcelPivotAggregation.Maximum => ST_DataConsolidateFunction.max,
                    _ => throw new ArgumentOutOfRangeException(nameof(definition.Aggregation))
                }
            };
            if (definition.NumberFormat != null)
            {
                field.numFmtId = checked((uint)workbook.CreateDataFormat().GetFormat(definition.NumberFormat));
                field.numFmtIdSpecified = true;
            }
            fields.Add(field);
        }
        table.dataFields = new CT_DataFields
        { dataField = fields, count = checked((uint)fields.Count), countSpecified = true };
        if (fields.Count > 1)
        {
            table.colFields ??= new CT_ColFields { field = new List<CT_Field>() };
            table.colFields.field.Add(new CT_Field { x = -2 });
            table.colFields.count = checked((uint)table.colFields.field.Count);
        }
    }

    /// <summary>
    /// 配置报表筛选字段及整体目标位置。
    /// </summary>
    /// <param name="table">原生透视定义。</param>
    /// <param name="source">完整来源计划。</param>
    /// <param name="token">取消令牌。</param>
    private static void ApplyFilterFields(CT_PivotTableDefinition table, NpoiPivotSourcePlan source,
        CancellationToken token)
    {
        var definition = source.Definition;
        var fields = new List<CT_PageField>();
        foreach (var filter in definition.Filters)
        {
            token.ThrowIfCancellationRequested();
            var index = source.Columns[filter.Field];
            var field = table.pivotFields.pivotField[index];
            field.axis = ST_Axis.axisPage;
            field.axisSpecified = true;
            field.defaultSubtotal = false;
            field.showAll = false;
            fields.Add(new CT_PageField { fld = index, hier = -1, hierSpecified = true });
        }
        table.pageFields = fields.Count == 0 ? null : new CT_PageFields
        { pageField = fields, count = checked((uint)fields.Count), countSpecified = true };
        table.multipleFieldFilters = fields.Count > 1;
        table.pageOverThenDown = false;
        var bodyRow = checked(definition.TargetRow + (fields.Count == 0 ? 0 : fields.Count + 1));
        table.location.@ref = new AreaReference(new CellReference(bodyRow, definition.TargetColumn),
            new CellReference(definition.TargetRow + source.OutputRows - 1,
                definition.TargetColumn + source.OutputColumns - 1), SpreadsheetVersion.EXCEL2007).FormatAsString();
        table.location.firstHeaderRow = 0;
        table.location.firstDataRow = checked((uint)(Math.Max(1, definition.ColumnFields.Count)
            + (definition.ValueFields.Count > 1 ? 1 : 0)));
        table.location.firstDataCol = checked((uint)Math.Max(1, definition.RowFields.Count));
        table.location.rowPageCount = checked((uint)fields.Count);
        table.location.colPageCount = fields.Count == 0 ? 0u : 1u;
    }

    /// <summary>
    /// 将类型化筛选选项映射到完整缓存的项目索引。
    /// </summary>
    /// <param name="table">原生透视定义。</param>
    /// <param name="source">保留全部来源的计划。</param>
    /// <param name="records">逐行逐字段的共享项引用。</param>
    /// <param name="token">取消令牌。</param>
    private static void ApplyFilterSelections(CT_PivotTableDefinition table, NpoiPivotSourcePlan source,
        NpoiPivotCacheRecords records, CancellationToken token)
    {
        for (var filterIndex = 0; filterIndex < source.Definition.Filters.Count; filterIndex++)
        {
            token.ThrowIfCancellationRequested();
            var definition = source.Definition.Filters[filterIndex];
            var selected = source.Filters[definition.Field];
            if (selected.Count == 0)
                continue;
            var keys = new HashSet<string>(selected.Select(NpoiPivotSourceValidator.ValueKey),
                StringComparer.OrdinalIgnoreCase);
            var column = source.Columns[definition.Field];
            var indexes = new HashSet<uint>();
            for (var row = 0; row < source.Rows.Count; row++)
            {
                token.ThrowIfCancellationRequested();
                if (keys.Contains(NpoiPivotSourceValidator.ValueKey(source.Rows[row][column])))
                    indexes.Add(records.Rows[row][column]);
            }
            var field = table.pivotFields.pivotField[column];
            var page = table.pageFields.pageField[filterIndex];
            if (indexes.Count == 1)
            {
                page.item = indexes.Single();
                page.itemSpecified = true;
            }
            else
            {
                field.multipleItemSelectionAllowed = true;
                foreach (var item in field.items.item)
                    item.h = !indexes.Contains(item.x);
            }
        }
    }
}

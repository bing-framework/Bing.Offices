using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using Bing.Offices.Attributes;
using Bing.Offices.Exports;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Resolvers;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Exports;

/// <summary>
/// 负责单个 Workbook Sheet 的表头、单元格、样式、布局和图表写入。
/// </summary>
internal sealed class NpoiExportSheetWriter
{
    /// <summary>
    /// 写入一个已编译映射计划对应的 Sheet。
    /// </summary>
    internal void Write<T>(IWorkbook workbook, ExcelSheetExportRequest request,
        CancellationToken cancellationToken, IExcelMappingPlan mapping,
        IReadOnlyList<ExcelColumnPlan> columns, int originRow, int originColumn)
        where T : class, new()
    {
        var sheet = workbook.GetSheet(request.Name) ?? workbook.CreateSheet(request.Name);
        if (request.Hidden)
            workbook.SetSheetVisibility(workbook.GetSheetIndex(sheet), SheetVisibility.Hidden);
        var headerRowIndex = originRow + request.HeaderRowIndex;
        var header = sheet.GetRow(headerRowIndex) ?? sheet.CreateRow(headerRowIndex);
        var dynamicKeys = request.FailOnUnknownDynamicValues
            ? new HashSet<string>(columns.Where(column => column.IsDynamic).Select(column => column.Key),
                StringComparer.Ordinal)
            : null;
        WriteCustomHeaders(sheet, request.HeaderRows, originRow, originColumn,
            request.CommentConflictPolicy, request.TemplateCellOverwritePolicy);
        for (var index = 0; index < columns.Count; index++)
        {
            var cell = header.GetCell(originColumn + index) ?? header.CreateCell(originColumn + index);
            PrepareTemplateCell(workbook, cell, request.TemplateCellOverwritePolicy);
            cell.SetCellValue(columns[index].Title);
        }

        var rowIndex = originRow + request.DataRowStartIndex;
        foreach (var item in request.Data.Cast<T>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var dynamicValues = request.DynamicGetter?.Invoke(item);
            ValidateUnknownDynamicValues(request, dynamicValues, dynamicKeys);
            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                var physicalColumnIndex = originColumn + columnIndex;
                var cell = row.GetCell(physicalColumnIndex) ?? row.CreateCell(physicalColumnIndex);
                PrepareTemplateCell(workbook, cell, request.TemplateCellOverwritePolicy);
                WriteCell(cell, item, columns[columnIndex], dynamicValues, sheet.SheetName,
                    rowIndex + 1, physicalColumnIndex + 1, request.Culture);
            }
            rowIndex++;
        }

        ApplyRequestStyles(workbook, sheet, header, columns, request, mapping, originColumn,
            originRow + request.DataRowStartIndex, rowIndex);
        ApplyHeaderStyle<T>(workbook, sheet, headerRowIndex);
        ApplyWrapText<T>(sheet);
        ApplyColumnWidths(sheet, columns, originColumn, originRow, request.ColumnWidth);
        MergeColumns<T>(sheet, columns, originRow + request.DataRowStartIndex, rowIndex - 1, originColumn);
        CreateCharts(workbook, sheet, columns, request.Charts, originRow + request.DataRowStartIndex,
            rowIndex, originColumn);
    }

    /// <summary>
    /// 验证动态字典中的键都已声明为导出列。
    /// </summary>
    private static void ValidateUnknownDynamicValues(ExcelSheetExportRequest request,
        IDictionary<string, object> values, ISet<string> dynamicKeys)
    {
        if (!request.FailOnUnknownDynamicValues || values == null || dynamicKeys == null)
            return;
        var unknown = values.Keys.FirstOrDefault(key => !dynamicKeys.Contains(key));
        if (unknown != null)
            throw new InvalidOperationException($"动态值未声明: {unknown}");
    }

    /// <summary>
    /// 将 provider-neutral 图表定义转换为 XSSF 图表。
    /// </summary>
    private static void CreateCharts(IWorkbook workbook, ISheet sheet, IReadOnlyList<ExcelColumnPlan> columns,
        IReadOnlyList<ExcelChartDefinition> charts, int dataStartRow, int dataEndRow, int firstColumnIndex)
    {
        if (charts == null || charts.Count == 0)
            return;
        if (!(workbook is XSSFWorkbook) || !(sheet.CreateDrawingPatriarch() is XSSFDrawing drawing))
            throw new NotSupportedException("当前 Excel 格式不支持图表。");
        foreach (var definition in charts)
        {
            definition.Validate();
            var categoryColumn = ResolveChartColumn(columns, definition.Categories.ColumnKey);
            var categoryRange = CreateChartRange(sheet, definition.Categories, firstColumnIndex + categoryColumn,
                dataStartRow, dataEndRow);
            var anchor = definition.Anchor;
            var chartAnchor = drawing.CreateAnchor(0, 0, 0, 0, anchor.StartColumn, anchor.StartRow,
                anchor.EndColumn, anchor.EndRow);
            var chart = drawing.CreateChart(chartAnchor);
            if (!string.IsNullOrWhiteSpace(definition.Title))
                chart.SetTitle(definition.Title);
            var categories = DataSources.FromStringCellRange(sheet, categoryRange);
            switch (definition.Type)
            {
                case ExcelChartType.Column:
                {
                    var data = chart.ChartDataFactory.CreateColumnChartData<string, double>();
                    foreach (var seriesDefinition in definition.Series)
                    {
                        var valueColumn = ResolveChartColumn(columns, seriesDefinition.Values.ColumnKey);
                        var values = DataSources.FromNumericCellRange(sheet,
                            CreateChartRange(sheet, seriesDefinition.Values, firstColumnIndex + valueColumn,
                                dataStartRow, dataEndRow));
                        data.AddSeries(categories, values).SetTitle(seriesDefinition.Name);
                    }
                    var categoryAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
                    var valueAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
                    valueAxis.Crosses = AxisCrosses.AutoZero;
                    chart.Plot(data, categoryAxis, valueAxis);
                    break;
                }
                case ExcelChartType.Line:
                {
                    var data = chart.ChartDataFactory.CreateLineChartData<string, double>();
                    foreach (var seriesDefinition in definition.Series)
                    {
                        var valueColumn = ResolveChartColumn(columns, seriesDefinition.Values.ColumnKey);
                        var values = DataSources.FromNumericCellRange(sheet,
                            CreateChartRange(sheet, seriesDefinition.Values, firstColumnIndex + valueColumn,
                                dataStartRow, dataEndRow));
                        data.AddSeries(categories, values).SetTitle(seriesDefinition.Name);
                    }
                    var categoryAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
                    var valueAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
                    valueAxis.Crosses = AxisCrosses.AutoZero;
                    chart.Plot(data, categoryAxis, valueAxis);
                    break;
                }
                case ExcelChartType.Pie:
                {
                    var data = chart.ChartDataFactory.CreatePieChartData<string, double>();
                    var seriesDefinition = definition.Series[0];
                    var valueColumn = ResolveChartColumn(columns, seriesDefinition.Values.ColumnKey);
                    var values = DataSources.FromNumericCellRange(sheet,
                        CreateChartRange(sheet, seriesDefinition.Values, firstColumnIndex + valueColumn,
                            dataStartRow, dataEndRow));
                    data.AddSeries(categories, values).SetTitle(seriesDefinition.Name);
                    chart.Plot(data);
                    break;
                }
                default:
                    throw new NotSupportedException($"不支持的图表类型: {definition.Type}");
            }
        }
    }

    /// <summary>
    /// 解析图表列引用。
    /// </summary>
    private static int ResolveChartColumn(IReadOnlyList<ExcelColumnPlan> columns, string key)
    {
        var index = columns.ToList().FindIndex(column => string.Equals(column.Key, key,
            StringComparison.OrdinalIgnoreCase)
            || string.Equals(column.Title, key, StringComparison.OrdinalIgnoreCase)
            || string.Equals(column.Property.Name, key, StringComparison.OrdinalIgnoreCase));
        if (index < 0)
            throw new ArgumentException($"图表引用的列不存在: {key}", nameof(key));
        return index;
    }

    /// <summary>
    /// 创建图表使用的单列区域。
    /// </summary>
    private static CellRangeAddress CreateChartRange(ISheet sheet, ExcelChartRange range, int columnIndex,
        int defaultStartRow, int defaultEndRow)
    {
        var startRow = range.StartRow ?? defaultStartRow;
        var endRow = range.EndRow ?? defaultEndRow;
        if (startRow < 0 || endRow <= startRow || endRow > sheet.LastRowNum + 1)
            throw new ArgumentException($"图表范围行索引无效: {range.ColumnKey}", nameof(range));
        return new CellRangeAddress(startRow, endRow - 1, columnIndex, columnIndex);
    }

    /// <summary>
    /// 应用 Workbook 请求级区域和动态列样式。
    /// </summary>
    private static void ApplyRequestStyles(IWorkbook workbook, ISheet sheet, IRow header,
        IReadOnlyList<ExcelColumnPlan> columns, ExcelSheetExportRequest request, IExcelMappingPlan mapping,
        int firstColumnIndex, int firstDataRowIndex, int lastRowIndex)
    {
        var mappingHeaderStyle = ResolveStyle(mapping.Style?.HeaderStyleKey, true);
        var mappingBodyStyle = ResolveStyle(mapping.Style?.BodyStyleKey, false);
        if (request.HeaderStyle == null && request.BodyStyle == null && request.SheetStyle == null
            && mappingHeaderStyle == null && mappingBodyStyle == null
            && columns.All(column => column.HeaderStyle == null && column.BodyStyle == null))
            return;
        for (var index = 0; index < columns.Count; index++)
        {
            var style = columns[index].HeaderStyle ?? request.HeaderStyle ?? mappingHeaderStyle ?? request.SheetStyle;
            if (style != null)
            {
                var cell = header.GetCell(firstColumnIndex + index);
                cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, style);
            }
        }
        for (var rowIndex = firstDataRowIndex; rowIndex < lastRowIndex; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
                continue;
            for (var index = 0; index < columns.Count; index++)
            {
                var style = columns[index].BodyStyle ?? request.BodyStyle ?? mappingBodyStyle ?? request.SheetStyle;
                if (style != null)
                {
                    var cell = row.GetCell(firstColumnIndex + index);
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, style);
                }
            }
        }
    }

    /// <summary>
    /// 解析有限的内置映射样式键。
    /// </summary>
    private static Bing.Offices.Styles.ExcelCellStyle ResolveStyle(string key, bool header)
    {
        if (string.IsNullOrWhiteSpace(key))
            return null;
        if (string.Equals(key, "header", StringComparison.OrdinalIgnoreCase)
            || string.Equals(key, "bold", StringComparison.OrdinalIgnoreCase))
            return new Bing.Offices.Styles.ExcelCellStyle { Bold = true };
        if (string.Equals(key, "body", StringComparison.OrdinalIgnoreCase)
            || string.Equals(key, "default", StringComparison.OrdinalIgnoreCase))
            return new Bing.Offices.Styles.ExcelCellStyle();
        throw new InvalidOperationException($"未注册的{(header ? "表头" : "正文")}样式键: {key}");
    }

    /// <summary>
    /// 写入自定义表头、模板批注和合并区域。
    /// </summary>
    private static void WriteCustomHeaders(ISheet sheet, IReadOnlyList<ExcelHeaderRow> headerRows,
        int originRow, int originColumn, ExcelCommentConflictPolicy commentConflictPolicy,
        ExcelTemplateCellOverwritePolicy overwritePolicy)
    {
        foreach (var headerRow in headerRows ?? Array.Empty<ExcelHeaderRow>())
        {
            var rowIndex = originRow + headerRow.RowIndex;
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            foreach (var headerCell in headerRow.Cells)
            {
                var columnIndex = originColumn + headerCell.ColumnIndex;
                var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
                PrepareTemplateCell(sheet.Workbook, cell, overwritePolicy);
                cell.SetValue(headerCell.Value);
                if (headerCell.Comment != null)
                    ApplyComment(sheet, cell, headerCell.Comment, commentConflictPolicy);
                if (headerCell.RowSpan > 1 || headerCell.ColumnSpan > 1)
                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex + headerCell.RowSpan - 1,
                        columnIndex, columnIndex + headerCell.ColumnSpan - 1));
            }
        }
    }

    /// <summary>
    /// 按模板覆盖策略保留或清理单元格样式和批注。
    /// </summary>
    private static void PrepareTemplateCell(IWorkbook workbook, ICell cell,
        ExcelTemplateCellOverwritePolicy overwritePolicy)
    {
        if (overwritePolicy != ExcelTemplateCellOverwritePolicy.ReplaceTemplate)
            return;
        cell.CellStyle = workbook.GetCellStyleAt(0);
        cell.RemoveCellComment();
    }

    /// <summary>
    /// 按批注冲突策略写入单元格批注。
    /// </summary>
    private static void ApplyComment(ISheet sheet, ICell cell, ExcelComment comment,
        ExcelCommentConflictPolicy conflictPolicy)
    {
        var existing = cell.CellComment;
        if (existing != null)
        {
            switch (conflictPolicy)
            {
                case ExcelCommentConflictPolicy.Preserve:
                    return;
                case ExcelCommentConflictPolicy.Fail:
                    throw new InvalidOperationException($"单元格已有批注: {cell.Address}");
                case ExcelCommentConflictPolicy.Append:
                    comment = new ExcelComment(existing.String.String + Environment.NewLine + comment.Text,
                        string.IsNullOrWhiteSpace(comment.Author) ? existing.Author : comment.Author, comment.Visible);
                    break;
                case ExcelCommentConflictPolicy.Replace:
                    cell.RemoveCellComment();
                    break;
            }
        }
        var anchor = sheet.Workbook.GetCreationHelper().CreateClientAnchor();
        anchor.Col1 = cell.ColumnIndex;
        anchor.Col2 = cell.ColumnIndex + 2;
        anchor.Row1 = cell.RowIndex;
        anchor.Row2 = cell.RowIndex + 3;
        var created = sheet.CreateDrawingPatriarch().CreateCellComment(anchor);
        created.String = sheet.Workbook.GetCreationHelper().CreateRichTextString(comment.Text);
        created.Author = comment.Author;
        created.Visible = comment.Visible;
        cell.CellComment = created;
    }

    /// <summary>
    /// 应用实体表头特性定义的字体样式。
    /// </summary>
    private static void ApplyHeaderStyle<T>(IWorkbook workbook, ISheet sheet, int headerRowIndex)
        where T : class, new()
    {
        var attribute = typeof(T).GetCustomAttributes(typeof(HeaderAttribute), false).Cast<HeaderAttribute>()
            .SingleOrDefault();
        if (attribute == null)
            return;
        var row = sheet.GetRow(headerRowIndex);
        if (row == null)
            return;
        foreach (var cell in row.Cells)
            cell.CellStyle = NpoiStyleCache.ApplyHeaderAttribute(workbook, cell.CellStyle, attribute);
    }

    /// <summary>
    /// 应用固定、自动或自适应列宽。
    /// </summary>
    private static void ApplyColumnWidths(ISheet sheet, IReadOnlyList<ExcelColumnPlan> columns,
        int firstColumnIndex, int originRow, ExcelColumnWidthOptions options)
    {
        if (options == null || options.Mode == ExcelColumnWidthMode.None)
            return;
        for (var index = 0; index < columns.Count; index++)
        {
            var physicalColumnIndex = firstColumnIndex + index;
            double width;
            switch (options.Mode)
            {
                case ExcelColumnWidthMode.Fixed:
                    width = options.FixedWidth.Value;
                    break;
                case ExcelColumnWidthMode.AutoFit:
                    sheet.AutoSizeColumn(physicalColumnIndex);
                    width = sheet.GetColumnWidth(physicalColumnIndex) / 256d;
                    break;
                case ExcelColumnWidthMode.Adaptive:
                    width = MeasureAdaptiveWidth(sheet, physicalColumnIndex, originRow, options.SampleRows);
                    break;
                default:
                    continue;
            }
            width = Math.Max(options.MinWidth ?? 0, width);
            width = Math.Min(options.MaxWidth ?? 255, width);
            sheet.SetColumnWidth(physicalColumnIndex, (int)Math.Min(255 * 256, Math.Max(0, width * 256)));
        }
    }

    /// <summary>
    /// 计算自适应列宽。
    /// </summary>
    private static double MeasureAdaptiveWidth(ISheet sheet, int columnIndex, int originRow, int sampleRows)
    {
        var width = 0d;
        var endRow = Math.Min(sheet.LastRowNum, originRow + sampleRows);
        for (var rowIndex = originRow; rowIndex <= endRow; rowIndex++)
        {
            var cell = sheet.GetRow(rowIndex)?.GetCell(columnIndex);
            if (cell == null)
                continue;
            var text = cell.ToString() ?? string.Empty;
            foreach (var line in text.Split(new[] { '\r', '\n' }, StringSplitOptions.None))
            {
                var lineWidth = 0d;
                foreach (var character in line)
                    lineWidth += character > 0xFF ? 2 : 1;
                width = Math.Max(width, lineWidth);
            }
        }
        return width + 2;
    }

    /// <summary>
    /// 应用实体的自动换行特性。
    /// </summary>
    private static void ApplyWrapText<T>(ISheet sheet) where T : class, new()
    {
        if (!typeof(T).IsDefined(typeof(WrapTextAttribute), false))
            return;
        for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
                continue;
            foreach (var cell in row.Cells)
                cell.CellStyle = cell.GetStyleWithWrapText();
        }
    }

    /// <summary>
    /// 合并标记为 Merge 的连续数据列。
    /// </summary>
    private static void MergeColumns<T>(ISheet sheet, IReadOnlyList<ExcelColumnPlan> columns,
        int dataRowStartIndex, int dataRowEndIndex, int firstColumnIndex) where T : class, new()
    {
        if (dataRowEndIndex - dataRowStartIndex < 1)
            return;
        for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
        {
            var column = columns[columnIndex];
            if (column.IsDynamic || !column.IsMerged)
                continue;
            var physicalColumnIndex = firstColumnIndex + columnIndex;
            var groupStart = dataRowStartIndex;
            var groupValue = sheet.GetRow(groupStart)?.GetCell(physicalColumnIndex).GetStringValue();
            for (var rowIndex = dataRowStartIndex + 1; rowIndex <= dataRowEndIndex + 1; rowIndex++)
            {
                var currentValue = rowIndex <= dataRowEndIndex
                    ? sheet.GetRow(rowIndex)?.GetCell(physicalColumnIndex).GetStringValue()
                    : null;
                if (rowIndex <= dataRowEndIndex && string.Equals(groupValue, currentValue, StringComparison.Ordinal))
                    continue;
                if (!string.IsNullOrWhiteSpace(groupValue) && rowIndex - groupStart > 1)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(groupStart, rowIndex - 1, physicalColumnIndex,
                        physicalColumnIndex));
                    var cell = sheet.GetRow(groupStart).GetCell(physicalColumnIndex);
                    cell.CellStyle = cell.GetStyleWithVerticalAlignment(VerticalAlignment.Center);
                }
                groupStart = rowIndex;
                groupValue = currentValue;
            }
        }
    }

    /// <summary>
    /// 转换并写入一个固定列或动态列的值。
    /// </summary>
    private static void WriteCell<T>(ICell cell, T item, ExcelColumnPlan column,
        IDictionary<string, object> dynamicValues, string sheetName, int rowIndex, int columnIndex,
        CultureInfo culture) where T : class, new()
    {
        object value;
        if (column.IsDynamic)
        {
            var values = dynamicValues ?? column.Getter(item) as IDictionary<string, object>;
            value = values != null && values.TryGetValue(column.Key, out var dynamicValue)
                ? dynamicValue
                : null;
        }
        else
            value = column.Getter(item);
        value = column.ConvertTo(value, sheetName, rowIndex, columnIndex, culture);
        column.WriteValue(cell, value);
    }
}

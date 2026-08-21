using Bing.Offices.Exports;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Mappings;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Resolvers;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using NPOI.SS.Util;
using NPOI.SS.UserModel.Charts;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Exports;

/// <summary>
/// 基于 NPOI 的单工作簿流式 Excel 导出器。
/// </summary>
internal sealed class NpoiExcelExporter : IExcelExporter
{
    /// <summary>
    /// 当前导出器使用的值转换器。
    /// </summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;

    /// <summary>
    /// 初始化一个<see cref="NpoiExcelExporter"/>类型的实例。
    /// </summary>
    /// <param name="valueConverters">值转换器集合。</param>
    public NpoiExcelExporter(IEnumerable<IExcelValueConverter> valueConverters = null) =>
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();

    /// <inheritdoc />
    public void Export(ExcelWorkbookExportRequest request, Stream destination,
        CancellationToken cancellationToken = default)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (destination == null)
            throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite)
            throw new ArgumentException("目标流不可写入。", nameof(destination));
        cancellationToken.ThrowIfCancellationRequested();

        try
        {
            using var workbook = CreateWorkbook(request);
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sheetRequest in request.Sheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ValidateSheetName(sheetRequest.Name);
                if (!names.Add(sheetRequest.Name))
                    throw new ArgumentException($"Workbook 包含重复 Sheet 名称: {sheetRequest.Name}");
                WriteSheet(workbook, sheetRequest, request.Template != null, cancellationToken);
            }

            using var bufferedDestination = new MemoryStream();
            workbook.Write(new NonDisposingStream(bufferedDestination), false);
            bufferedDestination.Position = 0;
            CopyTo(bufferedDestination, destination, cancellationToken);
        }
        finally
        {
            if (request.Template != null && !request.LeaveTemplateOpen)
                request.Template.Dispose();
        }
    }

    /// <summary>
    /// 创建普通或模板工作簿。模板加载后沿用同一 Sheet Writer。
    /// </summary>
    private static NPOI.SS.UserModel.IWorkbook CreateWorkbook(ExcelWorkbookExportRequest request)
    {
        if (request.Template == null)
            return ExcelHelper.PrepareWorkbook(request.Format);
        if (!request.Template.CanRead)
            throw new ArgumentException("模板流不可读取。", nameof(request));
        return NPOI.SS.UserModel.WorkbookFactory.Create(new NonDisposingStream(request.Template));
    }

    /// <summary>
    /// 验证 Excel Sheet 名称边界。
    /// </summary>
    private static void ValidateSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("工作表名称不能为空。", nameof(name));
        if (name.Length > 31)
            throw new ArgumentException("工作表名称不能超过 31 个字符。", nameof(name));
        if (name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0)
            throw new ArgumentException($"工作表名称包含非法字符: {name}", nameof(name));
    }

    /// <summary>
    /// 通过一次类型擦除调用执行具体 Sheet 的泛型计划。
    /// </summary>
    private void WriteSheet(NPOI.SS.UserModel.IWorkbook workbook, ExcelSheetExportRequest request,
        bool isTemplate, CancellationToken cancellationToken)
    {
        var method = GetType().GetMethod(nameof(WriteTypedSheet), BindingFlags.Instance | BindingFlags.NonPublic);
        try
        {
            method.MakeGenericMethod(request.ItemType).Invoke(this,
                new object[] { workbook, request, isTemplate, cancellationToken });
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }

    /// <summary>
    /// 执行一个泛型 Sheet 的统一列计划和 Cell Writer。
    /// </summary>
    private void WriteTypedSheet<T>(NPOI.SS.UserModel.IWorkbook workbook, ExcelSheetExportRequest request,
        bool isTemplate, CancellationToken cancellationToken) where T : class, new()
    {
        var sheet = workbook.GetSheet(request.Name);
        if (sheet == null && isTemplate)
            throw new InvalidOperationException($"模板缺少请求的 Sheet: {request.Name}");
        sheet ??= workbook.CreateSheet(request.Name);
        if (request.Hidden)
            workbook.SetSheetVisibility(workbook.GetSheetIndex(sheet), NPOI.SS.UserModel.SheetVisibility.Hidden);
        if (request.HeaderRowIndex < 0 || request.DataRowStartIndex <= request.HeaderRowIndex)
            throw new ArgumentOutOfRangeException(nameof(request.DataRowStartIndex));

        var templateOrigin = ResolveTemplateOrigin(workbook, sheet, request, isTemplate);
        var headerRowIndex = templateOrigin.Row + request.HeaderRowIndex;
        var firstColumnIndex = templateOrigin.Column;
        var map = ExcelTypeMapFactory.Get(request.MappingProfile as Configurations.ExcelMappingProfile<T>,
            request.MappingConfiguration);
        ValidateDynamicDefinitions(request.DynamicColumns);
        var columns = CreateColumns(map, request.DynamicColumns);
        ValidateDynamicColumns(request.DynamicColumns, columns);
        var header = sheet.GetRow(headerRowIndex) ?? sheet.CreateRow(headerRowIndex);
        WriteCustomHeaders(sheet, request.HeaderRows, templateOrigin.Row, firstColumnIndex,
            request.CommentConflictPolicy);
        for (var index = 0; index < columns.Count; index++)
            header.CreateCell(firstColumnIndex + index).SetCellValue(columns[index].Title);

        var rowIndex = templateOrigin.Row + request.DataRowStartIndex;
        foreach (var item in request.Data.Cast<T>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var dynamicValues = request.DynamicGetter?.Invoke(item);
            ValidateUnknownDynamicValues(request, dynamicValues, columns);
            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                var physicalColumnIndex = firstColumnIndex + columnIndex;
                var cell = row.GetCell(physicalColumnIndex) ?? row.CreateCell(physicalColumnIndex);
                WriteRequestCell(cell, item, columns[columnIndex], request, dynamicValues, sheet.SheetName,
                    rowIndex + 1, physicalColumnIndex + 1);
            }
            rowIndex++;
        }

        ApplyRequestStyles(workbook, sheet, header, columns, request, firstColumnIndex,
            templateOrigin.Row + request.DataRowStartIndex, rowIndex);
        ApplyHeaderStyle<T>(workbook, sheet, headerRowIndex);
        ApplyWrapText<T>(sheet);
        ApplyColumnWidths(sheet, columns, firstColumnIndex, templateOrigin.Row, request.ColumnWidth);
        MergeColumns<T>(sheet, columns, templateOrigin.Row + request.DataRowStartIndex, rowIndex - 1,
            firstColumnIndex);
        CreateCharts(workbook, sheet, columns, request.Charts, templateOrigin.Row + request.DataRowStartIndex,
            rowIndex, firstColumnIndex);
    }

    /// <summary>
    /// 将 provider-neutral 图表定义转换为 XSSF 图表。
    /// </summary>
    private static void CreateCharts(NPOI.SS.UserModel.IWorkbook workbook, NPOI.SS.UserModel.ISheet sheet,
        IReadOnlyList<ExcelColumnPlan> columns, IReadOnlyList<ExcelChartDefinition> charts, int dataStartRow,
        int dataEndRow, int firstColumnIndex)
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

    private static CellRangeAddress CreateChartRange(NPOI.SS.UserModel.ISheet sheet, ExcelChartRange range,
        int columnIndex, int defaultStartRow, int defaultEndRow)
    {
        var startRow = range.StartRow ?? defaultStartRow;
        var endRow = range.EndRow ?? defaultEndRow;
        if (startRow < 0 || endRow <= startRow || endRow > sheet.LastRowNum + 1)
            throw new ArgumentException($"图表范围行索引无效: {range.ColumnKey}", nameof(range));
        return new CellRangeAddress(startRow, endRow - 1, columnIndex, columnIndex);
    }

    /// <summary>
    /// 解析模板命名区域的起点；普通导出从零坐标开始。
    /// </summary>
    private static (int Row, int Column) ResolveTemplateOrigin(NPOI.SS.UserModel.IWorkbook workbook,
        NPOI.SS.UserModel.ISheet sheet, ExcelSheetExportRequest request, bool isTemplate)
    {
        if (!isTemplate || string.IsNullOrWhiteSpace(request.TemplateRegion))
            return (0, 0);
        var name = workbook.GetName(request.TemplateRegion);
        if (name == null || string.IsNullOrWhiteSpace(name.RefersToFormula))
            throw new InvalidOperationException($"模板缺少命名区域: {request.TemplateRegion}");
        var formula = name.RefersToFormula.TrimStart('=');
        var separator = formula.LastIndexOf('!');
        if (separator < 0)
            throw new InvalidOperationException($"模板命名区域缺少 Sheet 引用: {request.TemplateRegion}");
        var sheetName = formula.Substring(0, separator).Trim('\'', ' ');
        if (!string.Equals(sheetName, sheet.SheetName, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException($"模板命名区域不属于请求 Sheet: {request.TemplateRegion}");
        var address = formula.Substring(separator + 1).Split(':')[0].Replace("$", string.Empty);
        var match = Regex.Match(address, "^([A-Za-z]+)([0-9]+)$");
        if (!match.Success)
            throw new InvalidOperationException($"模板命名区域地址无效: {request.TemplateRegion}");
        var column = 0;
        foreach (var character in match.Groups[1].Value.ToUpperInvariant())
            column = column * 26 + character - 'A' + 1;
        return (int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) - 1, column - 1);
    }

    /// <summary>
    /// 检查动态列定义和列键唯一性。
    /// </summary>
    private static void ValidateDynamicColumns(IReadOnlyList<ExcelDynamicColumnDefinition> definitions,
        IReadOnlyList<ExcelColumnPlan> columns)
    {
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var titles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in columns.Where(column => !column.IsDynamic))
            titles.Add(column.Title);
        foreach (var definition in definitions ?? Array.Empty<ExcelDynamicColumnDefinition>())
        {
            if (!keys.Add(definition.Key))
                throw new ArgumentException($"动态列包含重复 Key: {definition.Key}", nameof(definitions));
            if (!titles.Add(definition.Title))
                throw new ArgumentException($"动态列包含重复标题: {definition.Title}", nameof(definitions));
            foreach (var alias in definition.Aliases ?? Array.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(alias) || !titles.Add(alias))
                    throw new ArgumentException($"动态列包含重复或空标题别名: {alias}", nameof(definitions));
            }
            if (definition.PhysicalColumnIndex.HasValue && definition.Placement != null)
                throw new ArgumentException($"动态列 {definition.Key} 不能同时指定相对位置和物理索引。",
                    nameof(definitions));
            if (definition.Placement?.PhysicalColumnIndex != null && definition.PhysicalColumnIndex.HasValue)
                throw new ArgumentException($"动态列 {definition.Key} 不能重复指定物理索引。",
                    nameof(definitions));
        }
        ValidateColumns(columns);
    }

    private static void ValidateDynamicDefinitions(IReadOnlyList<ExcelDynamicColumnDefinition> definitions)
    {
        foreach (var definition in definitions ?? Array.Empty<ExcelDynamicColumnDefinition>())
        {
            if (definition == null || string.IsNullOrWhiteSpace(definition.Key)
                || string.IsNullOrWhiteSpace(definition.Title))
                throw new ArgumentException("动态列 Key 和 Title 不能为空。", nameof(definitions));
        }
    }

    /// <summary>
    /// 拒绝动态字典中未声明的值。
    /// </summary>
    private static void ValidateUnknownDynamicValues(ExcelSheetExportRequest request,
        IDictionary<string, object> values, IReadOnlyList<ExcelColumnPlan> columns)
    {
        if (!request.FailOnUnknownDynamicValues || values == null)
            return;
        var keys = new HashSet<string>(columns.Where(column => column.IsDynamic).Select(column => column.Key),
            StringComparer.Ordinal);
        var unknown = values.Keys.FirstOrDefault(key => !keys.Contains(key));
        if (unknown != null)
            throw new InvalidOperationException($"动态值未声明: {unknown}");
    }

    /// <summary>
    /// 新 Workbook 请求使用统一的 Cell Writer 写入动态和固定列。
    /// </summary>
    private void WriteRequestCell<T>(NPOI.SS.UserModel.ICell cell, T item, ExcelColumnPlan column,
        ExcelSheetExportRequest request, IDictionary<string, object> dynamicValues, string sheetName,
        int rowIndex, int columnIndex) where T : class, new()
    {
        WriteCell(cell, item, column, dynamicValues, sheetName, rowIndex, columnIndex, request.Culture);
    }

    /// <summary>
    /// 应用 Workbook 请求级区域和动态列样式。
    /// </summary>
    private static void ApplyRequestStyles(NPOI.SS.UserModel.IWorkbook workbook,
        NPOI.SS.UserModel.ISheet sheet, NPOI.SS.UserModel.IRow header, IReadOnlyList<ExcelColumnPlan> columns,
        ExcelSheetExportRequest request, int firstColumnIndex, int firstDataRowIndex, int lastRowIndex)
    {
        if (request.HeaderStyle == null && request.BodyStyle == null && request.SheetStyle == null
            && columns.All(column => column.HeaderStyle == null && column.BodyStyle == null))
            return;
        for (var index = 0; index < columns.Count; index++)
        {
            var style = columns[index].HeaderStyle ?? request.HeaderStyle ?? request.SheetStyle;
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
                var style = columns[index].BodyStyle ?? request.BodyStyle ?? request.SheetStyle;
                if (style != null)
                {
                    var cell = row.GetCell(firstColumnIndex + index);
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, style);
                }
            }
        }
    }

    /// <summary>
    /// 将已完成的工作簿缓冲写入调用方目标流，并在每个数据块之间检查取消状态。
    /// </summary>
    /// <param name="source">实现拥有的已完成工作簿缓冲。</param>
    /// <param name="destination">调用方拥有的目标流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void CopyTo(Stream source, Stream destination, CancellationToken cancellationToken)
    {
        var buffer = new byte[81920];
        int count;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            count = source.Read(buffer, 0, buffer.Length);
            if (count == 0)
                break;
            cancellationToken.ThrowIfCancellationRequested();
            destination.Write(buffer, 0, count);
        }
    }

    /// <summary>
    /// 在当前工作簿中写入自定义表头布局。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="headerRows">自定义表头行。</param>
    /// <param name="originRow">模板区域的行起点。</param>
    /// <param name="originColumn">模板区域的列起点。</param>
    private static void WriteCustomHeaders(NPOI.SS.UserModel.ISheet sheet, IReadOnlyList<ExcelHeaderRow> headerRows,
        int originRow, int originColumn, ExcelCommentConflictPolicy commentConflictPolicy)
    {
        foreach (var headerRow in headerRows ?? Array.Empty<ExcelHeaderRow>())
        {
            var rowIndex = originRow + headerRow.RowIndex;
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            foreach (var headerCell in headerRow.Cells)
            {
                var columnIndex = originColumn + headerCell.ColumnIndex;
                var cell = row.CreateCell(columnIndex);
                cell.SetValue(headerCell.Value);
                if (headerCell.Comment != null)
                    ApplyComment(sheet, cell, headerCell.Comment, commentConflictPolicy);
                if (headerCell.RowSpan > 1 || headerCell.ColumnSpan > 1)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex + headerCell.RowSpan - 1,
                        columnIndex, columnIndex + headerCell.ColumnSpan - 1));
                }
            }
        }
    }

    private static void ApplyComment(NPOI.SS.UserModel.ISheet sheet, NPOI.SS.UserModel.ICell cell,
        ExcelComment comment, ExcelCommentConflictPolicy conflictPolicy)
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

    private static void ApplyHeaderStyle<T>(NPOI.SS.UserModel.IWorkbook workbook,
        NPOI.SS.UserModel.ISheet sheet, int headerRowIndex) where T : class, new()
    {
        var attribute = typeof(T).GetCustomAttributes(typeof(HeaderAttribute), false).Cast<HeaderAttribute>()
            .SingleOrDefault();
        if (attribute == null)
            return;
        var row = sheet.GetRow(headerRowIndex);
        if (row == null)
            return;
        foreach (var cell in row.Cells)
        {
            var style = workbook.CreateCellStyle();
            style.CloneStyleFrom(cell.CellStyle);
            var font = workbook.CreateFont();
            font.CloneStyleFrom(workbook.GetFontAt(cell.CellStyle.FontIndex));
            font.FontName = attribute.FontName;
            font.Color = ColorResolver.Resolve(attribute.Color);
            font.FontHeightInPoints = attribute.FontSize;
            font.IsBold = attribute.Bold;
            style.SetFont(font);
            cell.CellStyle = style;
        }
    }

    private static void ApplyColumnWidths(NPOI.SS.UserModel.ISheet sheet, IReadOnlyList<ExcelColumnPlan> columns,
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

    private static double MeasureAdaptiveWidth(NPOI.SS.UserModel.ISheet sheet, int columnIndex, int originRow,
        int sampleRows)
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

    private static void ApplyWrapText<T>(NPOI.SS.UserModel.ISheet sheet) where T : class, new()
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

    private static void MergeColumns<T>(NPOI.SS.UserModel.ISheet sheet, IReadOnlyList<ExcelColumnPlan> columns,
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
                    cell.CellStyle = cell.GetStyleWithVerticalAlignment(NPOI.SS.UserModel.VerticalAlignment.Center);
                }
                groupStart = rowIndex;
                groupValue = currentValue;
            }
        }
    }

    /// <summary>
    /// 创建 Workbook 请求使用的固定列和 typed 动态列计划。
    /// </summary>
    private IReadOnlyList<ExcelColumnPlan> CreateColumns<T>(ExcelTypeMap<T> typeMap,
        IReadOnlyList<ExcelDynamicColumnDefinition> dynamicColumns)
    {
        var fixedColumns = new List<ExcelColumnPlan>();
        foreach (var property in typeMap.Properties)
        {
            if (property.Ignored)
                continue;
            if (!property.IsDynamicColumn)
            {
                fixedColumns.Add(new ExcelColumnPlan(property.Title, property, false, -1, null, null,
                    BindValueConverters(property.ConverterName, property.Property.PropertyType)));
                continue;
            }
        }
        var columns = fixedColumns.ToList();
        var dynamicProperty = typeMap.Properties.FirstOrDefault(property => property.IsDynamicColumn);
        if (dynamicProperty == null)
            return columns;
        var definitions = (dynamicColumns ?? Array.Empty<ExcelDynamicColumnDefinition>())
            .OrderBy(definition => definition.Order)
            .ThenBy(definition => definition.Key, StringComparer.Ordinal)
            .ToList();
        foreach (var definition in definitions)
        {
            var column = new ExcelColumnPlan(definition.Title, dynamicProperty, true, -1, definition, definition.Key,
                BindValueConverters(definition.ConverterName ?? dynamicProperty.ConverterName, definition.DataType));
            var placement = definition.Placement;
            var physicalIndex = definition.PhysicalColumnIndex ?? placement?.PhysicalColumnIndex;
            if (physicalIndex.HasValue)
            {
                if (physicalIndex.Value > columns.Count)
                    throw new ArgumentOutOfRangeException(nameof(definition.PhysicalColumnIndex),
                        $"动态列 {definition.Key} 的物理索引超出当前列计划。");
                columns.Insert(physicalIndex.Value, column);
                continue;
            }
            if (placement != null && !string.IsNullOrWhiteSpace(placement.BeforeKey))
            {
                var index = columns.FindIndex(item => string.Equals(item.Key, placement.BeforeKey,
                    StringComparison.OrdinalIgnoreCase));
                if (index < 0)
                    throw new ArgumentException($"动态列 {definition.Key} 的 Before 目标不存在: {placement.BeforeKey}");
                columns.Insert(index, column);
                continue;
            }
            if (placement != null && !string.IsNullOrWhiteSpace(placement.AfterKey))
            {
                var index = columns.FindIndex(item => string.Equals(item.Key, placement.AfterKey,
                    StringComparison.OrdinalIgnoreCase));
                if (index < 0)
                    throw new ArgumentException($"动态列 {definition.Key} 的 After 目标不存在: {placement.AfterKey}");
                columns.Insert(index + 1, column);
                continue;
            }
            columns.Add(column);
        }
        return columns;
    }

    /// <summary>
    /// 验证解析后的列标题唯一，确保导出结果可被导入器无歧义读取。
    /// </summary>
    /// <param name="columns">当前请求的导出列。</param>
    private static void ValidateColumns(IReadOnlyList<ExcelColumnPlan> columns)
    {
        var titles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in columns)
        {
            if (!titles.Add(column.Title))
                throw new ArgumentException($"导出列标题重复: {column.Title}", nameof(columns));
        }
    }

    /// <summary>
    /// 写入一个已解析列的单元格值。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="cell">目标单元格。</param>
    /// <param name="item">当前实体。</param>
    /// <param name="column">导出列定义。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowIndex">从一开始的行索引。</param>
    /// <param name="columnIndex">从一开始的列索引。</param>
    /// <param name="culture">值转换使用的区域性。</param>
    private void WriteCell<T>(NPOI.SS.UserModel.ICell cell, T item, ExcelColumnPlan column,
        IDictionary<string, object> dynamicValues, string sheetName, int rowIndex, int columnIndex,
        CultureInfo culture) where T : class, new()
    {
        object value;
        if (column.IsDynamic)
        {
            var values = dynamicValues ?? column.Getter(item) as IDictionary<string, object>;
            values ??= new Dictionary<string, object>();
            values.TryGetValue(column.Key, out value);
        }
        else
            value = column.Getter(item);

        value = column.ConvertTo(value, sheetName, rowIndex, columnIndex, culture);
        column.WriteValue(cell, value);
    }

    /// <summary>
    /// 解析属性允许使用的值转换器。
    /// </summary>
    /// <param name="property">属性映射。</param>
    private IReadOnlyList<IExcelValueConverter> BindValueConverters(string converterName, Type propertyType)
    {
        if (string.IsNullOrWhiteSpace(converterName))
            return _valueConverters.Where(converter => converter.CanConvert(propertyType)).ToArray();
        var converters = _valueConverters.OfType<INamedExcelValueConverter>().Where(converter =>
            string.Equals(converter.Name, converterName, StringComparison.OrdinalIgnoreCase)).ToList();
        if (converters.Count != 1)
            throw new InvalidOperationException($"未找到唯一命名值转换器: {converterName}");
        if (!converters[0].CanConvert(propertyType))
            throw new InvalidOperationException($"值转换器 {converterName} 不支持属性类型: {propertyType.FullName}");
        return converters;
    }

    /// <summary>
    /// 让 NPOI 可以释放包装器但不能关闭调用方拥有的内部缓冲流。
    /// </summary>
    private sealed class NonDisposingStream : Stream
    {
        private readonly Stream _inner;

        public NonDisposingStream(Stream inner) => _inner = inner;

        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => _inner.CanSeek;
        public override bool CanWrite => _inner.CanWrite;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        public override void SetLength(long value) => _inner.SetLength(value);
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Flush();
            base.Dispose(disposing);
        }
    }

}

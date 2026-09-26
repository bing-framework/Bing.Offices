using ClosedXML.Excel;
using Bing.Offices.Exceptions;

using Bing.Offices.Internals;

namespace Bing.Offices.Exports;

/// <summary>
/// 将公共报表定义映射到 ClosedXML 工作簿。
/// </summary>
internal static class ClosedXmlReportWriter
{
    /// <summary>
    /// 在工作簿创建和数据枚举前校验报表定义。
    /// </summary>
    /// <param name="request">当前操作请求。</param>
    internal static void Validate(ExcelWorkbookExportRequest request)
    {
        ExcelReportPreflight.Validate(request, "ClosedXML", 1048576, 16384,
            supportsFreezeOrigin: false);
        var invalid = request.Sheets.Select(sheet => sheet.PrintLayout?.ScalePercent)
            .FirstOrDefault(scale => scale.HasValue && (scale.Value < 10 || scale.Value > 400));
        if (invalid.HasValue)
            throw new BingOfficesConfigurationException("打印缩放百分比必须在 10 到 400 之间。",
                stage: BingOfficesStage.Preflight);
    }

    /// <summary>
    /// 应用当前 Sheet 的报表定义。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="request">当前操作请求。</param>
    internal static void Apply(XLWorkbook workbook, IXLWorksheet worksheet, ExcelSheetExportRequest request)
    {
        foreach (var tableDefinition in request.Tables)
        {
            var range = ToRange(worksheet, tableDefinition.Range);
            var table = range.CreateTable(tableDefinition.Name);
            table.SetShowHeaderRow(tableDefinition.HasHeaders);
            table.SetShowTotalsRow(tableDefinition.ShowTotals);
            if (!string.IsNullOrWhiteSpace(tableDefinition.StyleName))
            {
                var theme = typeof(XLTableTheme).GetField(tableDefinition.StyleName,
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static
                    | System.Reflection.BindingFlags.IgnoreCase)
                    ?.GetValue(null) as XLTableTheme;
                if (theme != null)
                    table.Theme = theme;
            }
        }

        foreach (var filter in request.AutoFilters)
        {
            // ClosedXML tables own their AutoFilter range; applying the same range a second
            // time raises an overlap exception, so the explicit definition is idempotent.
            if (request.Tables.Any(table => ExcelReportPreflight.SameRange(table.Range, filter.Range)))
                continue;
            ToRange(worksheet, filter.Range).SetAutoFilter();
        }

        if (request.FreezePane != null)
        {
            if (request.FreezePane.TopRow.HasValue || request.FreezePane.LeftColumn.HasValue)
                throw Unsupported("ClosedXML 当前版本仅支持按行数和列数冻结窗格。", BingOfficesStage.Preflight);
            worksheet.SheetView.FreezeRows(request.FreezePane.Rows);
            worksheet.SheetView.FreezeColumns(request.FreezePane.Columns);
        }

        foreach (var conditional in request.ConditionalFormats)
            ApplyConditionalFormat(worksheet, conditional);

        foreach (var namedRange in request.NamedRanges)
        {
            var range = worksheet.Range(namedRange.Address);
            if (string.IsNullOrWhiteSpace(namedRange.SheetName))
                workbook.DefinedNames.Add(namedRange.Name, range);
            else if (string.Equals(namedRange.SheetName, worksheet.Name, StringComparison.OrdinalIgnoreCase))
                worksheet.DefinedNames.Add(namedRange.Name, range);
            else
                throw Unsupported("ClosedXML 名称范围的 Sheet 引用必须与当前 Sheet 一致。", BingOfficesStage.Preflight);
        }

        if (request.PrintLayout != null)
            ApplyPrintLayout(worksheet, request.PrintLayout);
    }

    /// <summary>
    /// 将公共区域定义转换为工作表区域。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="range">待处理的单元格区域。</param>
    /// <returns>对应的工作表区域。</returns>
    private static IXLRange ToRange(IXLWorksheet worksheet, ExcelRangeDefinition range)
    {
        range.Validate(nameof(range));
        return worksheet.Range(range.StartRow + 1, range.StartColumn + 1,
            range.EndRow + 1, range.EndColumn + 1);
    }

    /// <summary>
    /// 将条件格式定义应用到工作表。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="definition">待应用的定义。</param>
    private static void ApplyConditionalFormat(IXLWorksheet worksheet,
        ExcelConditionalFormatDefinition definition)
    {
        var conditional = ToRange(worksheet, definition.Range).AddConditionalFormat();
        IXLStyle style;
        switch (definition.Type)
        {
            case ExcelConditionalFormatType.CellValue:
                style = definition.Operator switch
                {
                    ExcelConditionalComparisonOperator.Equal => conditional.WhenEquals(definition.Formula1),
                    ExcelConditionalComparisonOperator.NotEqual => conditional.WhenNotEquals(definition.Formula1),
                    ExcelConditionalComparisonOperator.GreaterThan => conditional.WhenGreaterThan(definition.Formula1),
                    ExcelConditionalComparisonOperator.GreaterThanOrEqual => conditional.WhenEqualOrGreaterThan(definition.Formula1),
                    ExcelConditionalComparisonOperator.LessThan => conditional.WhenLessThan(definition.Formula1),
                    ExcelConditionalComparisonOperator.LessThanOrEqual => conditional.WhenEqualOrLessThan(definition.Formula1),
                    ExcelConditionalComparisonOperator.Between => conditional.WhenBetween(definition.Formula1, definition.Formula2),
                    ExcelConditionalComparisonOperator.NotBetween => conditional.WhenNotBetween(definition.Formula1, definition.Formula2),
                    _ => throw new ArgumentOutOfRangeException(nameof(definition.Operator))
                };
                ApplyColors(style, definition);
                return;
            case ExcelConditionalFormatType.Formula:
                style = conditional.WhenIsTrue(definition.Formula1);
                ApplyColors(style, definition);
                return;
            case ExcelConditionalFormatType.ColorScale:
                conditional.ColorScale().LowestValue(ToColor(ExcelReportPreflight.ScaleMinimumColor(definition)))
                    .HighestValue(ToColor(ExcelReportPreflight.ScaleMaximumColor(definition)));
                return;
            case ExcelConditionalFormatType.DataBar:
                conditional.DataBar(ToColor(ExcelReportPreflight.DataBarColor(definition)), true).LowestValue().HighestValue();
                return;
            case ExcelConditionalFormatType.IconSet:
                conditional.IconSet(XLIconSetStyle.ThreeTrafficLights1, false, false)
                    .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 0, XLCFContentType.Percent)
                    .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 33, XLCFContentType.Percent)
                    .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 67, XLCFContentType.Percent);
                return;
            default:
                throw new ArgumentOutOfRangeException(nameof(definition.Type));
        }
    }

    /// <summary>
    /// 应用条件格式的前景色和背景色。
    /// </summary>
    /// <param name="style">条件格式样式。</param>
    /// <param name="definition">待应用的定义。</param>
    private static void ApplyColors(IXLStyle style, ExcelConditionalFormatDefinition definition)
    {
        if (!string.IsNullOrWhiteSpace(definition.ForegroundColor))
            style.Font.FontColor = ToColor(definition.ForegroundColor);
        if (!string.IsNullOrWhiteSpace(definition.BackgroundColor))
            style.Fill.BackgroundColor = ToColor(definition.BackgroundColor);
    }

    /// <summary>
    /// 将颜色文本转换为 ClosedXML 颜色。
    /// </summary>
    /// <param name="value">待转换的值。</param>
    /// <returns>解析后的颜色；空文本返回无颜色标记。</returns>
    private static XLColor ToColor(string value) => string.IsNullOrWhiteSpace(value)
        ? XLColor.NoColor : XLColor.FromHtml(value);

    /// <summary>
    /// 应用工作表打印布局。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="options">工作表打印布局选项。</param>
    private static void ApplyPrintLayout(IXLWorksheet worksheet, ExcelPrintLayoutOptions options)
    {
        var setup = worksheet.PageSetup;
        setup.PageOrientation = options.Orientation == ExcelPrintOrientation.Landscape
            ? XLPageOrientation.Landscape : XLPageOrientation.Portrait;
        setup.PaperSize = options.PaperSize switch
        {
            ExcelPrintPaperSize.Letter => XLPaperSize.LetterPaper,
            ExcelPrintPaperSize.A3 => XLPaperSize.A3Paper,
            ExcelPrintPaperSize.Legal => XLPaperSize.LegalPaper,
            _ => XLPaperSize.A4Paper
        };
        setup.Margins.Top = options.TopMargin;
        setup.Margins.Bottom = options.BottomMargin;
        setup.Margins.Left = options.LeftMargin;
        setup.Margins.Right = options.RightMargin;
        if (options.ScalePercent.HasValue)
            setup.Scale = options.ScalePercent.Value;
        if (options.FitToWidth.HasValue || options.FitToHeight.HasValue)
            setup.FitToPages(options.FitToWidth ?? 0, options.FitToHeight ?? 0);
        if (!string.IsNullOrWhiteSpace(options.PrintArea))
            setup.PrintAreas.Add(options.PrintArea);
        if (!string.IsNullOrWhiteSpace(options.RepeatRows))
            setup.SetRowsToRepeatAtTop(NormalizeRepeatRange(options.RepeatRows));
        if (!string.IsNullOrWhiteSpace(options.RepeatColumns))
            setup.SetColumnsToRepeatAtLeft(NormalizeRepeatRange(options.RepeatColumns));
        if (!string.IsNullOrWhiteSpace(options.Header))
            setup.Header.Center.AddText(options.Header, XLHFOccurrence.AllPages);
        if (!string.IsNullOrWhiteSpace(options.Footer))
            setup.Footer.Center.AddText(options.Footer, XLHFOccurrence.AllPages);
    }

    /// <summary>
    /// 移除重复打印区域地址中的绝对引用标记。
    /// </summary>
    /// <param name="value">待转换的值。</param>
    /// <returns>移除绝对引用标记后的区域地址。</returns>
    private static string NormalizeRepeatRange(string value) => value.Replace("$", string.Empty);

    /// <summary>
    /// 创建不支持当前功能的异常。
    /// </summary>
    /// <param name="message">错误消息。</param>
    /// <param name="stage">发生错误的处理阶段。</param>
    /// <returns>包含当前处理阶段的功能不支持异常。</returns>
    private static BingOfficesUnsupportedFeatureException Unsupported(string message, BingOfficesStage stage) =>
        new(message, provider: "ClosedXML", operation: BingOfficesOperation.Export, stage: stage);
}

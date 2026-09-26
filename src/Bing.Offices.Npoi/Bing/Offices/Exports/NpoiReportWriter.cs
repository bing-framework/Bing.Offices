using NPOI.SS.UserModel;
using NPOI.SS;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Bing.Offices.Exceptions;
using Bing.Offices.Internals;

namespace Bing.Offices.Exports;

/// <summary>
/// 将公共报表定义映射到 NPOI 工作簿。
/// </summary>
internal static class NpoiReportWriter
{
    /// <summary>
    /// 在工作簿创建和数据枚举前校验报表定义。
    /// </summary>
    /// <param name="request">当前操作请求。</param>
    internal static void Validate(ExcelWorkbookExportRequest request)
    {
        var isXls = request.Format == ExcelFormat.Xls;
        var maxRows = isXls
            ? SpreadsheetVersion.EXCEL97.LastRowIndex + 1
            : SpreadsheetVersion.EXCEL2007.LastRowIndex + 1;
        ExcelReportPreflight.Validate(request, "NPOI",
            maxRows,
            isXls ? SpreadsheetVersion.EXCEL97.LastColumnIndex + 1 : SpreadsheetVersion.EXCEL2007.LastColumnIndex + 1,
            supportsFreezeOrigin: true);
        ValidatePrintScale(request);
        if (isXls && request.Sheets.Any(sheet => sheet.Tables.Count != 0))
            throw Unsupported("NPOI 的 XLS/HSSF 分支不支持表格定义。",
                BingOfficesStage.Preflight);
        if (isXls && request.Sheets.SelectMany(sheet => sheet.ConditionalFormats).Any(definition =>
                definition.Type == ExcelConditionalFormatType.ColorScale
                || definition.Type == ExcelConditionalFormatType.DataBar
                || definition.Type == ExcelConditionalFormatType.IconSet))
            throw Unsupported("NPOI 的 XLS/HSSF 分支不支持 ColorScale、DataBar 或 IconSet。",
                BingOfficesStage.Preflight);
    }

    /// <summary>
    /// 应用当前 Sheet 的报表定义。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="request">当前操作请求。</param>
    internal static void Apply(IWorkbook workbook, ISheet sheet, ExcelSheetExportRequest request)
    {
        var tableFilterRanges = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (request.Tables.Count > 0)
        {
            if (sheet is not XSSFSheet xssfSheet)
                throw Unsupported("NPOI 的表格定义仅支持 XSSF/XLSX 工作簿。", BingOfficesStage.Preflight);
            foreach (var definition in request.Tables)
            {
                var table = xssfSheet.CreateTable();
                table.Name = definition.Name;
                table.DisplayName = definition.Name;
                var filterAddress = ToAddress(definition.Range);
                var tableAddress = ToAddress(definition.Range, definition.ShowTotals ? 1 : 0);
                if (definition.ShowTotals && sheet.GetRow(definition.Range.EndRow + 1) == null)
                    sheet.CreateRow(definition.Range.EndRow + 1);
                table.CellReferences = new AreaReference(tableAddress, SpreadsheetVersion.EXCEL2007);
                table.UpdateReferences();
                table.UpdateHeaders();
                if (!string.IsNullOrWhiteSpace(definition.StyleName))
                    table.StyleName = definition.StyleName;
                if (!definition.HasHeaders)
                    throw Unsupported("NPOI 表格必须包含表头。", BingOfficesStage.Preflight);
                table.IsHasTotalsRow = definition.ShowTotals;
                table.GetCTTable().totalsRowCount = definition.ShowTotals ? 1u : 0u;
                table.GetCTTable().autoFilter = new CT_AutoFilter { @ref = filterAddress };
                tableFilterRanges.Add(filterAddress);
            }
        }

        foreach (var filter in request.AutoFilters)
        {
            if (tableFilterRanges.Contains(ToAddress(filter.Range)))
                continue;
            sheet.SetAutoFilter(ToCellRange(filter.Range));
        }

        if (request.FreezePane != null)
            sheet.CreateFreezePane(request.FreezePane.Columns, request.FreezePane.Rows,
                request.FreezePane.LeftColumn ?? request.FreezePane.Columns,
                request.FreezePane.TopRow ?? request.FreezePane.Rows);

        foreach (var conditional in request.ConditionalFormats)
            ApplyConditionalFormatting(sheet, conditional);

        foreach (var namedRange in request.NamedRanges)
        {
            var name = workbook.CreateName();
            if (!string.IsNullOrWhiteSpace(namedRange.SheetName))
                name.SheetIndex = workbook.GetSheetIndex(namedRange.SheetName);
            name.NameName = namedRange.Name;
            name.RefersToFormula = ToFormula(sheet, namedRange);
        }

        if (request.PrintLayout != null)
            ApplyPrintLayout(sheet, request.PrintLayout);
    }

    /// <summary>
    /// 将条件格式定义应用到工作表。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="definition">待应用的定义。</param>
    private static void ApplyConditionalFormatting(ISheet sheet,
        ExcelConditionalFormatDefinition definition)
    {
        var formatting = sheet.SheetConditionalFormatting;
        IConditionalFormattingRule rule;
        switch (definition.Type)
        {
            case ExcelConditionalFormatType.ColorScale:
                rule = formatting.CreateConditionalFormattingColorScaleRule();
                var scale = rule.ColorScaleFormatting;
                var minimum = scale.CreateThreshold();
                minimum.RangeType = RangeType.MIN;
                var maximum = scale.CreateThreshold();
                maximum.RangeType = RangeType.MAX;
                scale.Thresholds = new[] { minimum, maximum };
                scale.Colors = new ExtendedColor[]
                {
                    new XSSFColor(ParseColor(ExcelReportPreflight.ScaleMinimumColor(definition)), null),
                    new XSSFColor(ParseColor(ExcelReportPreflight.ScaleMaximumColor(definition)), null)
                };
                formatting.AddConditionalFormatting(new[] { ToCellRange(definition.Range) }, rule);
                return;
            case ExcelConditionalFormatType.DataBar:
                rule = formatting.CreateConditionalFormattingRule(
                    new XSSFColor(ParseColor(ExcelReportPreflight.DataBarColor(definition)), null));
                rule.DataBarFormatting.MinThreshold.RangeType = RangeType.MIN;
                rule.DataBarFormatting.MaxThreshold.RangeType = RangeType.MAX;
                formatting.AddConditionalFormatting(new[] { ToCellRange(definition.Range) }, rule);
                return;
            case ExcelConditionalFormatType.IconSet:
                rule = formatting.CreateConditionalFormattingRule(IconSet.GYR_3_TRAFFIC_LIGHTS);
                var thresholds = rule.MultiStateFormatting.Thresholds;
                thresholds[0].RangeType = RangeType.PERCENT;
                thresholds[0].Value = 0d;
                thresholds[1].RangeType = RangeType.PERCENT;
                thresholds[1].Value = 33d;
                thresholds[2].RangeType = RangeType.PERCENT;
                thresholds[2].Value = 67d;
                formatting.AddConditionalFormatting(new[] { ToCellRange(definition.Range) }, rule);
                return;
            case ExcelConditionalFormatType.Formula:
                rule = formatting.CreateConditionalFormattingRule(definition.Formula1);
                break;
            case ExcelConditionalFormatType.CellValue:
                rule = formatting.CreateConditionalFormattingRule(ToComparison(definition.Operator),
                    definition.Formula1, definition.Formula2);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(definition.Type));
        }
        var font = rule.CreateFontFormatting();
        if (!string.IsNullOrWhiteSpace(definition.ForegroundColor))
            font.FontColor = ToNpoiColor(sheet.Workbook, definition.ForegroundColor);
        var pattern = rule.CreatePatternFormatting();
        if (!string.IsNullOrWhiteSpace(definition.BackgroundColor))
        {
            pattern.FillPattern = FillPattern.SolidForeground;
            pattern.FillForegroundColorColor = ToNpoiColor(sheet.Workbook, definition.BackgroundColor);
        }
        formatting.AddConditionalFormatting(new[] { ToCellRange(definition.Range) }, rule);
    }

    /// <summary>
    /// 校验打印缩放比例。
    /// </summary>
    /// <param name="request">当前操作请求。</param>
    private static void ValidatePrintScale(ExcelWorkbookExportRequest request)
    {
        var invalid = request.Sheets.Select(sheet => sheet.PrintLayout?.ScalePercent)
            .FirstOrDefault(scale => scale.HasValue && (scale.Value < 10 || scale.Value > 400));
        if (invalid.HasValue)
            throw new BingOfficesConfigurationException("打印缩放百分比必须在 10 到 400 之间。",
                stage: BingOfficesStage.Preflight);
    }

    /// <summary>
    /// 将颜色文本转换为 NPOI 颜色。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="value">待转换的值。</param>
    /// <returns>适用于当前工作簿格式的颜色对象。</returns>
    private static IColor ToNpoiColor(IWorkbook workbook, string value)
    {
        var rgb = ParseColor(value);
        if (workbook is HSSFWorkbook hssfWorkbook)
            return hssfWorkbook.GetCustomPalette().FindSimilarColor(rgb[0], rgb[1], rgb[2]);
        return new XSSFColor(rgb, null);
    }

    /// <summary>
    /// 转换条件格式比较运算符。
    /// </summary>
    /// <param name="value">待转换的值。</param>
    /// <returns>对应的 NPOI 比较运算符。</returns>
    private static ComparisonOperator ToComparison(ExcelConditionalComparisonOperator value) => value switch
    {
        ExcelConditionalComparisonOperator.Equal => ComparisonOperator.Equal,
        ExcelConditionalComparisonOperator.NotEqual => ComparisonOperator.NotEqual,
        ExcelConditionalComparisonOperator.GreaterThan => ComparisonOperator.GreaterThan,
        ExcelConditionalComparisonOperator.GreaterThanOrEqual => ComparisonOperator.GreaterThanOrEqual,
        ExcelConditionalComparisonOperator.LessThan => ComparisonOperator.LessThan,
        ExcelConditionalComparisonOperator.LessThanOrEqual => ComparisonOperator.LessThanOrEqual,
        ExcelConditionalComparisonOperator.Between => ComparisonOperator.Between,
        ExcelConditionalComparisonOperator.NotBetween => ComparisonOperator.NotBetween,
        _ => throw new ArgumentOutOfRangeException(nameof(value))
    };

    /// <summary>
    /// 将公共区域定义转换为 NPOI 区域。
    /// </summary>
    /// <param name="range">待处理的单元格区域。</param>
    /// <returns>对应的 NPOI 单元格区域。</returns>
    private static CellRangeAddress ToCellRange(ExcelRangeDefinition range)
    {
        range.Validate(nameof(range));
        return new CellRangeAddress(range.StartRow, range.EndRow, range.StartColumn, range.EndColumn);
    }

    /// <summary>
    /// 生成命名范围的工作表引用公式。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="definition">待应用的定义。</param>
    /// <returns>包含工作表名称的区域引用公式。</returns>
    private static string ToFormula(ISheet sheet, ExcelNamedRangeDefinition definition)
    {
        var sheetName = string.IsNullOrWhiteSpace(definition.SheetName) ? sheet.SheetName : definition.SheetName;
        return $"'{sheetName.Replace("'", "''")}'!{definition.Address}";
    }

    /// <summary>
    /// 生成区域的 A1 地址。
    /// </summary>
    /// <param name="range">待处理的单元格区域。</param>
    /// <param name="endRowOffset">区域结束行的附加偏移量。</param>
    /// <returns>区域的 A1 地址。</returns>
    private static string ToAddress(ExcelRangeDefinition range, int endRowOffset = 0) =>
        $"{ToColumn(range.StartColumn)}{range.StartRow + 1}:{ToColumn(range.EndColumn)}{range.EndRow + endRowOffset + 1}";

    /// <summary>
    /// 将零基列索引转换为列名。
    /// </summary>
    /// <param name="index">起始索引，从 0 开始。</param>
    /// <returns>对应的 Excel 列名。</returns>
    private static string ToColumn(int index)
    {
        var value = index + 1;
        var result = string.Empty;
        while (value > 0)
        {
            value--;
            result = (char)('A' + value % 26) + result;
            value /= 26;
        }
        return result;
    }

    /// <summary>
    /// 解析十六进制颜色的 RGB 分量。
    /// </summary>
    /// <param name="value">待转换的值。</param>
    /// <returns>按红、绿、蓝顺序排列的颜色字节。</returns>
    private static byte[] ParseColor(string value)
    {
        var hex = value.Trim().TrimStart('#');
        if (hex.Length == 8) hex = hex.Substring(2);
        if (hex.Length != 6) throw new ArgumentException("颜色必须是六位十六进制值。", nameof(value));
        return new[]
        {
            Convert.ToByte(hex.Substring(0, 2), 16), Convert.ToByte(hex.Substring(2, 2), 16),
            Convert.ToByte(hex.Substring(4, 2), 16)
        };
    }

    /// <summary>
    /// 应用工作表打印布局。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="options">工作表打印布局选项。</param>
    private static void ApplyPrintLayout(ISheet sheet, ExcelPrintLayoutOptions options)
    {
        var setup = sheet.PrintSetup;
        setup.Landscape = options.Orientation == ExcelPrintOrientation.Landscape;
        setup.PaperSize = options.PaperSize switch
        {
            ExcelPrintPaperSize.Letter => (short)PaperSize.US_Letter_Small,
            ExcelPrintPaperSize.A3 => (short)PaperSize.A3,
            ExcelPrintPaperSize.Legal => (short)PaperSize.US_Legal,
            _ => (short)PaperSize.A4
        };
        if (options.ScalePercent.HasValue) setup.Scale = options.ScalePercent.Value;
        if (options.FitToWidth.HasValue) setup.FitWidth = options.FitToWidth.Value;
        if (options.FitToHeight.HasValue) setup.FitHeight = options.FitToHeight.Value;
        sheet.FitToPage = options.FitToWidth.HasValue || options.FitToHeight.HasValue;
        sheet.SetMargin(MarginType.TopMargin, options.TopMargin);
        sheet.SetMargin(MarginType.BottomMargin, options.BottomMargin);
        sheet.SetMargin(MarginType.LeftMargin, options.LeftMargin);
        sheet.SetMargin(MarginType.RightMargin, options.RightMargin);
        if (!string.IsNullOrWhiteSpace(options.PrintArea))
            sheet.Workbook.SetPrintArea(sheet.Workbook.GetSheetIndex(sheet), options.PrintArea);
        if (!string.IsNullOrWhiteSpace(options.RepeatRows))
            sheet.RepeatingRows = CellRangeAddress.ValueOf(NormalizeRepeatRange(options.RepeatRows));
        if (!string.IsNullOrWhiteSpace(options.RepeatColumns))
            sheet.RepeatingColumns = CellRangeAddress.ValueOf(NormalizeRepeatRange(options.RepeatColumns));
        if (!string.IsNullOrWhiteSpace(options.Header)) sheet.Header.Center = options.Header;
        if (!string.IsNullOrWhiteSpace(options.Footer)) sheet.Footer.Center = options.Footer;
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
        new(message, provider: "NPOI", operation: BingOfficesOperation.Export, stage: stage);
}

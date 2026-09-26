using System.Text.RegularExpressions;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;

namespace Bing.Offices.Internals;

/// <summary>
/// 跨 Provider 的报表定义预检。
/// </summary>
internal static class ExcelReportPreflight
{
    /// <summary>
    /// 解析色阶最低值颜色。
    /// </summary>
    /// <param name="definition">条件格式定义。</param>
    /// <returns>配置的前景色；未配置时为红色。</returns>
    internal static string ScaleMinimumColor(ExcelConditionalFormatDefinition definition) =>
        string.IsNullOrWhiteSpace(definition.ForegroundColor) ? "#FF0000" : definition.ForegroundColor;
    /// <summary>
    /// 解析色阶最高值颜色。
    /// </summary>
    /// <param name="definition">条件格式定义。</param>
    /// <returns>配置的背景色；未配置时为绿色。</returns>
    internal static string ScaleMaximumColor(ExcelConditionalFormatDefinition definition) =>
        string.IsNullOrWhiteSpace(definition.BackgroundColor) ? "#00FF00" : definition.BackgroundColor;
    /// <summary>
    /// 解析数据条颜色。
    /// </summary>
    /// <param name="definition">条件格式定义。</param>
    /// <returns>配置的前景色；未配置时为蓝色。</returns>
    internal static string DataBarColor(ExcelConditionalFormatDefinition definition) =>
        string.IsNullOrWhiteSpace(definition.ForegroundColor) ? "#0000FF" : definition.ForegroundColor;
    /// <summary>
    /// 校验表格与名称范围标识符允许字符的共享正则表达式。
    /// </summary>
    private static readonly Regex NamePattern = new(@"^[A-Za-z_\\][A-Za-z0-9_.\\]*$",
        RegexOptions.CultureInvariant);
    /// <summary>
    /// 识别与单元格引用冲突的标识符的共享正则表达式。
    /// </summary>
    private static readonly Regex CellReferencePattern = new(@"^[A-Za-z]{1,3}[1-9][0-9]*$",
        RegexOptions.CultureInvariant);
    /// <summary>
    /// 提取有界 A1 地址起止行列的共享正则表达式。
    /// </summary>
    private static readonly Regex AddressPattern = new(
        @"^\$?(?<startColumn>[A-Za-z]{1,3})\$?(?<startRow>[1-9][0-9]*)(?::\$?(?<endColumn>[A-Za-z]{1,3})\$?(?<endRow>[1-9][0-9]*))?$",
        RegexOptions.CultureInvariant);

    /// <summary>
    /// 在工作簿创建和数据枚举前验证所有报表定义。
    /// </summary>
    /// <param name="request">待导出的工作簿请求。</param>
    /// <param name="provider">用于能力判断与异常上下文的 Provider 名称。</param>
    /// <param name="maxRows">目标格式允许的最大行数。</param>
    /// <param name="maxColumns">目标格式允许的最大列数。</param>
    /// <param name="supportsFreezeOrigin">是否支持自定义冻结窗格可视起点。</param>
    internal static void Validate(ExcelWorkbookExportRequest request, string provider,
        int maxRows, int maxColumns, bool supportsFreezeOrigin)
    {
        var tableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var workbookNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var localNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var sheet in request.Sheets)
        {
            foreach (var table in sheet.Tables)
            {
                table.Validate(nameof(sheet.Tables));
                ValidateIdentifier(table.Name, "表格");
                ValidateTableStyle(provider == "SpreadCheetah" && !string.IsNullOrWhiteSpace(table.StyleName)
                    && !table.StyleName.StartsWith("TableStyle", StringComparison.OrdinalIgnoreCase)
                    ? "TableStyle" + table.StyleName : table.StyleName);
                ValidateRange(table.Range, maxRows, maxColumns, "表格");
                if (table.ShowTotals && table.Range.EndRow >= maxRows - 1)
                    throw Configuration("表格汇总行超出工作表行边界。");
                if (!tableNames.Add(table.Name))
                    throw Configuration($"表格名称重复: {table.Name}");
                if (!table.HasHeaders)
                    throw Unsupported(provider, "报表表格必须包含表头。");
            }

            ValidateFilters(sheet, maxRows, maxColumns);
            ValidateFreezePane(sheet.FreezePane, provider, maxRows, maxColumns, supportsFreezeOrigin);

            foreach (var conditional in sheet.ConditionalFormats)
            {
                conditional.Validate(nameof(sheet.ConditionalFormats));
                ValidateRange(conditional.Range, maxRows, maxColumns, "条件格式");
                ValidateConditionalColors(conditional);
            }

            foreach (var namedRange in sheet.NamedRanges)
            {
                namedRange.Validate(nameof(sheet.NamedRanges));
                ValidateIdentifier(namedRange.Name, "名称范围");
                ValidateAddress(namedRange.Address, maxRows, maxColumns, "名称范围");
                if (!string.IsNullOrWhiteSpace(namedRange.SheetName)
                    && !string.Equals(namedRange.SheetName, sheet.Name, StringComparison.OrdinalIgnoreCase))
                    throw Configuration($"名称范围 {namedRange.Name} 的 Sheet 作用域必须与声明 Sheet 一致。");
                var names = string.IsNullOrWhiteSpace(namedRange.SheetName) ? workbookNames : localNames;
                var key = string.IsNullOrWhiteSpace(namedRange.SheetName)
                    ? namedRange.Name : $"{sheet.Name}\0{namedRange.Name}";
                if (!names.Add(key))
                    throw Configuration($"名称范围在相同作用域内重复: {namedRange.Name}");
            }

            ValidatePrintLayout(sheet.PrintLayout, maxRows, maxColumns);
        }
    }

    /// <summary>
    /// 验证自动筛选边界及其与表格区域的兼容性。
    /// </summary>
    /// <param name="sheet">工作表导出请求。</param>
    /// <param name="maxRows">格式允许的最大行数。</param>
    /// <param name="maxColumns">格式允许的最大列数。</param>
    private static void ValidateFilters(ExcelSheetExportRequest sheet, int maxRows, int maxColumns)
    {
        ExcelRangeDefinition standalone = null;
        foreach (var filter in sheet.AutoFilters)
        {
            filter.Validate(nameof(sheet.AutoFilters));
            ValidateRange(filter.Range, maxRows, maxColumns, "自动筛选");
            var matchingTable = sheet.Tables.FirstOrDefault(table => SameRange(table.Range, filter.Range));
            if (matchingTable != null)
                continue;
            if (sheet.Tables.Any(table => Overlaps(table.Range, filter.Range)))
                throw Configuration("自动筛选与表格区域重叠但不完全相同。");
            if (standalone != null && !SameRange(standalone, filter.Range))
                throw Configuration("同一 Sheet 不能声明多个不同的独立自动筛选区域。");
            standalone = filter.Range;
        }
    }

    /// <summary>
    /// 验证冻结窗格边界与 Provider 支持能力。
    /// </summary>
    /// <param name="pane">冻结窗格定义；为空时跳过验证。</param>
    /// <param name="provider">用于异常上下文的 Provider 名称。</param>
    /// <param name="maxRows">格式允许的最大行数。</param>
    /// <param name="maxColumns">格式允许的最大列数。</param>
    /// <param name="supportsFreezeOrigin">是否支持自定义可视起点。</param>
    private static void ValidateFreezePane(ExcelFreezePaneDefinition pane, string provider,
        int maxRows, int maxColumns, bool supportsFreezeOrigin)
    {
        if (pane == null)
            return;
        pane.Validate(nameof(pane));
        if (pane.Rows >= maxRows || pane.Columns >= maxColumns)
            throw Configuration("冻结窗格超出工作表边界。");
        if (pane.TopRow >= maxRows || pane.LeftColumn >= maxColumns)
            throw Configuration("冻结窗格可视起点超出工作表边界。");
        if (!supportsFreezeOrigin && (pane.TopRow.HasValue || pane.LeftColumn.HasValue))
            throw Unsupported(provider, $"{provider} 当前版本不能表达自定义冻结窗格可视起点。");
        if (pane.TopRow.HasValue && pane.TopRow.Value < pane.Rows
            || pane.LeftColumn.HasValue && pane.LeftColumn.Value < pane.Columns)
            throw Configuration("冻结窗格可视起点不能位于冻结区域内部。");
    }

    /// <summary>
    /// 验证打印布局、缩放选项和打印区域。
    /// </summary>
    /// <param name="options">打印选项；为空时跳过验证。</param>
    /// <param name="maxRows">格式允许的最大行数。</param>
    /// <param name="maxColumns">格式允许的最大列数。</param>
    private static void ValidatePrintLayout(ExcelPrintLayoutOptions options, int maxRows, int maxColumns)
    {
        if (options == null)
            return;
        options.Validate(nameof(options));
        if (options.ScalePercent.HasValue
            && (options.FitToWidth.HasValue || options.FitToHeight.HasValue))
            throw Configuration("打印缩放百分比不能与适应页宽或页高同时设置。");
        if (!string.IsNullOrWhiteSpace(options.PrintArea))
            ValidateAddress(options.PrintArea, maxRows, maxColumns, "打印区域");
        if (!string.IsNullOrWhiteSpace(options.RepeatRows))
        {
            var match = Regex.Match(options.RepeatRows, @"^\$?([1-9][0-9]*):\$?([1-9][0-9]*)$");
            if (!match.Success || !int.TryParse(match.Groups[1].Value, out var first)
                || !int.TryParse(match.Groups[2].Value, out var last)
                || first > last || last > maxRows)
                throw Configuration("重复打印标题行必须是格式边界内的正向行区域。");
        }
        if (!string.IsNullOrWhiteSpace(options.RepeatColumns))
        {
            var match = Regex.Match(options.RepeatColumns, @"^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$");
            if (!match.Success || ToColumnIndex(match.Groups[1].Value) > ToColumnIndex(match.Groups[2].Value)
                || ToColumnIndex(match.Groups[2].Value) >= maxColumns)
                throw Configuration("重复打印标题列必须是格式边界内的正向列区域。");
        }
    }

    /// <summary>
    /// 验证条件格式使用的颜色。
    /// </summary>
    /// <param name="definition">条件格式定义。</param>
    private static void ValidateConditionalColors(ExcelConditionalFormatDefinition definition)
    {
        if (definition.Type == ExcelConditionalFormatType.ColorScale)
        {
            ValidateColor(ScaleMinimumColor(definition), "色阶最低值颜色");
            ValidateColor(ScaleMaximumColor(definition), "色阶最高值颜色");
        }
        else if (definition.Type == ExcelConditionalFormatType.DataBar)
            ValidateColor(DataBarColor(definition), "数据条颜色");
        else
        {
            if (!string.IsNullOrWhiteSpace(definition.ForegroundColor))
                ValidateColor(definition.ForegroundColor, "条件格式前景色");
            if (!string.IsNullOrWhiteSpace(definition.BackgroundColor))
                ValidateColor(definition.BackgroundColor, "条件格式背景色");
        }
    }

    /// <summary>
    /// 验证 RGB 或 ARGB 十六进制颜色。
    /// </summary>
    /// <param name="value">待验证的颜色文本。</param>
    /// <param name="label">用于错误消息的颜色用途名称。</param>
    private static void ValidateColor(string value, string label)
    {
        var hex = value?.Trim().TrimStart('#');
        if (hex?.Length == 8)
            hex = hex.Substring(2);
        if (hex?.Length != 6 || !hex.All(Uri.IsHexDigit))
            throw Configuration($"{label}必须是六位 RGB 或八位 ARGB 十六进制值。");
    }

    /// <summary>
    /// 验证内置表格样式名称和编号。
    /// </summary>
    /// <param name="styleName">表格样式名称；为空白时跳过验证。</param>
    private static void ValidateTableStyle(string styleName)
    {
        if (string.IsNullOrWhiteSpace(styleName))
            return;
        var match = Regex.Match(styleName, @"^TableStyle(?<family>Light|Medium|Dark)(?<index>[1-9][0-9]?)$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
        if (!match.Success)
            throw Configuration($"无效的内置表格样式: {styleName}");
        var index = int.Parse(match.Groups["index"].Value, System.Globalization.CultureInfo.InvariantCulture);
        var maximum = match.Groups["family"].Value.Equals("Light", StringComparison.OrdinalIgnoreCase) ? 21
            : match.Groups["family"].Value.Equals("Medium", StringComparison.OrdinalIgnoreCase) ? 28 : 11;
        if (index > maximum)
            throw Configuration($"无效的内置表格样式: {styleName}");
    }

    /// <summary>
    /// 验证名称字符并排除单元格引用。
    /// </summary>
    /// <param name="value">待验证的标识符。</param>
    /// <param name="label">用于错误消息的定义类别。</param>
    private static void ValidateIdentifier(string value, string label)
    {
        if (!NamePattern.IsMatch(value) || CellReferencePattern.IsMatch(value))
            throw Configuration($"{label}名称无效: {value}");
    }

    /// <summary>
    /// 验证区域方向及工作表边界。
    /// </summary>
    /// <param name="range">待验证的区域。</param>
    /// <param name="maxRows">格式允许的最大行数。</param>
    /// <param name="maxColumns">格式允许的最大列数。</param>
    /// <param name="label">用于错误消息的区域用途。</param>
    private static void ValidateRange(ExcelRangeDefinition range, int maxRows, int maxColumns, string label)
    {
        range.Validate(label);
        if (range.EndRow >= maxRows || range.EndColumn >= maxColumns)
            throw Configuration($"{label}区域超出工作表边界。");
    }

    /// <summary>
    /// 验证有界 A1 地址及工作表边界。
    /// </summary>
    /// <param name="address">单元格或区域的 A1 地址。</param>
    /// <param name="maxRows">格式允许的最大行数。</param>
    /// <param name="maxColumns">格式允许的最大列数。</param>
    /// <param name="label">用于错误消息的地址用途。</param>
    private static void ValidateAddress(string address, int maxRows, int maxColumns, string label)
    {
        var match = AddressPattern.Match(address);
        if (!match.Success)
            throw Configuration($"{label}地址无效。");
        var endColumn = match.Groups["endColumn"].Success
            ? match.Groups["endColumn"].Value : match.Groups["startColumn"].Value;
        var endRow = match.Groups["endRow"].Success
            ? match.Groups["endRow"].Value : match.Groups["startRow"].Value;
        if (!int.TryParse(match.Groups["startRow"].Value, out var firstRow)
            || !int.TryParse(endRow, out var lastRow)
            || firstRow > lastRow || lastRow > maxRows
            || ToColumnIndex(match.Groups["startColumn"].Value) > ToColumnIndex(endColumn)
            || ToColumnIndex(endColumn) >= maxColumns)
            throw Configuration($"{label}地址超出工作表边界。");
    }

    /// <summary>
    /// 将列字母转换为零基列索引。
    /// </summary>
    /// <param name="name">不含美元符号的列字母。</param>
    /// <returns>零基列索引。</returns>
    private static int ToColumnIndex(string name)
    {
        var result = 0;
        foreach (var character in name.ToUpperInvariant())
            result = checked(result * 26 + character - 'A' + 1);
        return result - 1;
    }

    /// <summary>
    /// 判断两个区域的边界是否相同。
    /// </summary>
    /// <param name="left">第一个区域。</param>
    /// <param name="right">第二个区域。</param>
    /// <returns>两个区域均非空且边界相同时返回 <see langword="true"/>，否则返回 <see langword="false"/>。</returns>
    internal static bool SameRange(ExcelRangeDefinition left, ExcelRangeDefinition right) =>
        left != null && right != null && left.StartRow == right.StartRow
        && left.StartColumn == right.StartColumn && left.EndRow == right.EndRow
        && left.EndColumn == right.EndColumn;

    /// <summary>
    /// 判断两个区域是否存在交集。
    /// </summary>
    /// <param name="left">第一个非空区域。</param>
    /// <param name="right">第二个非空区域。</param>
    /// <returns>至少共享一个单元格时返回 <see langword="true"/>，否则返回 <see langword="false"/>。</returns>
    private static bool Overlaps(ExcelRangeDefinition left, ExcelRangeDefinition right) =>
        left.StartRow <= right.EndRow && right.StartRow <= left.EndRow
        && left.StartColumn <= right.EndColumn && right.StartColumn <= left.EndColumn;

    /// <summary>
    /// 创建报表预检配置异常。
    /// </summary>
    /// <param name="message">错误原因。</param>
    /// <returns>标记为预检阶段的配置异常。</returns>
    private static BingOfficesConfigurationException Configuration(string message) =>
        new(message, stage: BingOfficesStage.Preflight);

    /// <summary>
    /// 创建报表导出能力不支持异常。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="message">不支持的能力说明。</param>
    /// <returns>包含 Provider、导出操作与预检阶段的异常。</returns>
    private static BingOfficesUnsupportedFeatureException Unsupported(string provider, string message) =>
        new(message, provider: provider, operation: BingOfficesOperation.Export,
            stage: BingOfficesStage.Preflight);
}

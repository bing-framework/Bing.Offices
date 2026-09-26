using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using Bing.Offices.ClosedXml.Internals;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.Validations;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Imports;

/// <summary>
/// 将 ClosedXML 的原生 Data Validation 规则转换为公共校验结果。
/// </summary>
internal static class ClosedXmlWorkbookValidationPipeline
{
    /// <summary>
    /// 识别受支持的单元格与数值或文本常量比较公式。
    /// </summary>
    private static readonly Regex SimpleComparison = new(
        @"^(?<left>(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)!)?\$?[A-Za-z]{1,3}\$?\d+)\s*(?<op>=|<>|>=|<=|>|<)\s*(?<right>[-+]?\d+(?:\.\d+)?|""[^""]*"")$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// 识别列表校验使用的 A1 矩形区域地址。
    /// </summary>
    private static readonly Regex A1Range = new(
        @"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*:\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// 识别命名范围使用的绝对单元格或区域地址。
    /// </summary>
    private static readonly Regex AbsoluteNamedAddress = new(
        @"^\$(?<startColumn>[A-Za-z]{1,3})\$(?<startRow>[1-9][0-9]*)(?::\$(?<endColumn>[A-Za-z]{1,3})\$(?<endRow>[1-9][0-9]*))?$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// 识别可选工作表限定的 A1 单元格引用。
    /// </summary>
    private static readonly Regex A1Reference = new(
        @"^(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)!)?\$?[A-Za-z]{1,3}\$?\d+$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// 用于解析工作簿日期与时间校验值的共享规则。
    /// </summary>
    private static readonly DateTimeExcelValidationRule DateRule = new();

    /// <summary>
    /// 表示一次原生规则的可观察结果。
    /// </summary>
    internal readonly struct Result
    {
        /// <summary>
        /// 初始化一个 <see cref="Result"/> 类型的实例。
        /// </summary>
        /// <param name="valid">是否通过校验。</param>
        /// <param name="unsupported">是否超出支持范围。</param>
        /// <param name="message">错误消息。</param>
        internal Result(bool valid, bool unsupported, string message)
        {
            IsValid = valid;
            IsUnsupported = unsupported;
            Message = message;
        }

        /// <summary>
        /// 获取规则是否通过校验。
        /// </summary>
        internal bool IsValid { get; }
        /// <summary>
        /// 获取规则是否超出支持范围。
        /// </summary>
        internal bool IsUnsupported { get; }
        /// <summary>
        /// 获取校验失败消息。
        /// </summary>
        internal string Message { get; }
    }

    /// <summary>
    /// 校验当前单元格适用的原生规则。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="cell">待处理的单元格。</param>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="text">单元格文本。</param>
    /// <param name="culture">调用方传入的区域设置。</param>
    /// <param name="isDate1904">工作簿是否使用 1904 日期系统。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>首个未通过的规则结果；全部通过时返回成功结果。</returns>
    internal static Result Validate(IXLWorksheet worksheet, IXLCell cell, object raw, string text,
        CultureInfo culture, bool isDate1904, CancellationToken cancellationToken)
    {
        var validations = worksheet.DataValidations.GetAllInRange(cell.AsRange().RangeAddress).ToArray();
        foreach (var validation in validations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var result = ValidateOne(worksheet, validation, raw, text, culture, isDate1904,
                cancellationToken);
            if (!result.IsValid)
                return result;
        }
        return new Result(true, false, null);
    }

    /// <summary>
    /// 校验单元格是否满足指定原生规则。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="validation">工作簿原生数据校验规则。</param>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="text">单元格文本。</param>
    /// <param name="culture">调用方传入的区域设置。</param>
    /// <param name="isDate1904">工作簿是否使用 1904 日期系统。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>指定规则的校验结果。</returns>
    private static Result ValidateOne(IXLWorksheet worksheet, IXLDataValidation validation,
        object raw, string text, CultureInfo culture, bool isDate1904,
        CancellationToken cancellationToken)
    {
        text ??= string.Empty;
        if (string.IsNullOrWhiteSpace(text))
            return validation.IgnoreBlanks
                ? Valid()
                : Invalid("不允许 Workbook 校验目标为空。");
        if (validation.AllowedValues == XLAllowedValues.AnyValue)
            return new Result(true, false, null);

        return validation.AllowedValues switch
        {
            XLAllowedValues.WholeNumber => ValidateNumber(validation, text, wholeNumber: true, "整数"),
            XLAllowedValues.Decimal => ValidateNumber(validation, text, wholeNumber: false, "小数"),
            XLAllowedValues.TextLength => ValidateTextLength(validation, text),
            XLAllowedValues.Date => ValidateDate(validation, raw, text, culture, isDate1904, timeOnly: false),
            XLAllowedValues.Time => ValidateDate(validation, raw, text, culture, isDate1904, timeOnly: true),
            XLAllowedValues.List => ValidateList(worksheet, validation, text, cancellationToken),
            XLAllowedValues.Custom => ValidateCustom(worksheet, validation, raw, text, culture),
            _ => Unsupported()
        };
    }

    /// <summary>
    /// 校验数值及其允许范围。
    /// </summary>
    /// <param name="validation">工作簿原生数据校验规则。</param>
    /// <param name="text">单元格文本。</param>
    /// <param name="wholeNumber">是否要求数值为整数。</param>
    /// <param name="label">校验失败消息使用的规则名称。</param>
    /// <returns>数值规则的校验结果。</returns>
    private static Result ValidateNumber(IXLDataValidation validation, string text, bool wholeNumber,
        string label)
    {
        if (!decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var value))
            return Invalid($"不符合 Workbook {label}校验。");
        if (wholeNumber && decimal.Truncate(value) != value)
            return Invalid("不符合 Workbook 整数校验。");
        var first = ParseDecimal(validation.MinValue ?? validation.Value);
        var second = ParseDecimal(validation.MaxValue);
        if (!first.HasValue || (RequiresSecond(validation.Operator) && !second.HasValue))
            return Unsupported();
        var valid = Compare(value, first.Value, second.GetValueOrDefault(), validation.Operator,
            second.HasValue);
        return valid ? Valid() : Invalid($"不符合 Workbook {label}校验。");
    }

    /// <summary>
    /// 校验文本长度。
    /// </summary>
    /// <param name="validation">工作簿原生数据校验规则。</param>
    /// <param name="text">单元格文本。</param>
    /// <returns>文本长度规则的校验结果。</returns>
    private static Result ValidateTextLength(IXLDataValidation validation, string text)
    {
        var first = ParseDecimal(validation.MinValue ?? validation.Value);
        var second = ParseDecimal(validation.MaxValue);
        if (!first.HasValue || (RequiresSecond(validation.Operator) && !second.HasValue))
            return Unsupported();
        var valid = Compare(text.Length, first.Value, second.GetValueOrDefault(), validation.Operator,
            second.HasValue);
        return valid ? Valid() : Invalid("不符合 Workbook 文本长度校验。");
    }

    /// <summary>
    /// 校验日期或时间及其允许范围。
    /// </summary>
    /// <param name="validation">工作簿原生数据校验规则。</param>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="text">单元格文本。</param>
    /// <param name="culture">调用方传入的区域设置。</param>
    /// <param name="isDate1904">工作簿是否使用 1904 日期系统。</param>
    /// <param name="timeOnly">为 true 时仅解析时间；否则解析日期。</param>
    /// <returns>日期或时间规则的校验结果。</returns>
    private static Result ValidateDate(IXLDataValidation validation, object raw, string text,
        CultureInfo culture, bool isDate1904, bool timeOnly)
    {
        if (!TryParseDate(raw, text, isDate1904, timeOnly, out var value))
            return Invalid("不符合 Workbook 日期/时间校验。");
        if (!TryParseDate(null, validation.MinValue ?? validation.Value, isDate1904, timeOnly, out var first))
            return Unsupported();
        var second = default(DateTime);
        var secondParsed = !RequiresSecond(validation.Operator)
            || TryParseDate(null, validation.MaxValue, isDate1904, timeOnly, out second);
        if (!secondParsed)
            return Unsupported();
        var valid = Compare(value, first, second, validation.Operator, secondParsed);
        return valid ? Valid() : Invalid("不符合 Workbook 日期/时间校验。");
    }

    /// <summary>
    /// 校验文本是否属于允许的列表。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="validation">工作簿原生数据校验规则。</param>
    /// <param name="text">单元格文本。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>列表规则的校验结果。</returns>
    private static Result ValidateList(IXLWorksheet worksheet, IXLDataValidation validation, string text,
        CancellationToken cancellationToken)
    {
        var expression = (validation.Value ?? string.Empty).Trim();
        if (expression.Length == 0)
            return Unsupported();
        if (expression.StartsWith("=", StringComparison.Ordinal))
            expression = expression.Substring(1).Trim();
        if (expression.Length >= 2 && expression[0] == '"' && expression[^1] == '"')
        {
            var explicitText = expression.Substring(1, expression.Length - 2);
            if (!explicitText.Contains(",", StringComparison.Ordinal))
                return string.Equals(explicitText, text, StringComparison.Ordinal)
                    ? Valid()
                    : Invalid("不符合 Workbook 列表校验。");
            foreach (var value in explicitText.Split(','))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var normalized = value.Trim();
                if (normalized.Length != 0 && string.Equals(normalized, text, StringComparison.Ordinal))
                    return Valid();
            }
            return Invalid("不符合 Workbook 列表校验。");
        }

        IXLRange range;
        if (TryResolveDirectRange(worksheet, expression, out range))
            return ValidateListRange(range, text, cancellationToken, rejectFormulaCells: false);
        if (!TryResolveNamedRange(worksheet, expression, out range))
            return Unsupported();
        return ValidateListRange(range, text, cancellationToken, rejectFormulaCells: true);
    }

    /// <summary>
    /// 校验受支持的自定义比较公式。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="validation">工作簿原生数据校验规则。</param>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="text">单元格文本。</param>
    /// <param name="culture">调用方传入的区域设置。</param>
    /// <returns>自定义公式的校验结果。</returns>
    private static Result ValidateCustom(IXLWorksheet worksheet, IXLDataValidation validation,
        object raw, string text, CultureInfo culture)
    {
        var formula = (validation.Value ?? string.Empty).Trim();
        if (formula.StartsWith("=", StringComparison.Ordinal))
            formula = formula.Substring(1).Trim();
        var match = SimpleComparison.Match(formula);
        if (!match.Success)
            return Unsupported();
        var left = match.Groups["left"].Value;
        var right = match.Groups["right"].Value.Trim('"');
        if (!TryResolveCellText(worksheet, left, out var leftValue))
            return Unsupported();
        if (!decimal.TryParse(leftValue, NumberStyles.Number, CultureInfo.InvariantCulture, out var number)
            || !decimal.TryParse(right, NumberStyles.Number, CultureInfo.InvariantCulture, out var bound))
        {
            var textComparison = string.Equals(leftValue, right, StringComparison.Ordinal);
            return EvaluateTextComparison(textComparison, match.Groups["op"].Value)
                ? Valid()
                : Invalid("不符合 Workbook 自定义校验。");
        }
        var valid = match.Groups["op"].Value switch
        {
            "=" => number == bound,
            "<>" => number != bound,
            ">" => number > bound,
            "<" => number < bound,
            ">=" => number >= bound,
            "<=" => number <= bound,
            _ => false
        };
        return valid ? Valid() : Invalid("不符合 Workbook 自定义校验。");
    }

    /// <summary>
    /// 校验文本是否存在于列表区域。
    /// </summary>
    /// <param name="range">待处理的单元格区域。</param>
    /// <param name="text">单元格文本。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="rejectFormulaCells">是否拒绝包含公式单元格的列表区域。</param>
    /// <returns>列表区域的校验结果。</returns>
    private static Result ValidateListRange(IXLRange range, string text,
        CancellationToken cancellationToken, bool rejectFormulaCells)
    {
        var found = false;
        try
        {
            var cells = rejectFormulaCells
                ? range.CellsUsed(XLCellsUsedOptions.Contents)
                : range.Cells();
            foreach (var cell in cells)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (rejectFormulaCells && cell.HasFormula)
                    return Unsupported();
                if (string.Equals(cell.GetString(), text, StringComparison.Ordinal))
                    found = true;
            }
        }
        catch (Exception exception) when (exception is ArgumentException
                                          || exception is FormatException
                                          || exception is InvalidOperationException)
        {
            return Unsupported();
        }

        return found ? Valid() : Invalid("不符合 Workbook 列表校验。");
    }

    /// <summary>
    /// 尝试解析直接引用的列表区域。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="expression">待解析的引用表达式。</param>
    /// <param name="range">解析得到的区域；失败时为 null。</param>
    /// <returns>成功解析区域时返回 true；引用无效或无法解析时返回 false。</returns>
    private static bool TryResolveDirectRange(IXLWorksheet worksheet, string expression,
        out IXLRange range)
    {
        range = null;
        if (!TrySplitReference(expression, out var sheetName, out var address)
            || !A1Range.IsMatch(address))
            return false;
        return TryGetRange(worksheet, sheetName, address, out range);
    }

    /// <summary>
    /// 尝试解析工作表或工作簿命名范围。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="name">命名范围名称。</param>
    /// <param name="range">解析得到的区域；失败时为 null。</param>
    /// <returns>找到并解析命名范围时返回 true；否则返回 false。</returns>
    private static bool TryResolveNamedRange(IXLWorksheet worksheet, string name, out IXLRange range)
    {
        range = null;
        IXLDefinedName definedName;
        if (worksheet.DefinedNames.Contains(name))
        {
            definedName = worksheet.DefinedNames.DefinedName(name);
            return TryResolveDefinedName(worksheet, definedName, out range);
        }
        if (!worksheet.Workbook.DefinedNames.Contains(name))
            return false;
        definedName = worksheet.Workbook.DefinedNames.DefinedName(name);
        return TryResolveDefinedName(worksheet, definedName, out range);
    }

    /// <summary>
    /// 尝试解析命名范围的绝对引用。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="definedName">待解析的命名范围。</param>
    /// <param name="range">解析得到的区域；失败时为 null。</param>
    /// <returns>名称引用有效且受支持时返回 true；否则返回 false。</returns>
    private static bool TryResolveDefinedName(IXLWorksheet worksheet, IXLDefinedName definedName,
        out IXLRange range)
    {
        range = null;
        if (definedName == null || !definedName.IsValid)
            return false;
        var reference = (definedName.RefersTo ?? string.Empty).Trim();
        if (reference.StartsWith("=", StringComparison.Ordinal))
            reference = reference.Substring(1).Trim();
        if (!TrySplitReference(reference, out var sheetName, out var address))
            return false;
        var match = AbsoluteNamedAddress.Match(address);
        if (!match.Success)
            return false;
        var endColumn = match.Groups["endColumn"].Success
            ? match.Groups["endColumn"].Value : match.Groups["startColumn"].Value;
        var endRow = match.Groups["endRow"].Success
            ? match.Groups["endRow"].Value : match.Groups["startRow"].Value;
        if (!string.Equals(match.Groups["startColumn"].Value, endColumn,
                StringComparison.OrdinalIgnoreCase)
            && !string.Equals(match.Groups["startRow"].Value, endRow, StringComparison.Ordinal))
            return false;
        return TryGetRange(worksheet, sheetName, address, out range);
    }

    /// <summary>
    /// 尝试获取指定工作表区域。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="sheetName">工作表名称；未限定时使用当前工作表。</param>
    /// <param name="address">A1 格式的区域地址。</param>
    /// <param name="range">解析得到的区域；失败时为 null。</param>
    /// <returns>成功获取区域时返回 true；工作表不存在或地址无效时返回 false。</returns>
    private static bool TryGetRange(IXLWorksheet worksheet, string sheetName, string address,
        out IXLRange range)
    {
        range = null;
        try
        {
            var target = worksheet;
            if (!string.IsNullOrEmpty(sheetName)
                && !worksheet.Workbook.TryGetWorksheet(sheetName, out target))
                return false;
            range = target.Range(address.Replace("$", string.Empty, StringComparison.Ordinal));
            return range != null;
        }
        catch (Exception exception) when (exception is ArgumentException || exception is FormatException)
        {
            return false;
        }
    }

    /// <summary>
    /// 尝试拆分工作表名称和区域地址。
    /// </summary>
    /// <param name="expression">待解析的引用表达式。</param>
    /// <param name="sheetName">解析得到的工作表名称；未限定时为 null。</param>
    /// <param name="address">解析得到的 A1 地址。</param>
    /// <returns>引用可拆分时返回 true；引用格式不受支持或不完整时返回 false。</returns>
    private static bool TrySplitReference(string expression, out string sheetName, out string address)
    {
        sheetName = null;
        address = null;
        if (string.IsNullOrWhiteSpace(expression) || expression.IndexOf('[') >= 0
            || expression.IndexOf(']') >= 0 || expression.IndexOf(',') >= 0)
            return false;
        var quoted = false;
        var separator = -1;
        for (var index = 0; index < expression.Length; index++)
        {
            if (expression[index] == '\'')
            {
                if (quoted && index + 1 < expression.Length && expression[index + 1] == '\'')
                {
                    index++;
                    continue;
                }
                quoted = !quoted;
            }
            else if (expression[index] == '!' && !quoted)
                separator = index;
        }
        if (quoted)
            return false;
        address = (separator < 0 ? expression : expression.Substring(separator + 1)).Trim();
        if (separator < 0)
            return address.Length != 0;
        var sheetToken = expression.Substring(0, separator).Trim();
        if (sheetToken.Length >= 2 && sheetToken[0] == '\'' && sheetToken[^1] == '\'')
            sheetName = sheetToken.Substring(1, sheetToken.Length - 2).Replace("''", "'",
                StringComparison.Ordinal);
        else if (sheetToken.IndexOf('\'') < 0)
            sheetName = sheetToken;
        return !string.IsNullOrWhiteSpace(sheetName) && address.Length != 0;
    }

    /// <summary>
    /// 尝试读取引用单元格的文本。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="reference">A1 格式的单元格引用。</param>
    /// <param name="value">读取的单元格文本；失败时为 null。</param>
    /// <returns>成功读取单元格文本时返回 true；引用无效或读取失败时返回 false。</returns>
    private static bool TryResolveCellText(IXLWorksheet worksheet, string reference, out string value)
    {
        value = null;
        if (!A1Reference.IsMatch(reference))
            return false;
        try
        {
            var expression = reference.Trim();
            var separator = expression.LastIndexOf('!');
            var target = worksheet;
            if (separator >= 0)
            {
                var sheetName = expression.Substring(0, separator).Trim('"', '\'');
                if (!worksheet.Workbook.TryGetWorksheet(sheetName, out target))
                    return false;
                expression = expression.Substring(separator + 1);
            }
            value = target.Cell(expression.Replace("$", string.Empty, StringComparison.Ordinal)).GetString();
            return true;
        }
        catch (Exception exception) when (exception is ArgumentException || exception is FormatException)
        {
            return false;
        }
    }

    /// <summary>
    /// 尝试解析工作簿日期或时间。
    /// </summary>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="text">单元格文本。</param>
    /// <param name="isDate1904">工作簿是否使用 1904 日期系统。</param>
    /// <param name="timeOnly">为 true 时仅解析时间；否则解析日期。</param>
    /// <param name="value">解析得到的日期或时间。</param>
    /// <returns>成功解析日期或时间时返回 true；否则返回 false。</returns>
    private static bool TryParseDate(object raw, string text, bool isDate1904, bool timeOnly,
        out DateTime value)
    {
        var cell = ClosedXmlValueAdapter.CreateCell(raw, text, isDate1904);
        return DateRule.TryParseWorkbookDate(cell, text, timeOnly, isDate1904, out value);
    }

    /// <summary>
    /// 按固定区域格式解析十进制数。
    /// </summary>
    /// <param name="value">待转换的值。</param>
    /// <returns>解析得到的数值；无法解析时返回 null。</returns>
    private static decimal? ParseDecimal(string value) => decimal.TryParse(value,
        NumberStyles.Number, CultureInfo.InvariantCulture, out var number) ? number : null;

    /// <summary>
    /// 判断比较运算是否需要第二个边界值。
    /// </summary>
    /// <param name="operation">比较运算符。</param>
    /// <returns>区间运算返回 true；其他运算返回 false。</returns>
    private static bool RequiresSecond(XLOperator operation) => operation is XLOperator.Between or XLOperator.NotBetween;

    /// <summary>
    /// 判断值是否满足指定边界比较。
    /// </summary>
    /// <param name="value">待比较的值。</param>
    /// <param name="first">第一个比较边界。</param>
    /// <param name="second">第二个比较边界。</param>
    /// <param name="operation">比较运算符。</param>
    /// <param name="secondParsed">是否已成功解析第二个边界。</param>
    /// <returns>值满足比较条件时返回 true；否则返回 false。</returns>
    private static bool Compare(decimal value, decimal first, decimal second, XLOperator operation,
        bool secondParsed) => operation switch
    {
        XLOperator.Between => secondParsed && value >= first && value <= second,
        XLOperator.NotBetween => secondParsed && (value < first || value > second),
        XLOperator.EqualTo => value == first,
        XLOperator.NotEqualTo => value != first,
        XLOperator.GreaterThan => value > first,
        XLOperator.LessThan => value < first,
        XLOperator.EqualOrGreaterThan => value >= first,
        XLOperator.EqualOrLessThan => value <= first,
        _ => false
    };

    /// <summary>
    /// 判断值是否满足指定边界比较。
    /// </summary>
    /// <param name="value">待比较的值。</param>
    /// <param name="first">第一个比较边界。</param>
    /// <param name="second">第二个比较边界。</param>
    /// <param name="operation">比较运算符。</param>
    /// <param name="secondParsed">是否已成功解析第二个边界。</param>
    /// <returns>值满足比较条件时返回 true；否则返回 false。</returns>
    private static bool Compare(DateTime value, DateTime first, DateTime second, XLOperator operation,
        bool secondParsed) => operation switch
    {
        XLOperator.Between => secondParsed && value >= first && value <= second,
        XLOperator.NotBetween => secondParsed && (value < first || value > second),
        XLOperator.EqualTo => value == first,
        XLOperator.NotEqualTo => value != first,
        XLOperator.GreaterThan => value > first,
        XLOperator.LessThan => value < first,
        XLOperator.EqualOrGreaterThan => value >= first,
        XLOperator.EqualOrLessThan => value <= first,
        _ => false
    };

    /// <summary>
    /// 根据运算符判断文本比较结果。
    /// </summary>
    /// <param name="equal">参与比较的文本是否相同。</param>
    /// <param name="operation">比较运算符。</param>
    /// <returns>文本比较符合运算符时返回 true；否则返回 false。</returns>
    private static bool EvaluateTextComparison(bool equal, string operation) => operation switch
    {
        "=" => equal,
        "<>" => !equal,
        _ => false
    };

    /// <summary>
    /// 创建校验通过结果。
    /// </summary>
    /// <returns>表示校验通过的结果。</returns>
    private static Result Valid() => new(true, false, null);
    /// <summary>
    /// 创建校验失败结果。
    /// </summary>
    /// <param name="message">错误消息。</param>
    /// <returns>包含失败消息的校验结果。</returns>
    private static Result Invalid(string message) => new(false, false, message);
    /// <summary>
    /// 创建规则不受支持的校验结果。
    /// </summary>
    /// <returns>表示规则不受支持的校验结果。</returns>
    private static Result Unsupported() => new(false, true,
        "Workbook Data Validation 规则类型或公式暂不支持。");
}

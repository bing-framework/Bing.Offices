using System.Globalization;
using System.Runtime.CompilerServices;
using Bing.Offices.Conversions;
using Bing.Offices.Validations;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Bing.Offices.Imports;

/// <summary>
/// 执行 Workbook 原生 Data Validation，隔离规则解析和错误收集。
/// </summary>
internal static class NpoiWorkbookValidationPipeline
{
    /// <summary>
    /// 保存当前导入 Workbook 中已解析的名称列表，Workbook 回收后缓存随之释放。
    /// </summary>
    private static readonly ConditionalWeakTable<IWorkbook, NamedListWorkbookCache> NamedListCaches =
        new ConditionalWeakTable<IWorkbook, NamedListWorkbookCache>();

    /// <summary>
    /// 校验一行的 Workbook 原生规则，并按调用方模式收集错误。
    /// </summary>
    /// <param name="row">当前待校验的数据行。</param>
    /// <param name="columns">按零基列索引排列的已绑定列计划。</param>
    /// <param name="validationIndex">按单元格坐标索引的原生校验规则。</param>
    /// <param name="sheet">当前数据行所属的工作表。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="rowIndex">当前数据行的零基索引。</param>
    /// <param name="bodyWhitespace">单元格文本的空白处理策略。</param>
    /// <param name="validationFailureMode">发生校验失败后的继续策略。</param>
    /// <param name="unsupportedFeaturePolicy">不支持的原生校验规则的处理策略。</param>
    /// <param name="errors">接收工作簿校验错误的收集器。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="cancellationToken">用于取消名称列表解析的令牌。</param>
    /// <returns>当前行所有规则均通过时为 <see langword="true" />；存在失败或不支持规则时为 <see langword="false" />。</returns>
    public static bool Validate(IRow row, IReadOnlyDictionary<int, ExcelColumnPlan> columns,
        ValidationRangeIndex validationIndex, ISheet sheet, string sheetName, int rowIndex,
        ExcelWhitespacePolicy bodyWhitespace, ExcelValidationFailureMode validationFailureMode,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy, ExcelImportErrorCollector errors,
        bool isDate1904, CancellationToken cancellationToken)
    {
        var valid = true;
        foreach (var column in columns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var cell = row.GetCell(column.Key);
            var cellValue = NpoiExcelImporter.ReadCellValue(cell, isDate1904);
            var value = NpoiExcelImporter.NormalizeText(cellValue.Text, bodyWhitespace);
            foreach (var validation in validationIndex.Get(rowIndex, column.Key))
            {
                if (ValidateValue(validation, cellValue, value, sheet, isDate1904,
                        cancellationToken, out var message))
                    continue;
                errors.Add(new ExcelImportError(ExcelImportErrorCode.WorkbookValidation, message, sheetName,
                    rowIndex + 1, column.Key + 1, column.Value.Property.Name, column.Value.DynamicDefinition?.Key
                    ?? column.Value.Property.Name, column.Value.HeaderName, value));
                if (unsupportedFeaturePolicy == ExcelUnsupportedFeaturePolicy.Report
                    && message != null && message.StartsWith("Workbook Data Validation 规则类型或公式暂不支持",
                        StringComparison.Ordinal))
                    continue;
                valid = false;
                if (validationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                    return false;
            }
        }
        return valid;
    }

    /// <summary>
    /// 使用 Workbook 原生规则校验单个实体单元格，并复用当前工作簿的名称列表缓存。
    /// </summary>
    /// <param name="validation">应用于当前单元格的原生校验规则。</param>
    /// <param name="cellValue">保留原始类型信息的单元格值。</param>
    /// <param name="normalizedValue">按调用方空白策略规范化后的单元格文本。</param>
    /// <param name="sheet">当前单元格所属工作表。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="cancellationToken">用于取消名称列表解析的令牌。</param>
    /// <param name="message">校验失败或规则不支持时返回的说明。</param>
    /// <returns>规则通过时为 <see langword="true" />；失败或不支持时为 <see langword="false" />。</returns>
    internal static bool ValidateCell(IDataValidation validation, ExcelCellValue cellValue,
        string normalizedValue, ISheet sheet, bool isDate1904,
        CancellationToken cancellationToken, out string message)
    {
        if (validation == null)
            throw new ArgumentNullException(nameof(validation));
        if (sheet == null)
            throw new ArgumentNullException(nameof(sheet));
        return ValidateValue(validation, cellValue, normalizedValue, sheet, isDate1904,
            cancellationToken, out message);
    }

    /// <summary>
    /// 按照单个原生数据校验约束验证已规范化的单元格值。
    /// </summary>
    /// <param name="validation">NPOI 提供的数据校验规则。</param>
    /// <param name="cellValue">保留原始类型信息的单元格值。</param>
    /// <param name="value">按正文空白策略规范化后的文本值。</param>
    /// <param name="sheet">解析列表区域时使用的当前工作表。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="cancellationToken">用于取消名称列表解析的令牌。</param>
    /// <param name="message">校验失败或不支持时返回的说明消息。</param>
    /// <returns>规则通过时为 true。</returns>
    private static bool ValidateValue(IDataValidation validation, ExcelCellValue cellValue,
        string value, ISheet sheet, bool isDate1904, CancellationToken cancellationToken,
        out string message)
    {
        var constraint = validation.ValidationConstraint;
        value ??= string.Empty;
        var type = constraint.GetValidationType();
        if (type == NPOI.SS.UserModel.ValidationType.ANY)
        {
            message = null;
            return true;
        }
        if (value.Length == 0)
        {
            if (validation.EmptyCellAllowed)
            {
                message = null;
                return true;
            }
            message = "不允许 Workbook 校验目标为空。";
            return false;
        }
        if (type == NPOI.SS.UserModel.ValidationType.LIST)
        {
            if (!TryGetExplicitListValues(constraint, sheet, cancellationToken, out var values))
                return Unsupported(out message);
            if (values.Contains(value))
            {
                message = null;
                return true;
            }
            message = "不符合 Workbook 显式列表校验。";
            return false;
        }
        if (type == NPOI.SS.UserModel.ValidationType.TEXT_LENGTH)
        {
            if (!decimal.TryParse(GetFormula1(constraint), NumberStyles.Float, CultureInfo.InvariantCulture,
                    out var first))
                return Unsupported(out message);
            var secondParsed = decimal.TryParse(GetFormula2(constraint), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var second);
            var valid = Compare(value.Length, first, second, secondParsed, constraint.Operator);
            message = valid ? null : "不符合 Workbook 文本长度校验。";
            return valid;
        }
        if (type == NPOI.SS.UserModel.ValidationType.INTEGER || type == NPOI.SS.UserModel.ValidationType.DECIMAL)
        {
            var parsed = decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture,
                out var number);
            var firstParsed = decimal.TryParse(GetFormula1(constraint), NumberStyles.Number,
                CultureInfo.InvariantCulture, out var first);
            var secondParsed = decimal.TryParse(GetFormula2(constraint), NumberStyles.Number,
                CultureInfo.InvariantCulture, out var second);
            var valid = parsed && firstParsed && Compare(number, first, second, secondParsed,
                constraint.Operator);
            message = valid ? null : "不符合 Workbook 数值校验。";
            return valid;
        }
        if (type == NPOI.SS.UserModel.ValidationType.DATE || type == NPOI.SS.UserModel.ValidationType.TIME)
        {
            var parsed = TryGetExcelDate(cellValue, value, type == NPOI.SS.UserModel.ValidationType.TIME,
                isDate1904, out var date);
            var timeOnly = type == NPOI.SS.UserModel.ValidationType.TIME;
            var firstText = GetFormula1(constraint);
            var first = TryGetExcelDate(new ExcelCellValue(firstText, firstText, ExcelCellKind.Text,
                    isDate1904: isDate1904), firstText, timeOnly, isDate1904,
                out var minimum);
            var secondText = GetFormula2(constraint);
            var second = TryGetExcelDate(new ExcelCellValue(secondText, secondText, ExcelCellKind.Text,
                    isDate1904: isDate1904), secondText, timeOnly, isDate1904,
                out var maximum);
            var valid = parsed && first && Compare(date, minimum, maximum, second, constraint.Operator);
            message = valid ? null : "不符合 Workbook 日期/时间校验。";
            return valid;
        }
        return Unsupported(out message);
    }

    /// <summary>
    /// 从显式列表、带引号列表或单元格区域公式解析允许值。
    /// </summary>
    /// <param name="constraint">原生列表校验约束。</param>
    /// <param name="currentSheet">解析未限定区域时使用的工作表。</param>
    /// <param name="cancellationToken">用于取消名称列表解析的令牌。</param>
    /// <param name="values">成功时返回允许值集合。</param>
    /// <returns>支持并成功解析列表规则时为 true。</returns>
    private static bool TryGetExplicitListValues(IDataValidationConstraint constraint,
        ISheet currentSheet, CancellationToken cancellationToken, out HashSet<string> values)
    {
        var formula = constraint.Formula1?.Replace("&#34;", "\"").Replace("&quot;", "\"")
            .Replace("\\\"", "\"").Trim();
        if (!string.IsNullOrWhiteSpace(formula)
            && TryGetCellRangeValues(formula, currentSheet, out var rangeValues))
        {
            values = new HashSet<string>(rangeValues, StringComparer.Ordinal);
            return true;
        }
        if (!string.IsNullOrWhiteSpace(formula)
            && (formula.StartsWith("=", StringComparison.Ordinal) || IsNameToken(formula)))
            return TryGetNamedListValues(formula, currentSheet, out values, cancellationToken);
        var explicitValues = constraint.ExplicitListValues;
        if (explicitValues != null && explicitValues.Length > 0)
        {
            if (explicitValues.Length == 1 && (explicitValues[0].Contains(":") || explicitValues[0].Contains("!")))
            {
                values = null;
                return false;
            }
            if (explicitValues.Length == 1 && explicitValues[0].Contains(","))
                explicitValues = SplitExplicitList(explicitValues[0]);
            values = new HashSet<string>(explicitValues, StringComparer.Ordinal);
            return values.Count > 0;
        }
        if (string.IsNullOrWhiteSpace(formula))
        {
            values = null;
            return false;
        }
        if (formula.StartsWith("=", StringComparison.Ordinal))
        {
            values = null;
            return false;
        }
        if (!(formula.StartsWith("\"", StringComparison.Ordinal) && formula.EndsWith("\"",
                StringComparison.Ordinal)))
        {
            values = null;
            return false;
        }
        values = new HashSet<string>(SplitExplicitList(formula.Trim('\"')), StringComparer.Ordinal);
        return values.Count > 0;
    }

    /// <summary>
    /// 读取原生约束的第一个公式或 HSSF 数值后备值。
    /// </summary>
    /// <param name="constraint">原生数据校验约束。</param>
    /// <returns>第一个公式文本；无可用值时为 null。</returns>
    private static string GetFormula1(IDataValidationConstraint constraint)
    {
        if (!string.IsNullOrWhiteSpace(constraint.Formula1))
            return constraint.Formula1;
        var hssfConstraint = constraint as DVConstraint;
        return hssfConstraint == null
            ? null
            : hssfConstraint.Value1.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 读取原生约束的第二个公式或 HSSF 数值后备值。
    /// </summary>
    /// <param name="constraint">原生数据校验约束。</param>
    /// <returns>第二个公式文本；无可用值时为 null。</returns>
    private static string GetFormula2(IDataValidationConstraint constraint)
    {
        if (!string.IsNullOrWhiteSpace(constraint.Formula2))
            return constraint.Formula2;
        var hssfConstraint = constraint as DVConstraint;
        return hssfConstraint == null
            ? null
            : hssfConstraint.Value2.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 解析单工作表或带工作表限定的 A1 区域公式，并读取区域文本值。
    /// </summary>
    /// <param name="formula">列表校验中保存的区域公式。</param>
    /// <param name="currentSheet">解析未限定区域时使用的工作表。</param>
    /// <param name="values">成功时返回区域内单元格文本。</param>
    /// <returns>公式是受支持的区域且读取成功时为 true。</returns>
    private static bool TryGetCellRangeValues(string formula, ISheet currentSheet, out string[] values)
    {
        values = null;
        var expression = formula.Trim();
        if (expression.StartsWith("=", StringComparison.Ordinal))
            expression = expression.Substring(1).Trim();
        var separator = expression.LastIndexOf('!');
        var sheetName = separator < 0 ? null : expression.Substring(0, separator).Trim('\'');
        var range = separator < 0 ? expression : expression.Substring(separator + 1);
        var parts = range.Split(':');
        if (parts.Length < 1 || parts.Length > 2)
            return false;
        ISheet sheet = currentSheet;
        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            sheet = currentSheet.Workbook.GetSheet(sheetName);
            if (sheet == null)
                return false;
        }
        try
        {
            var first = new CellReference(parts[0].Trim());
            var last = parts.Length == 1 ? first : new CellReference(parts[1].Trim());
            if (first.Row < 0 || first.Col < 0 || last.Row < first.Row || last.Col < first.Col)
                return false;
            var result = new List<string>();
            for (var rowIndex = first.Row; rowIndex <= last.Row; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                for (var columnIndex = first.Col; columnIndex <= last.Col; columnIndex++)
                    result.Add(NpoiExcelImporter.GetRawStringValue(row?.GetCell(columnIndex)));
            }
            values = result.ToArray();
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    /// <summary>
    /// 解析并缓存当前工作簿中定义名称指向的一维有界区域。
    /// </summary>
    /// <param name="formula">列表校验中的名称表达式。</param>
    /// <param name="currentSheet">引用名称的当前工作表。</param>
    /// <param name="values">成功时返回名称区域中的允许值。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>名称引用受支持且解析成功时为 true。</returns>
    private static bool TryGetNamedListValues(string formula, ISheet currentSheet, out HashSet<string> values,
        CancellationToken cancellationToken = default)
    {
        values = null;
        var name = formula.Trim();
        if (name.StartsWith("=", StringComparison.Ordinal))
            name = name.Substring(1).Trim();
        if (!IsNameToken(name))
            return false;

        var workbook = currentSheet.Workbook;
        var cache = NamedListCaches.GetValue(workbook, _ => new NamedListWorkbookCache());
        var sheetIndex = workbook.GetSheetIndex(currentSheet);
        var result = cache.GetOrAdd(sheetIndex, name,
            () => ResolveNamedList(workbook, sheetIndex, name, cancellationToken));
        values = result.Values;
        return result.IsSupported;
    }

    /// <summary>
    /// 检查名称列表导出来源是否受支持。
    /// </summary>
    /// <param name="sheet">名称列表所属的工作表。</param>
    /// <param name="name">名称列表名称。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>名称来源受支持时为 <see langword="true" />；否则为 <see langword="false" />。</returns>
    internal static bool IsSupportedNamedListSource(ISheet sheet, string name, CancellationToken cancellationToken)
        => TryGetNamedListValues(name, sheet, out _, cancellationToken);

    /// <summary>
    /// 按局部优先规则解析名称；无效局部名称会遮蔽同名全局名称。
    /// </summary>
    /// <param name="workbook">名称所在的工作簿。</param>
    /// <param name="sheetIndex">当前校验工作表的索引。</param>
    /// <param name="nameText">待解析的名称。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>成功时返回允许值；名称缺失或引用不受支持时返回 Unsupported。</returns>
    private static NamedListResult ResolveNamedList(IWorkbook workbook, int sheetIndex, string nameText,
        CancellationToken cancellationToken)
    {
        var matches = workbook.GetAllNames().Where(item => string.Equals(item.NameName, nameText,
            StringComparison.OrdinalIgnoreCase)).ToArray();
        var local = matches.Where(item => item.SheetIndex == sheetIndex).ToArray();
        IName selected;
        if (local.Length > 0)
        {
            if (local.Length != 1)
                return NamedListResult.Unsupported;
            selected = local[0];
        }
        else
        {
            var global = matches.Where(item => item.SheetIndex < 0).ToArray();
            if (global.Length != 1)
                return NamedListResult.Unsupported;
            selected = global[0];
        }

        if (!TryParseBoundedAbsoluteRange(selected.RefersToFormula, workbook, sheetIndex,
                out var rangeSheet, out var firstRow, out var lastRow, out var firstColumn,
                out var lastColumn))
            return NamedListResult.Unsupported;

        var values = new List<string>();
        for (var rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = rangeSheet.GetRow(rowIndex);
            for (var columnIndex = firstColumn; columnIndex <= lastColumn; columnIndex++)
            {
                var cell = row?.GetCell(columnIndex);
                if (cell?.CellType == CellType.Formula)
                    return NamedListResult.Unsupported;
                if (cell != null)
                    values.Add(NpoiExcelImporter.GetRawStringValue(cell));
            }
        }
        return new NamedListResult(true, new HashSet<string>(values, StringComparer.Ordinal));
    }

    /// <summary>
    /// 解析名称公式中的绝对单格、单行或单列区域。
    /// </summary>
    /// <param name="formula">名称的引用公式。</param>
    /// <param name="workbook">名称所在的工作簿。</param>
    /// <param name="currentSheetIndex">名称引用未限定工作表时使用的索引。</param>
    /// <param name="rangeSheet">引用目标工作表。</param>
    /// <param name="firstRow">区域起始零基行号。</param>
    /// <param name="lastRow">区域结束零基行号。</param>
    /// <param name="firstColumn">区域起始零基列号。</param>
    /// <param name="lastColumn">区域结束零基列号。</param>
    /// <returns>公式严格匹配受支持区域格式时为 true。</returns>
    private static bool TryParseBoundedAbsoluteRange(string formula, IWorkbook workbook,
        int currentSheetIndex, out ISheet rangeSheet, out int firstRow, out int lastRow,
        out int firstColumn, out int lastColumn)
    {
        rangeSheet = null;
        firstRow = lastRow = firstColumn = lastColumn = 0;
        if (string.IsNullOrWhiteSpace(formula))
            return false;

        var expression = formula.Trim();
        if (expression.StartsWith("=", StringComparison.Ordinal))
            expression = expression.Substring(1).Trim();
        if (expression.Length == 0)
            return false;

        var separator = -1;
        var quoted = false;
        for (var index = 0; index < expression.Length; index++)
        {
            var current = expression[index];
            if (current == '\'')
            {
                if (quoted && index + 1 < expression.Length && expression[index + 1] == '\'')
                    index++;
                else
                    quoted = !quoted;
            }
            else if (!quoted)
            {
                if (current == '!' && separator < 0)
                    separator = index;
                else if (current is '[' or ']' or ',' or '(' or ')' or ';' or '!')
                    return false;
            }
        }
        if (quoted)
            return false;
        if (separator >= 0)
        {
            var sheetToken = expression.Substring(0, separator).Trim();
            string sheetName;
            if (sheetToken.StartsWith("'", StringComparison.Ordinal))
            {
                if (sheetToken.Length < 2 || !sheetToken.EndsWith("'", StringComparison.Ordinal))
                    return false;
                sheetName = sheetToken.Substring(1, sheetToken.Length - 2).Replace("''", "'");
            }
            else
            {
                if (sheetToken.IndexOf('\'') >= 0 || sheetToken.IndexOf(':') >= 0)
                    return false;
                sheetName = sheetToken;
            }
            if (string.IsNullOrWhiteSpace(sheetName))
                return false;
            rangeSheet = workbook.GetSheet(sheetName);
            if (rangeSheet == null)
                return false;
            expression = expression.Substring(separator + 1).Trim();
        }
        else
        {
            rangeSheet = workbook.GetSheetAt(currentSheetIndex);
        }

        var parts = expression.Split(':');
        if (parts.Length < 1 || parts.Length > 2
            || !IsAbsoluteCellReference(parts[0].Trim())
            || (parts.Length == 2 && !IsAbsoluteCellReference(parts[1].Trim())))
            return false;

        try
        {
            var first = new CellReference(parts[0].Trim());
            var last = parts.Length == 1 ? first : new CellReference(parts[1].Trim());
            var version = workbook.SpreadsheetVersion;
            if (first.Row < 0 || first.Col < 0 || last.Row < first.Row || last.Col < first.Col
                || last.Row > version.LastRowIndex || last.Col > version.LastColumnIndex
                || (first.Row != last.Row && first.Col != last.Col))
                return false;
            firstRow = first.Row;
            lastRow = last.Row;
            firstColumn = first.Col;
            lastColumn = last.Col;
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    /// <summary>
    /// 判断文本是否是单一 Excel 名称，而非公式或区域表达式。
    /// </summary>
    /// <param name="value">名称表达式。</param>
    /// <returns>文本符合名称 token 语法时为 true。</returns>
    private static bool IsNameToken(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return false;
        var expression = value.Trim();
        if (expression.StartsWith("=", StringComparison.Ordinal))
            expression = expression.Substring(1).Trim();
        if (expression.Length == 0
            || !(char.IsLetter(expression[0]) || expression[0] == '_' || expression[0] == '\\'))
            return false;
        return expression.Skip(1).All(character => char.IsLetterOrDigit(character)
            || character == '_' || character == '.' || character == '\\');
    }

    /// <summary>
    /// 判断引用是否同时固定了列和行。
    /// </summary>
    /// <param name="address">A1 地址。</param>
    /// <returns>绝对单元格引用时为 true。</returns>
    private static bool IsAbsoluteCellReference(string address)
    {
        if (string.IsNullOrEmpty(address) || address[0] != '$')
            return false;
        var rowMarker = address.IndexOf('$', 1);
        if (rowMarker < 2 || rowMarker == address.Length - 1)
            return false;
        if (address.Substring(1, rowMarker - 1).Any(character => !char.IsLetter(character)))
            return false;
        return address.Substring(rowMarker + 1).All(char.IsDigit)
            && address[rowMarker + 1] != '0';
    }

    /// <summary>
    /// 表示名称解析成功或不受支持的缓存结果。
    /// </summary>
    private sealed class NamedListResult
    {
        /// <summary>
        /// 初始化一个 <see cref="NamedListResult" /> 类型的实例。
        /// </summary>
        /// <param name="isSupported">引用是否受支持。</param>
        /// <param name="values">区域中的列表值。</param>
        internal NamedListResult(bool isSupported, HashSet<string> values)
        {
            IsSupported = isSupported;
            Values = values;
        }

        /// <summary>
        /// 获取名称引用是否受支持。
        /// </summary>
        internal bool IsSupported { get; }
        /// <summary>
        /// 获取名称区域中的列表值。
        /// </summary>
        internal HashSet<string> Values { get; }
        /// <summary>
        /// 获取无法安全解析的名称引用结果。
        /// </summary>
        internal static NamedListResult Unsupported { get; } = new NamedListResult(false, null);
    }

    /// <summary>
    /// 缓存当前工作簿中名称解析和允许值扫描结果。
    /// </summary>
    private sealed class NamedListWorkbookCache
    {
        /// <summary>
        /// 保存按当前工作表和名称区分的解析结果。
        /// </summary>
        private readonly Dictionary<string, NamedListResult> _results =
            new Dictionary<string, NamedListResult>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 获取名称缓存项；首次访问时解析并缓存成功或 Unsupported 结果。
        /// </summary>
        /// <param name="sheetIndex">当前工作表索引。</param>
        /// <param name="name">待解析名称。</param>
        /// <param name="factory">创建缓存值的委托。</param>
        /// <returns>名称解析结果。</returns>
        internal NamedListResult GetOrAdd(int sheetIndex, string name, Func<NamedListResult> factory)
        {
            var key = sheetIndex.ToString(CultureInfo.InvariantCulture) + ":" + name;
            lock (_results)
            {
                if (_results.TryGetValue(key, out var result))
                    return result;
                result = factory();
                _results.Add(key, result);
                return result;
            }
        }
    }

    /// <summary>
    /// 拆分逗号分隔的显式列表并移除外层引号和空白。
    /// </summary>
    /// <param name="value">列表公式中的显式文本。</param>
    /// <returns>非空的列表项集合。</returns>
    private static string[] SplitExplicitList(string value) => value.Trim().Trim('\"')
        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
        .Select(item => item.Trim()).ToArray();

    /// <summary>
    /// 从 typed 单元格值或文本解析 Excel 日期/时间为比较刻度。
    /// </summary>
    /// <param name="cellValue">保留原始数值或日期类型的单元格值。</param>
    /// <param name="text">单元格文本后备值。</param>
    /// <param name="timeOnly">是否仅比较时间部分。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="ticks">成功时返回日期或时间的刻度值。</param>
    /// <returns>成功解析日期或时间时为 true。</returns>
    private static bool TryGetExcelDate(ExcelCellValue cellValue, string text, bool timeOnly, bool isDate1904,
        out long ticks)
    {
        var parser = new DateTimeExcelValidationRule();
        if (!parser.TryParseWorkbookDate(cellValue, text, timeOnly, isDate1904, out var parsed))
        {
            ticks = 0;
            return false;
        }
        ticks = ((DateTime)parsed).Ticks;
        return true;
    }

    /// <summary>
    /// 表示两个边界之间的比较操作。
    /// </summary>
    private const int BetweenOperator = OperatorBetween;
    /// <summary>
    /// 表示两个边界之外的比较操作。
    /// </summary>
    private const int NotBetweenOperator = OperatorNotBetween;
    /// <summary>
    /// 表示相等比较操作。
    /// </summary>
    private const int EqualOperator = OperatorEqual;
    /// <summary>
    /// 表示不相等比较操作。
    /// </summary>
    private const int NotEqualOperator = OperatorNotEqual;
    /// <summary>
    /// 表示大于比较操作。
    /// </summary>
    private const int GreaterThanOperator = OperatorGreaterThan;
    /// <summary>
    /// 表示小于比较操作。
    /// </summary>
    private const int LessThanOperator = OperatorLessThan;
    /// <summary>
    /// 表示大于或等于比较操作。
    /// </summary>
    private const int GreaterThanOrEqualOperator = OperatorGreaterThanOrEqual;
    /// <summary>
    /// 表示小于或等于比较操作。
    /// </summary>
    private const int LessThanOrEqualOperator = OperatorLessThanOrEqual;

    /// <summary>
    /// NPOI 定义的两个边界之间操作码。
    /// </summary>
    private const int OperatorBetween = OperatorType.BETWEEN;
    /// <summary>
    /// NPOI 定义的两个边界之外操作码。
    /// </summary>
    private const int OperatorNotBetween = OperatorType.NOT_BETWEEN;
    /// <summary>
    /// NPOI 定义的相等操作码。
    /// </summary>
    private const int OperatorEqual = OperatorType.EQUAL;
    /// <summary>
    /// NPOI 定义的不相等操作码。
    /// </summary>
    private const int OperatorNotEqual = OperatorType.NOT_EQUAL;
    /// <summary>
    /// NPOI 定义的大于操作码。
    /// </summary>
    private const int OperatorGreaterThan = OperatorType.GREATER_THAN;
    /// <summary>
    /// NPOI 定义的小于操作码。
    /// </summary>
    private const int OperatorLessThan = OperatorType.LESS_THAN;
    /// <summary>
    /// NPOI 定义的大于或等于操作码。
    /// </summary>
    private const int OperatorGreaterThanOrEqual = OperatorType.GREATER_OR_EQUAL;
    /// <summary>
    /// NPOI 定义的小于或等于操作码。
    /// </summary>
    private const int OperatorLessThanOrEqual = OperatorType.LESS_OR_EQUAL;

    /// <summary>
    /// 使用原生操作码比较日期或时间刻度值。
    /// </summary>
    /// <param name="value">待比较的刻度值。</param>
    /// <param name="first">第一个边界值。</param>
    /// <param name="second">第二个边界值。</param>
    /// <param name="secondParsed">第二个边界是否成功解析。</param>
    /// <param name="operation">NPOI 数据校验比较操作码。</param>
    /// <returns>比较通过时为 true。</returns>
    private static bool Compare(long value, long first, long second, bool secondParsed, int operation) => operation switch
    {
        BetweenOperator => secondParsed && value >= first && value <= second,
        NotBetweenOperator => secondParsed && (value < first || value > second),
        EqualOperator => value == first,
        NotEqualOperator => value != first,
        GreaterThanOperator => value > first,
        LessThanOperator => value < first,
        GreaterThanOrEqualOperator => value >= first,
        LessThanOrEqualOperator => value <= first,
        _ => false,
    };

    /// <summary>
    /// 使用原生操作码比较数值或文本长度。
    /// </summary>
    /// <param name="value">待比较的值。</param>
    /// <param name="first">第一个边界值。</param>
    /// <param name="second">第二个边界值。</param>
    /// <param name="secondParsed">第二个边界是否成功解析。</param>
    /// <param name="operation">NPOI 数据校验比较操作码。</param>
    /// <returns>比较通过时为 true。</returns>
    private static bool Compare(decimal value, decimal first, decimal second, bool secondParsed, int operation) => operation switch
    {
        BetweenOperator => secondParsed && value >= first && value <= second,
        NotBetweenOperator => secondParsed && (value < first || value > second),
        EqualOperator => value == first,
        NotEqualOperator => value != first,
        GreaterThanOperator => value > first,
        LessThanOperator => value < first,
        GreaterThanOrEqualOperator => value >= first,
        LessThanOrEqualOperator => value <= first,
        _ => false
    };

    /// <summary>
    /// 返回统一的不支持原生校验规则错误。
    /// </summary>
    /// <param name="message">返回给调用方的错误消息。</param>
    /// <returns>始终为 false。</returns>
    private static bool Unsupported(out string message)
    {
        message = "Workbook Data Validation 规则类型或公式暂不支持。";
        return false;
    }
}

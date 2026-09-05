using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.Validations;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// 执行 Workbook 原生 Data Validation，隔离规则解析和错误收集。
/// </summary>
internal static class NpoiWorkbookValidationPipeline
{
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
    /// <param name="validateMode">发生校验失败后的继续策略。</param>
    /// <param name="unsupportedFeaturePolicy">不支持的原生校验规则的处理策略。</param>
    /// <param name="errors">接收工作簿校验错误的收集器。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <returns>当前行所有可执行规则均通过时为 true。</returns>
    public static bool Validate(IRow row, IReadOnlyDictionary<int, ExcelColumnPlan> columns,
        ValidationRangeIndex validationIndex, ISheet sheet, string sheetName, int rowIndex,
        ExcelWhitespacePolicy bodyWhitespace, ValidateMode validateMode,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy, ExcelImportErrorCollector errors,
        bool isDate1904)
    {
        var valid = true;
        foreach (var column in columns)
        {
            var cell = row.GetCell(column.Key);
            var cellValue = NpoiExcelImporter.ReadCellValue(cell, isDate1904);
            var value = NpoiExcelImporter.NormalizeText(cellValue.Text, bodyWhitespace);
            foreach (var validation in validationIndex.Get(rowIndex, column.Key))
            {
                if (ValidateValue(validation, cellValue, value, sheet, isDate1904, out var message))
                    continue;
                errors.Add(new ExcelImportError(ExcelImportErrorCode.WorkbookValidation, message, sheetName,
                    rowIndex + 1, column.Key + 1, column.Value.Property.Name, column.Value.DynamicDefinition?.Key
                    ?? column.Value.Property.Name, column.Value.HeaderName, value));
                if (unsupportedFeaturePolicy == ExcelUnsupportedFeaturePolicy.Report
                    && message != null && message.StartsWith("Workbook Data Validation 规则类型或公式暂不支持",
                        StringComparison.Ordinal))
                    continue;
                valid = false;
                if (validateMode == ValidateMode.StopOnFirstFailure)
                    return false;
            }
        }
        return valid;
    }

    /// <summary>按照单个原生数据校验约束验证已规范化的单元格值。</summary>
    /// <param name="validation">NPOI 提供的数据校验规则。</param>
    /// <param name="cellValue">保留原始类型信息的单元格值。</param>
    /// <param name="value">按正文空白策略规范化后的文本值。</param>
    /// <param name="sheet">解析列表区域时使用的当前工作表。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="message">校验失败或不支持时返回的说明消息。</param>
    /// <returns>规则通过时为 true。</returns>
    private static bool ValidateValue(IDataValidation validation, ExcelCellValue cellValue,
        string value, ISheet sheet, bool isDate1904, out string message)
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
            if (!TryGetExplicitListValues(constraint, sheet, out var values))
                return Unsupported(out message);
            if (values.Contains(value, StringComparer.Ordinal))
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

    /// <summary>从显式列表、带引号列表或单元格区域公式解析允许值。</summary>
    /// <param name="constraint">原生列表校验约束。</param>
    /// <param name="currentSheet">解析未限定区域时使用的工作表。</param>
    /// <param name="values">成功时返回允许值集合。</param>
    /// <returns>支持并成功解析列表规则时为 true。</returns>
    private static bool TryGetExplicitListValues(IDataValidationConstraint constraint, ISheet currentSheet,
        out string[] values)
    {
        var formula = constraint.Formula1?.Replace("&#34;", "\"").Replace("&quot;", "\"")
            .Replace("\\\"", "\"").Trim();
        if (!string.IsNullOrWhiteSpace(formula)
            && TryGetCellRangeValues(formula, currentSheet, out values))
            return true;
        if (string.IsNullOrWhiteSpace(formula))
        {
            values = null;
            return false;
        }
        values = constraint.ExplicitListValues;
        if (values != null && values.Length > 0)
        {
            if (values.Length == 1 && (values[0].Contains(":") || values[0].Contains("!")))
            {
                values = null;
                return false;
            }
            if (values.Length == 1 && values[0].Contains(","))
                values = SplitExplicitList(values[0]);
            return values.Length > 0;
        }
        if (formula.StartsWith("=", StringComparison.Ordinal))
            return false;
        if (!(formula.StartsWith("\"", StringComparison.Ordinal) && formula.EndsWith("\"",
                StringComparison.Ordinal)))
            return false;
        values = SplitExplicitList(formula.Trim('\"'));
        return values.Length > 0;
    }

    /// <summary>读取原生约束的第一个公式或 HSSF 数值后备值。</summary>
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

    /// <summary>读取原生约束的第二个公式或 HSSF 数值后备值。</summary>
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

    /// <summary>解析单工作表或带工作表限定的 A1 区域公式，并读取区域文本值。</summary>
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
        if (parts.Length != 2)
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
            var first = new NPOI.SS.Util.CellReference(parts[0].Trim());
            var last = new NPOI.SS.Util.CellReference(parts[1].Trim());
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

    /// <summary>拆分逗号分隔的显式列表并移除外层引号和空白。</summary>
    /// <param name="value">列表公式中的显式文本。</param>
    /// <returns>非空的列表项集合。</returns>
    private static string[] SplitExplicitList(string value) => value.Trim().Trim('\"')
        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
        .Select(item => item.Trim()).ToArray();

    /// <summary>从 typed 单元格值或文本解析 Excel 日期/时间为比较刻度。</summary>
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

    /// <summary>表示两个边界之间的比较操作。</summary>
    private const int BetweenOperator = OperatorBetween;
    /// <summary>表示两个边界之外的比较操作。</summary>
    private const int NotBetweenOperator = OperatorNotBetween;
    /// <summary>表示相等比较操作。</summary>
    private const int EqualOperator = OperatorEqual;
    /// <summary>表示不相等比较操作。</summary>
    private const int NotEqualOperator = OperatorNotEqual;
    /// <summary>表示大于比较操作。</summary>
    private const int GreaterThanOperator = OperatorGreaterThan;
    /// <summary>表示小于比较操作。</summary>
    private const int LessThanOperator = OperatorLessThan;
    /// <summary>表示大于或等于比较操作。</summary>
    private const int GreaterThanOrEqualOperator = OperatorGreaterThanOrEqual;
    /// <summary>表示小于或等于比较操作。</summary>
    private const int LessThanOrEqualOperator = OperatorLessThanOrEqual;

    /// <summary>NPOI 定义的两个边界之间操作码。</summary>
    private const int OperatorBetween = OperatorType.BETWEEN;
    /// <summary>NPOI 定义的两个边界之外操作码。</summary>
    private const int OperatorNotBetween = OperatorType.NOT_BETWEEN;
    /// <summary>NPOI 定义的相等操作码。</summary>
    private const int OperatorEqual = OperatorType.EQUAL;
    /// <summary>NPOI 定义的不相等操作码。</summary>
    private const int OperatorNotEqual = OperatorType.NOT_EQUAL;
    /// <summary>NPOI 定义的大于操作码。</summary>
    private const int OperatorGreaterThan = OperatorType.GREATER_THAN;
    /// <summary>NPOI 定义的小于操作码。</summary>
    private const int OperatorLessThan = OperatorType.LESS_THAN;
    /// <summary>NPOI 定义的大于或等于操作码。</summary>
    private const int OperatorGreaterThanOrEqual = OperatorType.GREATER_OR_EQUAL;
    /// <summary>NPOI 定义的小于或等于操作码。</summary>
    private const int OperatorLessThanOrEqual = OperatorType.LESS_OR_EQUAL;

    /// <summary>使用原生操作码比较日期或时间刻度值。</summary>
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

    /// <summary>使用原生操作码比较数值或文本长度。</summary>
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

    /// <summary>返回统一的不支持原生校验规则错误。</summary>
    /// <param name="message">返回给调用方的错误消息。</param>
    /// <returns>始终为 false。</returns>
    private static bool Unsupported(out string message)
    {
        message = "Workbook Data Validation 规则类型或公式暂不支持。";
        return false;
    }
}

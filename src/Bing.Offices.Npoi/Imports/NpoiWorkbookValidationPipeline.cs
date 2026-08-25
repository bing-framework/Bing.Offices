using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
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
    public static bool Validate(IRow row, IReadOnlyDictionary<int, ExcelColumnPlan> columns,
        ValidationRangeIndex validationIndex, ISheet sheet, string sheetName, int rowIndex,
        ExcelWhitespacePolicy bodyWhitespace, ValidateMode validateMode,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy, ExcelImportErrorCollector errors)
    {
        var valid = true;
        foreach (var column in columns)
        {
            var cell = row.GetCell(column.Key);
            var cellValue = NpoiExcelImporter.ReadCellValue(cell);
            var value = NpoiExcelImporter.NormalizeText(cellValue.Text, bodyWhitespace);
            foreach (var validation in validationIndex.Get(rowIndex, column.Key))
            {
                if (ValidateValue(validation.ValidationConstraint, cellValue, value, sheet, out var message))
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

    private static bool ValidateValue(IDataValidationConstraint constraint, ExcelCellValue cellValue,
        string value, ISheet sheet, out string message)
    {
        value ??= string.Empty;
        var type = constraint.GetValidationType();
        if (type == NPOI.SS.UserModel.ValidationType.ANY)
        {
            message = null;
            return true;
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
            if (!decimal.TryParse(constraint.Formula1, NumberStyles.Float, CultureInfo.InvariantCulture,
                    out var first))
                return Unsupported(out message);
            var secondParsed = decimal.TryParse(constraint.Formula2, NumberStyles.Float,
                CultureInfo.InvariantCulture, out var second);
            var valid = Compare(value.Length, first, second, secondParsed, constraint.Operator);
            message = valid ? null : "不符合 Workbook 文本长度校验。";
            return valid;
        }
        if (type == NPOI.SS.UserModel.ValidationType.INTEGER || type == NPOI.SS.UserModel.ValidationType.DECIMAL)
        {
            var parsed = decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture,
                out var number);
            var firstParsed = decimal.TryParse(constraint.Formula1, NumberStyles.Number,
                CultureInfo.InvariantCulture, out var first);
            var secondParsed = decimal.TryParse(constraint.Formula2, NumberStyles.Number,
                CultureInfo.InvariantCulture, out var second);
            var valid = parsed && firstParsed && Compare(number, first, second, secondParsed,
                constraint.Operator);
            message = valid ? null : "不符合 Workbook 数值校验。";
            return valid;
        }
        if (type == NPOI.SS.UserModel.ValidationType.DATE || type == NPOI.SS.UserModel.ValidationType.TIME)
        {
            var parsed = TryGetExcelDate(cellValue, value, type == NPOI.SS.UserModel.ValidationType.TIME, out var date);
            var first = TryGetExcelDate(constraint.Formula1, type == NPOI.SS.UserModel.ValidationType.TIME,
                out var minimum);
            var second = TryGetExcelDate(constraint.Formula2, type == NPOI.SS.UserModel.ValidationType.TIME,
                out var maximum);
            var valid = parsed && first && Compare(date, minimum, maximum, second, constraint.Operator);
            message = valid ? null : "不符合 Workbook 日期/时间校验。";
            return valid;
        }
        return Unsupported(out message);
    }

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

    private static string[] SplitExplicitList(string value) => value.Trim().Trim('\"')
        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
        .Select(item => item.Trim()).ToArray();

    private static bool TryGetExcelDate(ExcelCellValue cellValue, string text, bool timeOnly, out long ticks)
    {
        if (cellValue?.Value is DateTime date)
        {
            ticks = timeOnly ? date.TimeOfDay.Ticks : date.Ticks;
            return true;
        }
        if (cellValue?.Value is double number)
        {
            ticks = ExcelSerialToDateTime(number, timeOnly).Ticks;
            return true;
        }
        return TryGetExcelDate(text, timeOnly, out ticks);
    }

    private static bool TryGetExcelDate(string text, bool timeOnly, out long ticks)
    {
        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces,
                out var date))
        {
            ticks = timeOnly ? date.TimeOfDay.Ticks : date.Ticks;
            return true;
        }
        if (decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
        {
            ticks = ExcelSerialToDateTime((double)serial, timeOnly).Ticks;
            return true;
        }
        ticks = 0;
        return false;
    }

    private static DateTime ExcelSerialToDateTime(double serial, bool timeOnly)
    {
        var date = new DateTime(1899, 12, 30).AddDays(serial);
        return timeOnly ? DateTime.Today.Add(date.TimeOfDay) : date;
    }

    private static bool Compare(long value, long first, long second, bool secondParsed, int operation) => operation switch
    {
        0 => secondParsed && value >= first && value <= second,
        1 => secondParsed && (value < first || value > second),
        2 => value == first,
        3 => value != first,
        4 => value > first,
        5 => value < first,
        6 => value >= first,
        7 => secondParsed && value <= second,
        _ => false,
    };

    private static bool Compare(decimal value, decimal first, decimal second, bool secondParsed, int operation) => operation switch
    {
        0 => secondParsed && value >= first && value <= second,
        1 => secondParsed && (value < first || value > second),
        2 => value == first,
        3 => value != first,
        4 => value > first,
        5 => value < first,
        6 => value >= first,
        7 => secondParsed && value <= second,
        _ => false
    };

    private static bool Unsupported(out string message)
    {
        message = "Workbook Data Validation 规则类型或公式暂不支持。";
        return false;
    }
}

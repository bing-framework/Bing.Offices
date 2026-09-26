using ClosedXML.Excel;

namespace Bing.Offices.Exports;

/// <summary>
/// 写入公共图片和原生数据校验定义。
/// </summary>
internal static class ClosedXmlSheetContentWriter
{
    /// <summary>
    /// 将图片和数据校验定义应用到工作表。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="request">当前操作请求。</param>
    /// <param name="token">取消令牌。</param>
    internal static void Apply(XLWorkbook workbook, IXLWorksheet sheet, ExcelSheetExportRequest request, CancellationToken token)
    {
        foreach (var definition in request.Images)
        {
            token.ThrowIfCancellationRequested();
            using var input = new MemoryStream(definition.Content, writable: false);
            sheet.AddPicture(input).MoveTo(sheet.Cell(definition.Row + 1, definition.Column + 1),
                definition.OffsetX, definition.OffsetY).WithSize(definition.Width, definition.Height);
        }
        foreach (var definition in request.DataValidations)
        {
            token.ThrowIfCancellationRequested();
            var rule = sheet.Range(ExcelSheetContent.Address(definition.Range)).CreateDataValidation();
            rule.AllowedValues = definition.Type switch
            {
                ExcelDataValidationType.Integer => XLAllowedValues.WholeNumber,
                ExcelDataValidationType.Decimal => XLAllowedValues.Decimal,
                ExcelDataValidationType.Date => XLAllowedValues.Date,
                ExcelDataValidationType.TextLength => XLAllowedValues.TextLength,
                _ => XLAllowedValues.List
            };
            rule.Operator = definition.Operator switch
            {
                ExcelConditionalComparisonOperator.Equal => XLOperator.EqualTo,
                ExcelConditionalComparisonOperator.NotEqual => XLOperator.NotEqualTo,
                ExcelConditionalComparisonOperator.GreaterThan => XLOperator.GreaterThan,
                ExcelConditionalComparisonOperator.GreaterThanOrEqual => XLOperator.EqualOrGreaterThan,
                ExcelConditionalComparisonOperator.LessThan => XLOperator.LessThan,
                ExcelConditionalComparisonOperator.LessThanOrEqual => XLOperator.EqualOrLessThan,
                ExcelConditionalComparisonOperator.Between => XLOperator.Between,
                _ => XLOperator.NotBetween
            };
            if (definition.Type == ExcelDataValidationType.List)
                rule.List("\"" + string.Join(",", definition.Values) + "\"", true);
            else
            {
                rule.MinValue = ExcelSheetContent.Value(definition, false, workbook.Use1904DateSystem);
                if (definition.Operator == ExcelConditionalComparisonOperator.Between
                    || definition.Operator == ExcelConditionalComparisonOperator.NotBetween)
                    rule.MaxValue = ExcelSheetContent.Value(definition, true, workbook.Use1904DateSystem);
            }
            rule.IgnoreBlanks = definition.IgnoreBlanks;
            rule.ShowErrorMessage = definition.ShowErrorMessage;
            rule.ErrorStyle = XLErrorStyle.Stop;
            rule.ShowInputMessage = !string.IsNullOrEmpty(definition.InputTitle) || !string.IsNullOrEmpty(definition.InputMessage);
            rule.InputTitle = definition.InputTitle ?? string.Empty;
            rule.InputMessage = definition.InputMessage ?? string.Empty;
            rule.ErrorTitle = definition.ErrorTitle ?? string.Empty;
            rule.ErrorMessage = definition.ErrorMessage ?? string.Empty;
        }
    }
}

using ClosedXML.Excel;

using Bing.Offices.Exceptions;
using Bing.Offices.ClosedXml.Imports;

namespace Bing.Offices.Exports;

/// <summary>
/// 写入公共图片和原生数据校验定义。
/// </summary>
internal static class ClosedXmlSheetContentWriter
{
    /// <summary>
    /// 将图片定义应用到工作表。
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
    }

    /// <summary>
    /// 将原生校验定义应用到工作表。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="definitions">校验定义。</param>
    /// <param name="token">取消令牌。</param>
    /// <param name="cascadeMaps">级联字典与映射名称。</param>
    internal static void ApplyValidations(XLWorkbook workbook, IXLWorksheet sheet,
        IReadOnlyList<ExcelDataValidationDefinition> definitions, CancellationToken token,
        IReadOnlyDictionary<string, string> cascadeMaps = null)
    {
        foreach (var definition in definitions)
        {
            token.ThrowIfCancellationRequested();
            if (definition.ListSourceName != null)
            {
                if (!ClosedXmlWorkbookValidationPipeline.TryResolveNamedRange(sheet,
                        definition.ListSourceName, out var range))
                    throw new BingOfficesConfigurationException(
                        $"列表校验名称缺失或来源不是受支持的一维有界区域: {definition.ListSourceName}",
                        stage: BingOfficesStage.Plan);
                foreach (var cell in range.CellsUsed(XLCellsUsedOptions.Contents))
                {
                    token.ThrowIfCancellationRequested();
                    if (cell.HasFormula)
                        throw new BingOfficesConfigurationException(
                            $"列表校验名称来源不能包含公式单元格: {definition.ListSourceName}",
                            stage: BingOfficesStage.Plan);
                }
            }
            var cascadeFormula = definition.CascadeSourceName == null
                ? null : CascadeFormula(definition, cascadeMaps);
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
            {
                var source = cascadeFormula != null
                    ? cascadeFormula
                    : definition.ListSourceName == null
                        ? "\"" + string.Join(",", definition.Values) + "\""
                        : definition.ListSourceName;
                rule.List(source, true);
            }
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

    /// <summary>
    /// 生成按同一行上级单元格取值的下拉公式。
    /// </summary>
    /// <param name="definition">校验定义。</param>
    /// <param name="cascadeMaps">级联字典与映射名称。</param>
    /// <returns>下拉来源公式。</returns>
    private static string CascadeFormula(ExcelDataValidationDefinition definition,
        IReadOnlyDictionary<string, string> cascadeMaps)
    {
        if (cascadeMaps == null || !cascadeMaps.TryGetValue(definition.CascadeSourceName, out var mapName))
            throw new BingOfficesConfigurationException(
                $"级联校验名称缺失: {definition.CascadeSourceName}", stage: BingOfficesStage.Plan);
        if (definition.ParentColumnIndex >= 16384 || definition.Range.EndColumn >= 16384
            || definition.Range.EndRow >= 1048576)
            throw new BingOfficesConfigurationException("级联列表区域超过工作表格式边界。",
                stage: BingOfficesStage.Plan);
        var column = definition.ParentColumnIndex.Value + 1;
        var letters = string.Empty;
        while (column > 0)
        {
            column--;
            letters = (char)('A' + column % 26) + letters;
            column /= 26;
        }
        var formula = $"INDIRECT(VLOOKUP(${letters}{definition.Range.StartRow + 1},{mapName},2,FALSE))";
        if (formula.Length > 255)
            throw new BingOfficesConfigurationException("级联列表公式超过长度上限。",
                stage: BingOfficesStage.Plan);
        return formula;
    }
}

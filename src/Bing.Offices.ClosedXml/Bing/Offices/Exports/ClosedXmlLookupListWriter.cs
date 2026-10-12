using Bing.Offices.Exceptions;
using ClosedXML.Excel;

namespace Bing.Offices.Exports;

/// <summary>
/// 为 ClosedXML 工作簿写入隐藏文本字典。
/// </summary>
internal static class ClosedXmlLookupListWriter
{
    /// <summary>
    /// 写入字典工作表和工作簿级名称。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="definitions">字典定义。</param>
    /// <param name="token">取消令牌。</param>
    /// <param name="reservedSheetNames">稍后创建的工作表名称。</param>
    internal static void Apply(XLWorkbook workbook, IReadOnlyList<ExcelLookupListDefinition> definitions,
        CancellationToken token, IReadOnlyCollection<string> reservedSheetNames = null)
    {
        var nextIndex = 1;
        foreach (var definition in definitions)
        {
            token.ThrowIfCancellationRequested();
            if (definition.Values.Count > 1048576)
                throw new BingOfficesConfigurationException("字典项超过工作表行数上限。",
                    stage: BingOfficesStage.Plan);
            if (workbook.DefinedNames.Contains(definition.Name)
                || workbook.Worksheets.Any(sheet => sheet.DefinedNames.Contains(definition.Name)))
                throw new BingOfficesConfigurationException($"工作簿字典名称已存在: {definition.Name}",
                    stage: BingOfficesStage.Plan);
            string sheetName;
            do
            {
                sheetName = "__BingLookup_" + nextIndex.ToString("D3", System.Globalization.CultureInfo.InvariantCulture);
                nextIndex++;
            } while (workbook.Worksheets.Any(sheet => string.Equals(sheet.Name, sheetName,
                StringComparison.OrdinalIgnoreCase))
                || reservedSheetNames?.Contains(sheetName, StringComparer.OrdinalIgnoreCase) == true);
            var sheet = workbook.Worksheets.Add(sheetName);
            for (var row = 0; row < definition.Values.Count; row++)
            {
                token.ThrowIfCancellationRequested();
                sheet.Cell(row + 1, 1).Value = definition.Values[row];
            }
            workbook.DefinedNames.Add(definition.Name, sheet.Range(1, 1, definition.Values.Count, 1));
            sheet.Visibility = XLWorksheetVisibility.Hidden;
        }
    }
}

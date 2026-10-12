using Bing.Offices.Exceptions;
using NPOI.SS.UserModel;

namespace Bing.Offices.Exports;

/// <summary>
/// 为 NPOI 工作簿写入隐藏文本字典。
/// </summary>
internal static class NpoiLookupListWriter
{
    /// <summary>
    /// 写入字典工作表和工作簿级名称。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="definitions">字典定义。</param>
    /// <param name="token">取消令牌。</param>
    /// <param name="reservedSheetNames">为后续报表保留的工作表名称；为空时不额外保留。</param>
    internal static void Apply(IWorkbook workbook, IReadOnlyList<ExcelLookupListDefinition> definitions,
        CancellationToken token, IReadOnlyCollection<string> reservedSheetNames = null)
    {
        var reserved = new HashSet<string>(reservedSheetNames ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
        var maxRows = workbook is NPOI.HSSF.UserModel.HSSFWorkbook ? 65536 : 1048576;
        var nextIndex = 1;
        foreach (var definition in definitions)
        {
            token.ThrowIfCancellationRequested();
            if (definition.Values.Count > maxRows)
                throw new BingOfficesConfigurationException("字典项超过工作表行数上限。",
                    stage: BingOfficesStage.Plan);
            for (var index = 0; index < workbook.NumberOfNames; index++)
                if (string.Equals(workbook.GetNameAt(index).NameName, definition.Name,
                        StringComparison.OrdinalIgnoreCase))
                    throw new BingOfficesConfigurationException($"工作簿字典名称已存在: {definition.Name}",
                        stage: BingOfficesStage.Plan);
            string sheetName;
            do
            {
                sheetName = "__BingLookup_" + nextIndex.ToString("D3", System.Globalization.CultureInfo.InvariantCulture);
                nextIndex++;
            } while (workbook.GetSheet(sheetName) != null || reserved.Contains(sheetName));
            var sheet = workbook.CreateSheet(sheetName);
            for (var row = 0; row < definition.Values.Count; row++)
            {
                token.ThrowIfCancellationRequested();
                sheet.CreateRow(row).CreateCell(0, CellType.String).SetCellValue(definition.Values[row]);
            }
            var name = workbook.CreateName();
            name.NameName = definition.Name;
            var lastRow = definition.Values.Count;
            name.RefersToFormula = sheetName + "!$A$1:$A$" + lastRow;
            workbook.SetSheetVisibility(workbook.GetSheetIndex(sheet), SheetVisibility.Hidden);
        }
    }
}

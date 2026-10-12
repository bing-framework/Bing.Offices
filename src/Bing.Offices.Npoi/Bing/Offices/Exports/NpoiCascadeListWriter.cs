using Bing.Offices.Exceptions;
using NPOI.SS.UserModel;

namespace Bing.Offices.Exports;

/// <summary>
/// 写入 NPOI 工作簿的级联文本字典。
/// </summary>
internal static class NpoiCascadeListWriter
{
    /// <summary>
    /// 写入隐藏字典并返回级联名称与映射名称的对应关系。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="definitions">级联字典定义。</param>
    /// <param name="token">取消令牌。</param>
    /// <param name="reservedSheetNames">为后续报表保留的工作表名称；为空时不额外保留。</param>
    /// <returns>级联字典名称与映射名称的对应关系。</returns>
    internal static IReadOnlyDictionary<string, string> Apply(IWorkbook workbook,
        IReadOnlyList<ExcelCascadeListDefinition> definitions, CancellationToken token,
        IReadOnlyCollection<string> reservedSheetNames = null)
    {
        var reserved = new HashSet<string>(reservedSheetNames ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
        var maps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < workbook.NumberOfNames; index++)
            names.Add(workbook.GetNameAt(index).NameName);
        foreach (var definition in definitions)
            if (!names.Add(definition.Name))
                throw new BingOfficesConfigurationException($"工作簿级联名称已存在: {definition.Name}",
                    stage: BingOfficesStage.Plan);
        var maxRows = workbook is NPOI.HSSF.UserModel.HSSFWorkbook ? 65536 : 1048576;
        var nextIndex = 1;
        var nextMapIndex = 1;
        var nextChildIndex = 1;
        foreach (var definition in definitions)
        {
            token.ThrowIfCancellationRequested();
            var childCount = definition.Entries.Sum(entry => (long)entry.Values.Count);
            if (definition.Entries.Count > maxRows || childCount > maxRows)
                throw new BingOfficesConfigurationException("级联字典项超过工作表行数上限。",
                    stage: BingOfficesStage.Plan);
            string sheetName;
            do
            {
                sheetName = "__BingCascade_" + nextIndex.ToString("D3",
                    System.Globalization.CultureInfo.InvariantCulture);
                nextIndex++;
            } while (workbook.GetSheet(sheetName) != null || reserved.Contains(sheetName));
            var sheet = workbook.CreateSheet(sheetName);
            var mapName = NextName(names, "__BingCascadeMap_", ref nextMapIndex);
            var childRow = 1;
            for (var index = 0; index < definition.Entries.Count; index++)
            {
                token.ThrowIfCancellationRequested();
                var entry = definition.Entries[index];
                var childName = NextName(names, "__BingCascadeValues_", ref nextChildIndex);
                var row = sheet.GetRow(index) ?? sheet.CreateRow(index);
                row.CreateCell(0, CellType.String).SetCellValue(entry.Parent);
                row.CreateCell(1, CellType.String).SetCellValue(childName);
                var firstChildRow = childRow;
                foreach (var value in entry.Values)
                {
                    token.ThrowIfCancellationRequested();
                    var child = sheet.GetRow(childRow - 1) ?? sheet.CreateRow(childRow - 1);
                    child.CreateCell(2, CellType.String).SetCellValue(value);
                    childRow++;
                }
                AddName(workbook, childName, sheetName, "$C$" + firstChildRow + ":$C$" + (childRow - 1));
            }
            AddName(workbook, definition.Name, sheetName, "$A$1:$A$" + definition.Entries.Count);
            AddName(workbook, mapName, sheetName, "$A$1:$B$" + definition.Entries.Count);
            workbook.SetSheetVisibility(workbook.GetSheetIndex(sheet), SheetVisibility.Hidden);
            maps.Add(definition.Name, mapName);
        }
        return maps;
    }

    /// <summary>
    /// 获取未占用的内部名称。
    /// </summary>
    /// <param name="names">已占用的名称集合。</param>
    /// <param name="prefix">名称前缀。</param>
    /// <param name="nextIndex">下一个编号。</param>
    /// <returns>生成的名称。</returns>
    private static string NextName(HashSet<string> names, string prefix, ref int nextIndex)
    {
        string name;
        do
        {
            name = prefix + nextIndex.ToString("D3", System.Globalization.CultureInfo.InvariantCulture);
            nextIndex++;
        } while (!names.Add(name));
        return name;
    }

    /// <summary>
    /// 创建工作簿级名称。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="name">待创建的名称。</param>
    /// <param name="sheet">名称引用的工作表。</param>
    /// <param name="address">名称引用的区域地址。</param>
    private static void AddName(IWorkbook workbook, string name, string sheet, string address)
    {
        var range = workbook.CreateName();
        range.NameName = name;
        range.RefersToFormula = sheet + "!" + address;
    }
}

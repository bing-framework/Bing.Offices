using System.Globalization;
using Bing.Offices.Exceptions;
using ClosedXML.Excel;

namespace Bing.Offices.Exports;

/// <summary>
/// 写入 ClosedXML 级联字典及其工作簿级名称。
/// </summary>
internal static class ClosedXmlCascadeListWriter
{
    /// <summary>
    /// 写入级联字典并返回下级校验所需的映射名称。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="definitions">级联字典定义。</param>
    /// <param name="token">取消令牌。</param>
    /// <param name="reservedSheetNames">稍后创建的工作表名称。</param>
    /// <returns>级联字典名称与映射名称的对应关系。</returns>
    internal static IReadOnlyDictionary<string, string> Apply(XLWorkbook workbook,
        IReadOnlyList<ExcelCascadeListDefinition> definitions, CancellationToken token,
        IReadOnlyCollection<string> reservedSheetNames = null)
    {
        var maps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var reservedNames = new HashSet<string>(definitions.Select(item => item.Name),
            StringComparer.OrdinalIgnoreCase);
        var nextSheetIndex = 1;
        var nextNameIndex = 1;
        foreach (var definition in definitions)
        {
            token.ThrowIfCancellationRequested();
            if (HasName(workbook, definition.Name))
                throw new BingOfficesConfigurationException($"工作簿字典名称已存在: {definition.Name}",
                    stage: BingOfficesStage.Plan);
            if (definition.Entries.Count > 1048576
                || definition.Entries.Sum(entry => (long)entry.Values.Count) > 1048576)
                throw new BingOfficesConfigurationException("级联字典项超过工作表行数上限。",
                    stage: BingOfficesStage.Plan);

            string sheetName;
            do
            {
                sheetName = "__BingCascade_" + nextSheetIndex.ToString("D3", CultureInfo.InvariantCulture);
                nextSheetIndex++;
            } while (workbook.Worksheets.Any(sheet => string.Equals(sheet.Name, sheetName,
                StringComparison.OrdinalIgnoreCase))
                || reservedSheetNames?.Contains(sheetName, StringComparer.OrdinalIgnoreCase) == true);
            var sheet = workbook.Worksheets.Add(sheetName);
            workbook.DefinedNames.Add(definition.Name, sheet.Range(1, 1, definition.Entries.Count, 1));
            var mapName = NextName(workbook, "__BingCascadeMap_", reservedNames, ref nextNameIndex);
            var childRow = 1;
            for (var index = 0; index < definition.Entries.Count; index++)
            {
                token.ThrowIfCancellationRequested();
                var entry = definition.Entries[index];
                var childName = NextName(workbook, "__BingCascadeValues_", reservedNames, ref nextNameIndex);
                sheet.Cell(index + 1, 1).Value = entry.Parent;
                sheet.Cell(index + 1, 2).Value = childName;
                foreach (var value in entry.Values)
                {
                    token.ThrowIfCancellationRequested();
                    sheet.Cell(childRow++, 3).Value = value;
                }
                workbook.DefinedNames.Add(childName,
                    sheet.Range(childRow - entry.Values.Count, 3, childRow - 1, 3));
            }
            workbook.DefinedNames.Add(mapName, sheet.Range(1, 1, definition.Entries.Count, 2));
            sheet.Visibility = XLWorksheetVisibility.Hidden;
            maps.Add(definition.Name, mapName);
        }
        return maps;
    }

    /// <summary>
    /// 获取未占用的工作簿名称。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="prefix">名称前缀。</param>
    /// <param name="reservedNames">请求保留的字典名称。</param>
    /// <param name="nextIndex">下一个编号。</param>
    /// <returns>生成的名称。</returns>
    private static string NextName(XLWorkbook workbook, string prefix,
        ISet<string> reservedNames, ref int nextIndex)
    {
        string name;
        do
        {
            name = prefix + nextIndex.ToString("D3", CultureInfo.InvariantCulture);
            nextIndex++;
        } while (reservedNames.Contains(name) || HasName(workbook, name));
        return name;
    }

    /// <summary>
    /// 判断名称是否已在工作簿或工作表作用域使用。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="name">待查名称。</param>
    /// <returns>已占用时为 <see langword="true" />。</returns>
    private static bool HasName(XLWorkbook workbook, string name) =>
        workbook.DefinedNames.Contains(name)
        || workbook.Worksheets.Any(sheet => sheet.DefinedNames.Contains(name));
}

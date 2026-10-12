using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Exports;

/// <summary>
/// 将工作表保护及可填写区域写入 NPOI 工作簿。
/// </summary>
internal static class NpoiSheetProtectionWriter
{
    /// <summary>
    /// 应用工作表保护。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="definition">工作表保护定义。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    internal static void Apply(IWorkbook workbook, ISheet sheet,
        ExcelSheetProtectionDefinition definition, CancellationToken cancellationToken)
    {
        var isXls = workbook is NPOI.HSSF.UserModel.HSSFWorkbook;
        var maxRow = isXls ? 65535 : 1048575;
        var maxColumn = isXls ? 255 : 16383;
        foreach (var range in definition.EditableRanges)
            if (range.EndRow > maxRow || range.EndColumn > maxColumn)
                throw new BingOfficesConfigurationException("工作表保护区域超出格式边界。",
                    stage: BingOfficesStage.Preflight);

        var unlockedStyles = new Dictionary<short, ICellStyle>();
        foreach (var range in definition.EditableRanges)
            for (var rowIndex = range.StartRow; rowIndex <= range.EndRow; rowIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                for (var columnIndex = range.StartColumn; columnIndex <= range.EndColumn; columnIndex++)
                {
                    var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
                    var style = cell.CellStyle;
                    if (!style.IsLocked)
                        continue;
                    if (!unlockedStyles.TryGetValue(style.Index, out var unlockedStyle))
                    {
                        unlockedStyle = workbook.CreateCellStyle();
                        unlockedStyle.CloneStyleFrom(style);
                        unlockedStyle.IsLocked = false;
                        unlockedStyles.Add(style.Index, unlockedStyle);
                    }
                    cell.CellStyle = unlockedStyle;
                }
            }
        cancellationToken.ThrowIfCancellationRequested();
        sheet.ProtectSheet(definition.Password ?? string.Empty);
    }
}

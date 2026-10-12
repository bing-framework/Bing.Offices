using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Exports;

/// <summary>
/// 将工作表保护及可填写区域写入 ClosedXML 工作簿。
/// </summary>
internal static class ClosedXmlSheetProtectionWriter
{
    /// <summary>
    /// 应用工作表保护。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="definition">工作表保护定义。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    internal static void Apply(IXLWorksheet sheet, ExcelSheetProtectionDefinition definition,
        CancellationToken cancellationToken)
    {
        foreach (var range in definition.EditableRanges)
            if (range.EndRow > 1048575 || range.EndColumn > 16383)
                throw new BingOfficesConfigurationException("工作表保护区域超出 XLSX 格式边界。",
                    stage: BingOfficesStage.Preflight);

        foreach (var range in definition.EditableRanges)
        {
            cancellationToken.ThrowIfCancellationRequested();
            sheet.Range(range.StartRow + 1, range.StartColumn + 1,
                range.EndRow + 1, range.EndColumn + 1).Style.Protection.Locked = false;
        }
        cancellationToken.ThrowIfCancellationRequested();
        sheet.Protect(definition.Password ?? string.Empty);
    }
}

using Bing.Offices.Exports;
using Bing.Offices.Npoi.Extensions;

namespace Bing.Offices.Npoi.Exports;

/// <summary>
/// Excel 导出器
/// </summary>
public class ExcelExporter : ExcelExporterBase
{
    /// <summary>
    /// 导出
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <param name="dataItems">数据</param>
    /// <param name="context">导出上下文</param>
    protected override Task<byte[]> ExportAsBytesAsync<TEntity>(ICollection<TEntity> dataItems, ExportContext<TEntity> context)
    {
        var workbook = ExcelHelper.PrepareWorkbook(context.Format, context.ExcelSetting);
        // 多Sheet设置
        if (context.GlobalSheetSetting.MaxRowNumberOnASheet > 0 && dataItems.Count > context.GlobalSheetSetting.MaxRowNumberOnASheet)
        {
            var sheetCount = (int)(dataItems.Count / context.GlobalSheetSetting.MaxRowNumberOnASheet) +
                             ((dataItems.Count % context.GlobalSheetSetting.MaxRowNumberOnASheet) > 0 ? 1 : 0);
            for (var i = 0; i < sheetCount; i++)
            {
                var sheetDataItems = dataItems.Skip(i * context.GlobalSheetSetting.MaxRowNumberOnASheet)
                    .Take(context.GlobalSheetSetting.MaxRowNumberOnASheet)
                    .ToList();
            }
        }
        else
        {
            var sheet = workbook.CreateSheet();

        }
        return Task.FromResult(workbook.SaveToBuffer());
    }
}

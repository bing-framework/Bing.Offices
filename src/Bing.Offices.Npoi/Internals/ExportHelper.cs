using System.Diagnostics;
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Resolvers;
using Bing.Reflection;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Internals;

/// <summary>
/// 导出帮助类
/// </summary>
internal static class ExportHelper
{
    public static ISheet EntitiesToSheet<TEntity>(ISheet sheet, ICollection<TEntity> dataItems, ExportContext<TEntity> context)
    {
        CreateHeader(sheet, context);
        CreateBody(sheet, dataItems, context);
        if (dataItems.Any())
            PostSheetProcess(sheet, context);
        return sheet;
    }

    /// <summary>
    /// 创建表头
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <param name="sheet">工作表</param>
    /// <param name="context">导出上下文</param>
    private static void CreateHeader<TEntity>(ISheet sheet, ExportContext<TEntity> context)
    {
        var sheetSetting = context.GetDefaultSheetSetting();
        // 获取表头行
        var headerRow = sheet.GetRow(sheetSetting.HeaderRowIndex);
        var columnIndex = 0;
        // 默认样式
        var defaultStyle = CreateCellStyle(sheet.Workbook, sheetSetting.HeaderStyle);
        foreach (var kvp in context.PropertySettings)
        {
            if (kvp.Value.Ignored)
                continue;
            var cell = headerRow.GetOrCreateCell(sheetSetting.AutoIndex ? columnIndex : kvp.Value.ColumnIndex);
            var currentStyle = kvp.Value.HeaderStyle == null
                ? defaultStyle
                : CreateCellStyle(sheet.Workbook, kvp.Value.HeaderStyle);
            cell.SetCellValue(kvp.Value.Name);
            cell.CellStyle = currentStyle;
            columnIndex++;
        }
    }

    /// <summary>
    /// 创建单元格样式
    /// </summary>
    /// <param name="workbook">工作簿</param>
    /// <param name="baseStyle">基础样式</param>
    /// <returns>NPOI单元格样式</returns>
    private static ICellStyle CreateCellStyle(IWorkbook workbook, IBaseStyle baseStyle)
    {
        var style = workbook.CreateCellStyle();
        var font = workbook.CreateFont();
        font.IsBold = baseStyle.IsBold;
        font.FontName = baseStyle.FontName;
        font.Color = baseStyle.FontColor;
        font.FontHeightInPoints = baseStyle.FontSize;
        style.SetFont(font);
        style.SetFillForegroundColor(baseStyle.FillForegroundColor);
        style.SetHorizontalAlignment(AlignmentResolver.Resolve(baseStyle.HorizontalAlignment));
        style.SetVerticalAlignment(AlignmentResolver.Resolve(baseStyle.VerticalAlignment));
        style.WrapText = baseStyle.WrapText;
        return style;
    }

    /// <summary>
    /// 创建正文
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <param name="sheet">工作表</param>
    /// <param name="dataItems">数据</param>
    /// <param name="context">导出上下文</param>
    private static void CreateBody<TEntity>(ISheet sheet, ICollection<TEntity> dataItems, ExportContext<TEntity> context)
    {
        if (dataItems == null || dataItems.Count <= 0)
            return;
        var sheetSetting = context.GetDefaultSheetSetting();
        var propertySettings = context.PropertySettings;
        var rowIndex = 0;

        // 默认样式
        var defaultStyle = CreateCellStyle(sheet.Workbook, sheetSetting.ColumnStyle);
        foreach (var item in dataItems)
        {
            var row = sheet.CreateRow(sheetSetting.HeaderRowIndex + rowIndex);
            var columnIndex = 0;
            if (item != null)
            {
                foreach (var key in propertySettings.Keys)
                {
                    var propertyValue = key.GetValueGetter()?.Invoke(item);
                    var outputFormatterFunc = context.GetOutputFormatterFunc(key);
                    if (outputFormatterFunc != null)
                    {
                        try
                        {
                            propertyValue = outputFormatterFunc.DynamicInvoke(item, propertyValue);
                        }
                        catch (Exception e)
                        {
                            Debug.WriteLine(e);
                            InvokeHelper.OnInvokeException?.Invoke(e);
                        }
                    }

                    var cell = row.GetOrCreateCell(sheetSetting.AutoIndex
                        ? columnIndex
                        : propertySettings[key].ColumnIndex);
                    cell.SetCellValue(propertyValue, propertySettings[key].StringFormat);
                    var currentStyle = propertySettings[key].ColumnStyle == null
                        ? defaultStyle
                        : CreateCellStyle(sheet.Workbook, propertySettings[key].ColumnStyle);
                    cell.CellStyle = currentStyle;
                    columnIndex++;
                }
            }

            rowIndex++;
        }
    }

    /// <summary>
    /// 后续工作表处理
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <param name="sheet">工作表</param>
    /// <param name="context">导出上下文</param>
    private static void PostSheetProcess<TEntity>(ISheet sheet, ExportContext<TEntity> context)
    {
        var sheetSetting = context.GetDefaultSheetSetting();
        var propertySettings = context.PropertySettings;
        var columnIndex = 0;
        // 设置列宽
        foreach (var setting in propertySettings.Values)
        {
            var currentIndex = sheetSetting.AutoIndex ? columnIndex : setting.ColumnIndex;
            if (setting.ColumnWidth > 0)
            {
                sheet.SetColumnWidth(currentIndex, setting.ColumnWidth * 256);
            }
            else
            {
                if (sheetSetting.AutoColumnWidthEnabled)
                    sheet.AutoSizeColumn(currentIndex);
            }
            columnIndex++;
        }

        // 设置冻结窗格
        foreach (var freezeSetting in context.FreezeSettings)
            sheet.CreateFreezePane(freezeSetting.ColSplit, freezeSetting.RowSplit, freezeSetting.LeftMostColumn, freezeSetting.TopRow);
    }
}

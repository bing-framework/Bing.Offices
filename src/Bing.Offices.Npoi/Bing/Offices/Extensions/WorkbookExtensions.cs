using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// NPOI 工作簿扩展。
/// </summary>
public static class WorkbookExtensions
{
    #region GetExcelFormat(获取Excel格式类型)

    /// <summary>
    /// 获取工作簿的 Excel 文件格式；仅支持 HSSF/XLS 和 XSSF/XLSX 实现。
    /// </summary>
    /// <param name="workbook">工作簿</param>
    /// <returns>工作簿对应的 XLS 或 XLSX 格式。</returns>
    public static ExcelFormat GetExcelFormat(this IWorkbook workbook)
    {
        switch (workbook)
        {
            case HSSFWorkbook _:
                return ExcelFormat.Xls;
            case XSSFWorkbook _:
                return ExcelFormat.Xlsx;
        }
        throw new NotSupportedException("未知 Excel 格式类型。");
    }

    #endregion

    #region GetSheets(获取工作表集合)

    /// <summary>
    /// 获取工作簿中可见且非非常隐藏的工作表集合。
    /// </summary>
    /// <param name="workbook">工作簿</param>
    /// <returns>按工作簿顺序排列的可见工作表；没有可见工作表时返回空集合。</returns>
    public static IEnumerable<ISheet> GetSheets(this IWorkbook workbook)
    {
        var sheets = new List<ISheet>();
        for (var i = 0; i < workbook.NumberOfSheets; i++)
        {
            var sheet = workbook.GetSheetAt(i);
            if (sheet != null && !workbook.IsSheetHidden(i) && !workbook.IsSheetVeryHidden(i))
                sheets.Add(sheet);
        }
        return sheets;
    }

    #endregion

    #region SetAllSheetAutoCompute(设置所有工作表自动计算)

    /// <summary>
    /// 将所有工作表设置为自动计算。
    /// </summary>
    /// <remarks>
    /// 空工作簿不执行任何操作。
    /// </remarks>
    /// <param name="workbook">工作簿</param>
    public static void SetAllSheetAutoCompute(this IWorkbook workbook)
    {
        if (workbook.NumberOfSheets <= 0)
            return;
        for (var i = 0; i < workbook.NumberOfSheets; i++)
            workbook.GetSheetAt(i).ForceFormulaRecalculation = true;// 让公式自动计算
    }

    #endregion

    #region AddSheet(添加工作表)

    /// <summary>
    /// 创建带有表头行和默认表头样式的工作表。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="name">工作表名称。</param>
    /// <param name="heads">按列顺序写入的表头文本。</param>
    /// <returns>新创建的工作表。</returns>
    public static ISheet AddSheet(this IWorkbook workbook, string name, List<string> heads)
    {
        if (heads is null)
            throw new ArgumentNullException(nameof(heads));
        var sheet = workbook.CreateSheet(name);
        var style = workbook.DefaultHeadStyle();
        var row = sheet.CreateRow(0);
        row.Height = 20 * 20;
        for (var index = 0; index < heads.Count; index++)
            row.Value(index, heads[index], style);
        return sheet;
    }

    #endregion

    #region DefaultHeadStyle(默认表头样式)

    /// <summary>
    /// 创建默认表头单元格样式。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <returns>已配置加粗、居中、黄色填充和边框的单元格样式。</returns>
    public static ICellStyle DefaultHeadStyle(this IWorkbook workbook)
    {
        var style = workbook.CreateCellStyle();
        var font = workbook.CreateFont();
        font.IsBold = true;// 加粗

        style.FillForegroundColor = 13;// 13为黄色
        style.FillPattern = FillPattern.SolidForeground;
        style.BorderTop = BorderStyle.Thin;
        style.BorderBottom = BorderStyle.Thin;
        style.BorderLeft = BorderStyle.Thin;
        style.BorderRight = BorderStyle.Thin;
        style.Alignment = HorizontalAlignment.Center;
        style.VerticalAlignment = VerticalAlignment.Center;
        style.SetFont(font.DefaultFont());
        return style;
    }

    #endregion

    #region DefaultBodyStyle(默认正文样式)

    /// <summary>
    /// 创建默认正文单元格样式。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <returns>已配置居中和边框的单元格样式。</returns>
    public static ICellStyle DefaultBodyStyle(this IWorkbook workbook)
    {
        var style = workbook.CreateCellStyle();
        var font = workbook.CreateFont();
        style.Alignment = HorizontalAlignment.Center;
        style.VerticalAlignment = VerticalAlignment.Center;
        style.BorderTop = BorderStyle.Thin;
        style.BorderBottom = BorderStyle.Thin;
        style.BorderLeft = BorderStyle.Thin;
        style.BorderRight = BorderStyle.Thin;
        style.SetFont(font.DefaultFont());
        return style;
    }

    #endregion
}

using Bing.Collections;
using Bing.Offices.Attributes;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Metadata.Excels;

namespace Bing.Offices.Filters;

/// <summary>
/// 日期过滤器
/// </summary>
[BindFilter(typeof(DateTimeAttribute))]
internal class DateTimeFilter : FilterBase, IFilter
{
    /// <summary>
    /// 过滤
    /// </summary>
    /// <param name="workbook">工作簿</param>
    /// <param name="context">过滤器上下文</param>
    /// <param name="options">导入选项配置</param>
    public override IWorkbook Filter(IWorkbook workbook, IFilterContext context, IImportOptions options)
    {
        foreach (var sheet in workbook.Sheets)
        {
            foreach (var row in sheet.GetBody())
            {
                if (!row.Valid && options.ValidateMode == ValidateMode.StopOnFirstFailure)
                    continue;
                row.Cells.ForEach(cell =>
                {
                    var attribute = cell.GetFilterAttribute<DateTimeAttribute>(context.TypeFilterInfo);
                    if (attribute != null)
                        row.Valid(cell.IsDateTime(), cell, attribute.ErrorMsg);
                });
            }
        }
        return workbook;
    }
}
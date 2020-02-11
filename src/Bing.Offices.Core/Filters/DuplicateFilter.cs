using System.Collections.Generic;
using Bing.Offices.Abstractions.Filters;
using Bing.Offices.Abstractions.Imports;
using Bing.Offices.Abstractions.Metadata.Excels;
using Bing.Offices.Attributes;
using Bing.Offices.Extensions;
using Bing.Extensions;

namespace Bing.Offices.Filters
{
    /// <summary>
    /// 重复过滤器
    /// </summary>
    [BindFilter(typeof(DuplicationAttribute))]
    internal class DuplicateFilter : FilterBase, IFilter
    {
        /// <summary>
        /// 过滤
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="context">过滤器上下文</param>
        /// <param name="options">导入选项配置</param>
        public override IWorkbook Filter(IWorkbook workbook, IFilterContext context, IImportOptions options)
        {
            var kvps = new List<KeyValuePair<int, string>>();
            foreach (var sheet in workbook.Sheets)
            {
                foreach (var row in sheet.GetBody())
                {
                    if (!row.Valid && options.ValidateMode == ValidateMode.StopOnFirstFailure)
                        continue;
                    row.Cells.ForEach(cell =>
                    {
                        var kvp = new KeyValuePair<int, string>(cell.ColumnIndex, cell.Value.ToString());
                        var attribute = cell.GetFilterAttribute<DuplicationAttribute>(context.TypeFilterInfo);
                        if (attribute != null)
                            row.Valid(!kvps.Contains(kvp), cell, attribute.ErrorMsg);
                        kvps.Add(kvp);
                    });
                }
            }
            return workbook;
        }
    }
}

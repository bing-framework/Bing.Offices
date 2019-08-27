using System.Text.RegularExpressions;
using Bing.Offices.Abstractions.Filters;
using Bing.Offices.Abstractions.Imports;
using Bing.Offices.Abstractions.Metadata.Excels;
using Bing.Offices.Attributes;
using Bing.Offices.Extensions;
using Bing.Utils.Extensions;

namespace Bing.Offices.Filters
{
    /// <summary>
    /// 正则表达式过滤器
    /// </summary>
    [BindFilter(typeof(RegexAttribute))]
    public class RegexFilter : FilterBase, IFilter
    {
        /// <summary>
        /// 正则表达式字符串
        /// </summary>
        public string RegexString { get; set; }

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
                        var attributes = cell.GetFilterAttributes<RegexAttribute>(context.TypeFilterInfo);
                        if (attributes != null && attributes.Count > 0)
                            attributes.ForEach(x => row.Valid(Regex.IsMatch(cell.Value.ToString(), x.RegexString), cell, x.ErrorMsg));
                    });
                }
            }
            return workbook;
        }
    }
}

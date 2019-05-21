using System.Collections.Generic;
using Bing.Offices.Abstractions;
using Bing.Offices.Attributes.Filters;
using Bing.Offices.Enums;
using Bing.Offices.Models.Excels;

namespace Bing.Offices.Filters.Internal
{
    /// <summary>
    /// 必填过滤器
    /// </summary>
    [FilterBind(typeof(RequiredAttribute))]
    public class RequiredFilter:IFilter
    {
        /// <summary>
        /// 过滤
        /// </summary>
        /// <param name="dataRows">数据行</param>
        /// <param name="context">过滤器上下文</param>
        /// <param name="importOptions">导入选项</param>
        public List<DataRow> Filter(List<DataRow> dataRows, FilterContext context, ImportOptions importOptions)
        {
            foreach (var dataRow in dataRows)
            {
                if (!dataRow.IsValid && importOptions.ValidateMode == ValidateMode.Stop)
                {
                    continue;
                }
                dataRow.Cells.ForEach(x =>
                {
                });
            }

            return dataRows;
        }
    }
}

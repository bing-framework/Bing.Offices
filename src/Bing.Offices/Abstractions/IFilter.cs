using System.Collections.Generic;
using Bing.Offices.Filters;
using Bing.Offices.Models.Excels;

namespace Bing.Offices.Abstractions
{
    /// <summary>
    /// 过滤器
    /// </summary>
    public interface IFilter
    {
        /// <summary>
        /// 过滤
        /// </summary>
        /// <param name="dataRows">数据行</param>
        /// <param name="context">过滤器上下文</param>
        /// <param name="importOptions">导入选项</param>
        List<DataRow> Filter(List<DataRow> dataRows, FilterContext context, ImportOptions importOptions);
    }
}

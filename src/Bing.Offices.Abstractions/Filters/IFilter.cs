using Bing.Offices.Abstractions.Imports;
using Bing.Offices.Abstractions.Metadata.Excels;

namespace Bing.Offices.Abstractions.Filters
{
    /// <summary>
    /// 过滤器
    /// </summary>
    public interface IFilter
    {
        /// <summary>
        /// 过滤
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="context">过滤器上下文</param>
        /// <param name="options">导入选项配置</param>
        IWorkbook Filter(IWorkbook workbook, IFilterContext context, IImportOptions options);
    }
}

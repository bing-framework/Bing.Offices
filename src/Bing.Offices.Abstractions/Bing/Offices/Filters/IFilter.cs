using Bing.Offices.Imports;
using Bing.Offices.Metadata.Excels;

namespace Bing.Offices.Filters
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

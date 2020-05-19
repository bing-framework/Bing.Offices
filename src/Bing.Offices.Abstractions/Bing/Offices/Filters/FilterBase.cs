using Bing.Offices.Imports;
using Bing.Offices.Metadata.Excels;

namespace Bing.Offices.Filters
{
    /// <summary>
    /// 过滤器基类
    /// </summary>
    public abstract class FilterBase : IFilter
    {
        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMsg { get; set; }

        /// <summary>
        /// 过滤
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="context">过滤器上下文</param>
        /// <param name="options">导入选项配置</param>
        public abstract IWorkbook Filter(IWorkbook workbook, IFilterContext context, IImportOptions options);
    }
}

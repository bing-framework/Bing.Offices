using Bing.Offices.Imports;
using Bing.Offices.Metadata.Excels;

namespace Bing.Offices.Filters;

/// <summary>
/// 且过滤器
/// </summary>
internal class AndFilter : IFilter
{
    /// <summary>
    /// 过滤器列表
    /// </summary>
    public IList<IFilter> Filters { get; set; }

    /// <summary>
    /// 特性类型
    /// </summary>
    public Type AttributeType => null;

    /// <summary>
    /// 过滤
    /// </summary>
    /// <param name="workbook">工作簿</param>
    /// <param name="context">过滤器上下文</param>
    /// <param name="options">导入选项配置</param>
    public IWorkbook Filter(IWorkbook workbook, IFilterContext context, IImportOptions options)
    {
        foreach (var filter in Filters)
            workbook = filter.Filter(workbook, context, options);
        return workbook;
    }
}
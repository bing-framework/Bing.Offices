using System.ComponentModel;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.Validations;

namespace Bing.Offices.Providers;

/// <summary>
/// Provider 使用的不可变映射计划。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingPlan
{
    /// <summary>
    /// 获取编译时绑定的业务 Profile 名称。
    /// </summary>
    string ProfileName { get; }
    /// <summary>
    /// 获取编译时绑定的业务模型别名。
    /// </summary>
    string ModelAlias { get; }
    /// <summary>
    /// 获取固定列映射。
    /// </summary>
    IReadOnlyList<IExcelMappingColumn> Columns { get; }
    /// <summary>
    /// 获取已预绑定的动态列映射。
    /// </summary>
    IReadOnlyList<IExcelDynamicMappingColumn> DynamicColumns { get; }
    /// <summary>
    /// 获取provider-neutral 样式配置。
    /// </summary>
    IExcelMappingStyle Style { get; }
    /// <summary>
    /// 获取provider-neutral 布局配置。
    /// </summary>
    IExcelMappingLayout Layout { get; }
}

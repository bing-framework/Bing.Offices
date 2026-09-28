using System.ComponentModel;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.Validations;

namespace Bing.Offices.Providers;

/// <summary>
/// Provider 使用的只读布局配置。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingLayout
{
    /// <summary>
    /// 获取动态列的零基物理列索引。
    /// </summary>
    int? ColumnIndex { get; }
    /// <summary>
    /// 获取动态列相对定位使用的键。
    /// </summary>
    string PlacementKey { get; }
}

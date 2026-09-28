using System.ComponentModel;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.Validations;

namespace Bing.Offices.Providers;

/// <summary>
/// Provider 使用的只读样式配置。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingStyle
{
    /// <summary>
    /// 获取表头样式的注册键。
    /// </summary>
    string HeaderStyleKey { get; }
    /// <summary>
    /// 获取数据行样式的注册键。
    /// </summary>
    string BodyStyleKey { get; }
    /// <summary>
    /// 获取列值的默认数字格式。
    /// </summary>
    string NumberFormat { get; }
}

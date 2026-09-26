using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 实体导入的资源限制选项。
/// </summary>
/// <remarks>
/// 该选项只影响实体导入，不改变实体布局的可复用性，也不影响实体导出。
/// 未提供限制时使用 <see cref="ExcelResourceLimits" /> 的默认值。
/// </remarks>
public sealed class ExcelEntityImportOptions
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityImportOptions" /> 类型的实例。
    /// </summary>
    /// <param name="resourceLimits">实体导入所使用的资源限制；为空时使用默认限制。</param>
    public ExcelEntityImportOptions(ExcelResourceLimits resourceLimits = null)
    {
        ResourceLimits = resourceLimits ?? new ExcelResourceLimits();
        ResourceLimits.Validate();
    }

    /// <summary>
    /// 获取实体导入所使用的资源限制。
    /// </summary>
    public ExcelResourceLimits ResourceLimits { get; }
}

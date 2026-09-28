using System.ComponentModel;
using System.Text.RegularExpressions;

namespace Bing.Offices.Exports;

/// <summary>
/// 自动筛选区域定义。
/// </summary>
public sealed class ExcelAutoFilterDefinition
{
    /// <summary>
    /// 获取或初始化筛选区域。
    /// </summary>
    public ExcelRangeDefinition Range { get; init; }
    /// <summary>
    /// 验证定义。
    /// </summary>
    /// <param name="parameterName">验证失败时使用的参数名称；未指定时使用对应属性名称。</param>
    public void Validate(string parameterName = null)
    {
        if (Range == null)
            throw new ArgumentException("自动筛选必须指定区域。", parameterName ?? nameof(Range));
        Range.Validate(parameterName ?? nameof(Range));
    }
}

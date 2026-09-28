using System.ComponentModel;
using System.Text.RegularExpressions;

namespace Bing.Offices.Exports;

/// <summary>
/// Excel 表格定义。
/// </summary>
public sealed class ExcelTableDefinition
{
    /// <summary>
    /// 获取或初始化表格稳定名称。
    /// </summary>
    public string Name { get; init; }
    /// <summary>
    /// 获取或初始化表格区域。
    /// </summary>
    public ExcelRangeDefinition Range { get; init; }
    /// <summary>
    /// 获取或初始化是否包含表头。
    /// </summary>
    public bool HasHeaders { get; init; } = true;
    /// <summary>
    /// 获取或初始化是否显示汇总行。
    /// </summary>
    public bool ShowTotals { get; init; }
    /// <summary>
    /// 获取或初始化可选的内置表格样式名称。
    /// </summary>
    public string StyleName { get; init; }

    /// <summary>
    /// 验证定义。
    /// </summary>
    /// <param name="parameterName">验证失败时使用的参数名称；未指定时使用对应属性名称。</param>
    public void Validate(string parameterName = null)
    {
        if (string.IsNullOrWhiteSpace(Name))
            throw new ArgumentException("表格必须指定名称。", parameterName ?? nameof(Name));
        Range?.Validate(parameterName ?? nameof(Range));
        if (Range == null)
            throw new ArgumentException("表格必须指定区域。", parameterName ?? nameof(Range));
    }
}

using System.ComponentModel;

namespace Bing.Offices.Exports;

/// <summary>
/// 图表数据系列。
/// </summary>
public sealed class ExcelChartSeries
{
    /// <summary>
    /// 获取或初始化系列显示名称。
    /// </summary>
    public string Name { get; init; }

    /// <summary>
    /// 获取或初始化系列数值范围。
    /// </summary>
    public ExcelChartRange Values { get; init; }

    /// <summary>
    /// 验证图表系列名称和数值范围。
    /// </summary>
    /// <param name="parameterName">发生验证错误时使用的参数名称。</param>
    internal void Validate(string parameterName)
    {
        if (string.IsNullOrWhiteSpace(Name))
            throw new ArgumentException("图表系列名称不能为空。", parameterName);
        if (Values == null)
            throw new ArgumentException("图表系列必须指定数值范围。", parameterName);
        Values.Validate(parameterName);
    }
}

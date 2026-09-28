using System.ComponentModel;

namespace Bing.Offices.Exports;

/// <summary>
/// 图表定位区域。
/// </summary>
public sealed class ExcelChartAnchor
{
    /// <summary>
    /// 获取或初始化起始行索引。
    /// </summary>
    public int StartRow { get; init; }

    /// <summary>
    /// 获取或初始化起始列索引。
    /// </summary>
    public int StartColumn { get; init; }

    /// <summary>
    /// 获取或初始化结束行索引（不包含）。
    /// </summary>
    public int EndRow { get; init; }

    /// <summary>
    /// 获取或初始化结束列索引（不包含）。
    /// </summary>
    public int EndColumn { get; init; }

    /// <summary>
    /// 验证图表定位区域的行列边界。
    /// </summary>
    internal void Validate()
    {
        if (StartRow < 0 || StartColumn < 0 || EndRow <= StartRow || EndColumn <= StartColumn)
            throw new ArgumentException("图表定位区域必须是非空的正向区域。", nameof(ExcelChartAnchor));
    }
}

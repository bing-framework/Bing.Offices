using System.ComponentModel;

namespace Bing.Offices.Exports;

/// <summary>
/// 提供程序无关的 Excel 图表定义。
/// </summary>
public sealed class ExcelChartDefinition
{
    /// <summary>
    /// 获取或初始化图表标题。
    /// </summary>
    public string Title { get; init; }

    /// <summary>
    /// 获取或初始化图表类型。
    /// </summary>
    public ExcelChartType Type { get; init; }

    /// <summary>
    /// 获取或初始化分类轴范围。
    /// </summary>
    public ExcelChartRange Categories { get; init; }

    /// <summary>
    /// 获取或初始化数值系列。
    /// </summary>
    public IReadOnlyList<ExcelChartSeries> Series { get; init; } = Array.Empty<ExcelChartSeries>();

    /// <summary>
    /// 获取或初始化图表定位区域。
    /// </summary>
    public ExcelChartAnchor Anchor { get; init; }

    /// <summary>
    /// 验证图表定义的范围和系列约束。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public void Validate()
    {
        if (Categories == null)
            throw new ArgumentException("图表必须指定分类范围。", nameof(Categories));
        Categories.Validate(nameof(Categories));
        if (Anchor == null)
            throw new ArgumentException("图表必须指定定位区域。", nameof(Anchor));
        Anchor.Validate();
        if (Series == null || Series.Count == 0)
            throw new ArgumentException("图表至少需要一个数据系列。", nameof(Series));
        if (Type == ExcelChartType.Pie && Series.Count != 1)
            throw new NotSupportedException("饼图只支持一个数据系列。 ");
        foreach (var series in Series)
        {
            if (series == null)
                throw new ArgumentException("图表系列不能为 null。", nameof(Series));
            series.Validate(nameof(Series));
        }
    }
}

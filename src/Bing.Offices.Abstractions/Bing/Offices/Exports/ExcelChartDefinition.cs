using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Bing.Offices.Exports;

/// <summary>
/// 图表类型。
/// </summary>
public enum ExcelChartType
{
    Column,
    Line,
    Pie
}

/// <summary>
/// 图表定位区域。
/// </summary>
public sealed class ExcelChartAnchor
{
    /// <summary>
    /// 起始行索引。
    /// </summary>
    public int StartRow { get; init; }

    /// <summary>
    /// 起始列索引。
    /// </summary>
    public int StartColumn { get; init; }

    /// <summary>
    /// 结束行索引（不包含）。
    /// </summary>
    public int EndRow { get; init; }

    /// <summary>
    /// 结束列索引（不包含）。
    /// </summary>
    public int EndColumn { get; init; }

    internal void Validate()
    {
        if (StartRow < 0 || StartColumn < 0 || EndRow <= StartRow || EndColumn <= StartColumn)
            throw new ArgumentException("图表定位区域必须是非空的正向区域。", nameof(ExcelChartAnchor));
    }
}

/// <summary>
/// 图表数据范围。
/// </summary>
public sealed class ExcelChartRange
{
    /// <summary>
    /// 数据列稳定 Key。
    /// </summary>
    public string ColumnKey { get; init; }

    /// <summary>
    /// 数据起始行索引（不含表头）。为 null 时使用当前 Sheet 数据起始行。
    /// </summary>
    public int? StartRow { get; init; }

    /// <summary>
    /// 数据结束行索引（不包含）。为 null 时使用当前 Sheet 最后一行之后。
    /// </summary>
    public int? EndRow { get; init; }

    internal void Validate(string parameterName)
    {
        if (string.IsNullOrWhiteSpace(ColumnKey))
            throw new ArgumentException("图表范围必须指定列 Key。", parameterName);
        if (StartRow.HasValue && StartRow.Value < 0)
            throw new ArgumentOutOfRangeException(parameterName);
        if (EndRow.HasValue && EndRow.Value < 0)
            throw new ArgumentOutOfRangeException(parameterName);
        if (StartRow.HasValue && EndRow.HasValue && EndRow.Value <= StartRow.Value)
            throw new ArgumentException("图表范围结束行必须大于起始行。", parameterName);
    }
}

/// <summary>
/// 图表数据系列。
/// </summary>
public sealed class ExcelChartSeries
{
    /// <summary>
    /// 系列显示名称。
    /// </summary>
    public string Name { get; init; }

    /// <summary>
    /// 系列数值范围。
    /// </summary>
    public ExcelChartRange Values { get; init; }

    internal void Validate(string parameterName)
    {
        if (string.IsNullOrWhiteSpace(Name))
            throw new ArgumentException("图表系列名称不能为空。", parameterName);
        if (Values == null)
            throw new ArgumentException("图表系列必须指定数值范围。", parameterName);
        Values.Validate(parameterName);
    }
}

/// <summary>
/// 提供程序无关的 Excel 图表定义。
/// </summary>
public sealed class ExcelChartDefinition
{
    /// <summary>
    /// 图表标题。
    /// </summary>
    public string Title { get; init; }

    /// <summary>
    /// 图表类型。
    /// </summary>
    public ExcelChartType Type { get; init; }

    /// <summary>
    /// 分类轴范围。
    /// </summary>
    public ExcelChartRange Categories { get; init; }

    /// <summary>
    /// 数值系列。
    /// </summary>
    public IReadOnlyList<ExcelChartSeries> Series { get; init; } = Array.Empty<ExcelChartSeries>();

    /// <summary>
    /// 图表定位区域。
    /// </summary>
    public ExcelChartAnchor Anchor { get; init; }

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

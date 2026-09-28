using System.ComponentModel;

namespace Bing.Offices.Exports;

/// <summary>
/// 图表数据范围。
/// </summary>
public sealed class ExcelChartRange
{
    /// <summary>
    /// 获取或初始化数据列稳定 Key。
    /// </summary>
    public string ColumnKey { get; init; }

    /// <summary>
    /// 获取或初始化数据起始行索引（不含表头）。为 null 时使用当前 Sheet 数据起始行。
    /// </summary>
    public int? StartRow { get; init; }

    /// <summary>
    /// 获取或初始化数据结束行索引（不包含）。为 null 时使用当前 Sheet 最后一行之后。
    /// </summary>
    public int? EndRow { get; init; }

    /// <summary>
    /// 验证图表数据范围的列键和行边界。
    /// </summary>
    /// <param name="parameterName">发生验证错误时使用的参数名称。</param>
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

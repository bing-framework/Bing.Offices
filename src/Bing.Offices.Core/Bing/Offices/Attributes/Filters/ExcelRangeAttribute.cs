using System;

namespace Bing.Offices.Attributes;

/// <summary>
/// v2 数值区间校验特性，区间为闭区间。
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public sealed class ExcelRangeAttribute : FilterAttributeBase
{
    /// <summary>
    /// 初始化一个<see cref="ExcelRangeAttribute"/>实例。
    /// </summary>
    /// <param name="min">最小值。</param>
    /// <param name="max">最大值。</param>
    public ExcelRangeAttribute(double min, double max)
    {
        if (min > max)
            throw new ArgumentException("校验区间的最小值不能大于最大值。", nameof(min));
        Min = min;
        Max = max;
        ErrorMsg = $"超限，仅允许为{Min}-{Max}";
    }

    /// <summary>获取最小值。</summary>
    public double Min { get; }

    /// <summary>获取最大值。</summary>
    public double Max { get; }

    /// <inheritdoc />
    public override string ErrorMsg { get; set; }
}

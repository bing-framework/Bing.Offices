using System;

namespace Bing.Offices.Attributes;

/// <summary>
/// v2 最大值校验特性。
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public sealed class ExcelMaxValueAttribute : FilterAttributeBase
{
    /// <summary>
    /// 初始化一个<see cref="ExcelMaxValueAttribute"/>实例。
    /// </summary>
    /// <param name="maxValue">允许的最大值。</param>
    public ExcelMaxValueAttribute(double maxValue)
    {
        MaxValue = maxValue;
        ErrorMsg = $"不能大于{MaxValue}";
    }

    /// <summary>获取最大值。</summary>
    public double MaxValue { get; }

    /// <inheritdoc />
    public override string ErrorMsg { get; set; }
}

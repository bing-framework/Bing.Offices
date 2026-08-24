using System;

namespace Bing.Offices.Attributes;

/// <summary>
/// v2 最大长度校验特性。
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public sealed class ExcelMaxLengthAttribute : FilterAttributeBase
{
    /// <summary>
    /// 初始化一个<see cref="ExcelMaxLengthAttribute"/>实例。
    /// </summary>
    /// <param name="maxLength">允许的最大字符数。</param>
    public ExcelMaxLengthAttribute(int maxLength)
    {
        if (maxLength < 0)
            throw new ArgumentOutOfRangeException(nameof(maxLength));
        MaxLength = maxLength;
    }

    /// <summary>获取最大字符数。</summary>
    public int MaxLength { get; }

    /// <inheritdoc />
    public override string ErrorMsg { get; set; } = "超长";
}

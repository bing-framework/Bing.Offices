using System;

namespace Bing.Offices.Attributes;

/// <summary>
/// v2 唯一值校验特性。
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public sealed class ExcelUniqueAttribute : FilterAttributeBase
{
    /// <summary>
    /// 获取或设置是否忽略空值；默认忽略空值。
    /// </summary>
    public bool IgnoreEmpty { get; set; } = true;

    /// <inheritdoc />
    public override string ErrorMsg { get; set; } = "重复数据";
}

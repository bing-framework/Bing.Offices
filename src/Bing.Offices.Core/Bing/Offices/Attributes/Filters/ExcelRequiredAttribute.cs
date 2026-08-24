using System;

namespace Bing.Offices.Attributes;

/// <summary>
/// v2 必填校验特性。
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public sealed class ExcelRequiredAttribute : FilterAttributeBase
{
    /// <inheritdoc />
    public override string ErrorMsg { get; set; } = "必填";
}

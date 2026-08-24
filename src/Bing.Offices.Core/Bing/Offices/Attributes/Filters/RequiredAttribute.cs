

// ReSharper disable once CheckNamespace
namespace Bing.Offices.Attributes;

/// <summary>
/// 必填特性
/// </summary>
[System.Obsolete("请改用 ExcelRequiredAttribute。")]
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public class RequiredAttribute : FilterAttributeBase
{
    /// <summary>
    /// 错误消息
    /// </summary>
    public override string ErrorMsg { get; set; } = "必填";
}

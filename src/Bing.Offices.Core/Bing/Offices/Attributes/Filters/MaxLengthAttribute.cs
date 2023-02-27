using System;

// ReSharper disable once CheckNamespace
namespace Bing.Offices.Attributes;

/// <summary>
/// 最大长度特性
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public class MaxLengthAttribute : FilterAttributeBase
{
    /// <summary>
    /// 最大长度
    /// </summary>
    public int MaxLength { get; set; }

    /// <summary>
    /// 错误消息
    /// </summary>
    public override string ErrorMsg { get; set; } = "超长";

    /// <summary>
    /// 初始化一个<see cref="MaxLengthAttribute"/>类型的实例
    /// </summary>
    /// <param name="maxLength">最大长度</param>
    public MaxLengthAttribute(int maxLength) => MaxLength = maxLength;
}
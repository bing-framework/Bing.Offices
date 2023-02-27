using System;

// ReSharper disable once CheckNamespace
namespace Bing.Offices.Attributes;

/// <summary>
/// 重复数据校验特性
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public class DuplicationAttribute : FilterAttributeBase
{
    /// <summary>
    /// 错误消息
    /// </summary>
    public override string ErrorMsg { get; set; } = "重复数据";
}
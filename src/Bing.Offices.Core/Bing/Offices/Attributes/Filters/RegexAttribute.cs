using System;

// ReSharper disable once CheckNamespace
namespace Bing.Offices.Attributes;

/// <summary>
/// 正则表达式特性
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = false)]
public class RegexAttribute : FilterAttributeBase
{
    /// <summary>
    /// 正则表达式字符串
    /// </summary>
    public string RegexString { get; set; }

    /// <summary>
    /// 初始化一个<see cref="RegexAttribute"/>类型的实例
    /// </summary>
    /// <param name="regex">正则表达式</param>
    public RegexAttribute(string regex)
    {
        RegexString = regex;
    }
}
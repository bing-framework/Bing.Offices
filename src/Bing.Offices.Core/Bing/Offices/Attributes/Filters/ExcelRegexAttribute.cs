using System;

namespace Bing.Offices.Attributes;

/// <summary>
/// v2 正则表达式校验特性。
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = false)]
public sealed class ExcelRegexAttribute : FilterAttributeBase
{
    /// <summary>
    /// 初始化一个<see cref="ExcelRegexAttribute"/>实例。
    /// </summary>
    /// <param name="pattern">正则表达式。</param>
    public ExcelRegexAttribute(string pattern)
    {
        Pattern = pattern ?? throw new ArgumentNullException(nameof(pattern));
    }

    /// <summary>
    /// 获取正则表达式。
    /// </summary>
    public string Pattern { get; }

    /// <inheritdoc />
    public override string ErrorMsg { get; set; } = "格式不正确";
}

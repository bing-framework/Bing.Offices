using System.ComponentModel;
using System.Text.RegularExpressions;

namespace Bing.Offices.Exports;

/// <summary>
/// 名称范围定义。
/// </summary>
public sealed class ExcelNamedRangeDefinition
{
    /// <summary>
    /// 获取或初始化名称范围名称。
    /// </summary>
    public string Name { get; init; }
    /// <summary>
    /// 获取或初始化可选的 Sheet 作用域；为空表示 Workbook 作用域。
    /// </summary>
    public string SheetName { get; init; }
    /// <summary>
    /// 获取或初始化有界区域的 A1 引用。
    /// </summary>
    public string Address { get; init; }

    /// <summary>
    /// 验证定义。
    /// </summary>
    /// <param name="parameterName">验证失败时使用的参数名称；未指定时使用对应属性名称。</param>
    public void Validate(string parameterName = null)
    {
        if (string.IsNullOrWhiteSpace(Name))
            throw new ArgumentException("名称范围必须指定名称。", parameterName ?? nameof(Name));
        if (string.IsNullOrWhiteSpace(Address)
            || Address.IndexOf('[') >= 0
            || Address.IndexOf(']') >= 0
            || !Regex.IsMatch(Address, @"^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$",
                RegexOptions.CultureInvariant))
            throw new ArgumentException("名称范围必须指定有界区域引用。", parameterName ?? nameof(Address));
    }
}

using System.ComponentModel;
using System.Text.RegularExpressions;

namespace Bing.Offices.Exports;

/// <summary>
/// 定义工作簿中的文本字典列表。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelLookupListDefinition
{
    /// <summary>
    /// 匹配 Excel 名称标识符。
    /// </summary>
    private static readonly Regex NamePattern = new(@"^[A-Za-z_\\][A-Za-z0-9_.\\]*$",
        RegexOptions.CultureInvariant);

    /// <summary>
    /// 排除与单元格引用相同的名称。
    /// </summary>
    private static readonly Regex CellReferencePattern = new(@"^[A-Za-z]{1,3}[1-9][0-9]*$",
        RegexOptions.CultureInvariant);

    /// <summary>
    /// 初始化一个 <see cref="ExcelLookupListDefinition" /> 类型的实例。
    /// </summary>
    /// <param name="name">工作簿级名称。</param>
    /// <param name="values">按显示顺序排列的文本值。</param>
    internal ExcelLookupListDefinition(string name, IReadOnlyList<string> values)
    {
        ValidateName(name);
        if (values == null)
            throw new ArgumentNullException(nameof(values));
        var normalized = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var value in values)
        {
            var text = value?.Trim();
            if (string.IsNullOrEmpty(text))
                throw new ArgumentException("字典不能包含空值。", nameof(values));
            if (text.Length > 32767)
                throw new ArgumentException("字典项超过 Excel 单元格长度上限。", nameof(values));
            if (seen.Add(text))
                normalized.Add(text);
        }
        if (normalized.Count == 0)
            throw new ArgumentException("字典不能为空。", nameof(values));
        Name = name;
        Values = Array.AsReadOnly(normalized.ToArray());
    }

    /// <summary>
    /// 获取工作簿级名称。
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// 获取字典文本快照。
    /// </summary>
    public IReadOnlyList<string> Values { get; }

    /// <summary>
    /// 验证工作簿级名称。
    /// </summary>
    /// <param name="name">待验证的名称。</param>
    internal static void ValidateName(string name)
    {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 255 || !NamePattern.IsMatch(name)
            || CellReferencePattern.IsMatch(name))
            throw new ArgumentException("字典名称不是有效的 Excel 名称。", nameof(name));
    }
}

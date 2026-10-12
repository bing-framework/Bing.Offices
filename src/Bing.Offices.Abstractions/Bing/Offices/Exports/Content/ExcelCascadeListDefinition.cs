using System.ComponentModel;

namespace Bing.Offices.Exports;

/// <summary>
/// 定义工作簿中的两级联动文本字典。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelCascadeListDefinition
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelCascadeListDefinition" /> 类型的实例。
    /// </summary>
    /// <param name="name">上级列表的工作簿名称。</param>
    /// <param name="values">上级文本与下级文本列表。</param>
    internal ExcelCascadeListDefinition(string name,
        IReadOnlyDictionary<string, IReadOnlyList<string>> values)
    {
        ExcelLookupListDefinition.ValidateName(name);
        if (values == null)
            throw new ArgumentNullException(nameof(values));
        var parents = new List<ExcelCascadeListEntry>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in values)
        {
            var parent = pair.Key?.Trim();
            if (string.IsNullOrEmpty(parent) || parent.Length > 32767 || !seen.Add(parent))
                throw new ArgumentException("级联上级文本为空、过长或重复。", nameof(values));
            var children = new ExcelLookupListDefinition(name, pair.Value).Values;
            parents.Add(new ExcelCascadeListEntry(parent, children));
        }
        if (parents.Count == 0)
            throw new ArgumentException("级联字典不能为空。", nameof(values));
        Name = name;
        Entries = Array.AsReadOnly(parents.OrderBy(item => item.Parent, StringComparer.Ordinal).ToArray());
    }

    /// <summary>
    /// 获取工作簿级名称。
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// 获取上级文本及其下级列表快照。
    /// </summary>
    public IReadOnlyList<ExcelCascadeListEntry> Entries { get; }
}

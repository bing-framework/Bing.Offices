using System.ComponentModel;

namespace Bing.Offices.Exports;

/// <summary>
/// 定义一个上级文本及其下级文本列表。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExcelCascadeListEntry
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelCascadeListEntry" /> 类型的实例。
    /// </summary>
    /// <param name="parent">上级文本。</param>
    /// <param name="values">下级文本快照。</param>
    internal ExcelCascadeListEntry(string parent, IReadOnlyList<string> values)
    {
        Parent = parent;
        Values = Array.AsReadOnly(values.ToArray());
    }

    /// <summary>
    /// 获取上级文本。
    /// </summary>
    public string Parent { get; }

    /// <summary>
    /// 获取下级文本快照。
    /// </summary>
    public IReadOnlyList<string> Values { get; }
}

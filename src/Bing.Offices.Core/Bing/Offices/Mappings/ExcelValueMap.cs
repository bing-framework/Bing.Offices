using System.Collections.ObjectModel;

namespace Bing.Offices.Mappings;

/// <summary>
/// Excel 单元格显示文本与业务值之间的不可变映射。
/// </summary>
/// <typeparam name="TValue">业务值类型。</typeparam>
public sealed class ExcelValueMap<TValue>
{
    /// <summary>
    /// 初始化一个<see cref="ExcelValueMap{TValue}"/>类型的实例。
    /// </summary>
    /// <param name="values">显示文本到业务值的映射。</param>
    public ExcelValueMap(IDictionary<string, TValue> values)
    {
        if (values == null)
            throw new ArgumentNullException(nameof(values));
        Values = new ReadOnlyDictionary<string, TValue>(
            new Dictionary<string, TValue>(values, StringComparer.Ordinal));
    }

    /// <summary>
    /// 获取显示文本到业务值的映射。
    /// </summary>
    public IReadOnlyDictionary<string, TValue> Values { get; }
}

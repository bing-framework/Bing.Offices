namespace Bing.Offices.Imports;

/// <summary>
/// provider-neutral 的工作表选择器。
/// </summary>
public sealed class ExcelSheetSelector
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelSheetSelector" /> 类型的实例。
    /// </summary>
    /// <param name="kind">选择方式。</param>
    /// <param name="name">按名称选择时使用的工作表名称。</param>
    /// <param name="index">按索引选择时使用的零基索引。</param>
    private ExcelSheetSelector(ExcelSheetSelectorKind kind, string name, int? index)
    {
        Kind = kind;
        Name = name;
        Index = index;
    }

    /// <summary>
    /// 按名称创建选择器。
    /// </summary>
    /// <param name="name">工作表名称。</param>
    /// <returns>按名称选择工作表的选择器。</returns>
    public static ExcelSheetSelector ByName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Sheet 名称不能为空。", nameof(name));
        return new ExcelSheetSelector(ExcelSheetSelectorKind.ByName, name, null);
    }

    /// <summary>
    /// 按从零开始的索引创建选择器。
    /// </summary>
    /// <param name="index">工作表的零基索引。</param>
    /// <returns>按索引选择工作表的选择器。</returns>
    public static ExcelSheetSelector ByIndex(int index)
    {
        if (index < 0)
            throw new ArgumentOutOfRangeException(nameof(index));
        return new ExcelSheetSelector(ExcelSheetSelectorKind.ByIndex, null, index);
    }

    /// <summary>
    /// 获取选择方式。
    /// </summary>
    public ExcelSheetSelectorKind Kind { get; }

    /// <summary>
    /// 获取名称选择值。
    /// </summary>
    /// <remarks>
    /// 仅 <see cref="ExcelSheetSelectorKind.ByName"/> 模式有值，按索引选择时为 <see langword="null" />。
    /// </remarks>
    public string Name { get; }

    /// <summary>
    /// 获取索引选择值。
    /// </summary>
    /// <remarks>
    /// 仅 <see cref="ExcelSheetSelectorKind.ByIndex"/> 模式有值，按名称选择时为 <see langword="null" />。
    /// </remarks>
    public int? Index { get; }
}

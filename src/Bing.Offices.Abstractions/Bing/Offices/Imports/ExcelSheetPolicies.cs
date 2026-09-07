using System;
using System.IO;

namespace Bing.Offices.Imports;

/// <summary>
/// Sheet 名称匹配策略。
/// </summary>
public enum ExcelNameComparison
{
    /// <summary>区分大小写。</summary>
    Ordinal,
    /// <summary>忽略大小写。</summary>
    OrdinalIgnoreCase
}

/// <summary>
/// 单元格文本空白规范化策略。
/// </summary>
public enum ExcelWhitespacePolicy
{
    /// <summary>保留原始文本。</summary>
    Preserve,
    /// <summary>移除首尾空白。</summary>
    Trim,
    /// <summary>移除全部 Unicode 空白字符。</summary>
    RemoveAll
}

/// <summary>
/// 工作表选择方式。
/// </summary>
public enum ExcelSheetSelectorKind
{
    /// <summary>按工作表名称选择。</summary>
    ByName,
    /// <summary>按从零开始的工作表索引选择。</summary>
    ByIndex
}

/// <summary>
/// provider-neutral 的工作表选择器。
/// </summary>
public sealed class ExcelSheetSelector
{
    private ExcelSheetSelector(ExcelSheetSelectorKind kind, string name, int? index)
    {
        Kind = kind;
        Name = name;
        Index = index;
    }

    /// <summary>按名称创建选择器。</summary>
    public static ExcelSheetSelector ByName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Sheet 名称不能为空。", nameof(name));
        return new ExcelSheetSelector(ExcelSheetSelectorKind.ByName, name, null);
    }

    /// <summary>按从零开始的索引创建选择器。</summary>
    public static ExcelSheetSelector ByIndex(int index)
    {
        if (index < 0)
            throw new ArgumentOutOfRangeException(nameof(index));
        return new ExcelSheetSelector(ExcelSheetSelectorKind.ByIndex, null, index);
    }

    /// <summary>获取选择方式。</summary>
    public ExcelSheetSelectorKind Kind { get; }

    /// <summary>获取名称选择值。</summary>
    public string Name { get; }

    /// <summary>获取索引选择值。</summary>
    public int? Index { get; }
}

/// <summary>
/// 工作表读取列范围，列索引从零开始。
/// </summary>
public sealed class ExcelReadColumnRange
{
    private ExcelReadColumnRange(int startIndex, int count)
    {
        StartIndex = startIndex;
        Count = count;
    }

    /// <summary>创建列读取范围。</summary>
    public static ExcelReadColumnRange Create(int startIndex, int count)
    {
        if (startIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(startIndex));
        if (count <= 0)
            throw new ArgumentOutOfRangeException(nameof(count));
        if ((long)startIndex + count > int.MaxValue)
            throw new ArgumentOutOfRangeException(nameof(count));
        return new ExcelReadColumnRange(startIndex, count);
    }

    /// <summary>获取起始列索引。</summary>
    public int StartIndex { get; }

    /// <summary>获取读取列数。</summary>
    public int Count { get; }

    /// <summary>判断指定列是否在范围内。</summary>
    public bool Contains(int columnIndex) => columnIndex >= StartIndex && columnIndex < StartIndex + Count;
}


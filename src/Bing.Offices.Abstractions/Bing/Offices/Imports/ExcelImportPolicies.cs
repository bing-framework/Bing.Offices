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

/// <summary>
/// 导入失败工作簿输出模式。
/// </summary>
public enum ExcelImportFailureWorkbookMode
{
    /// <summary>不生成失败工作簿。</summary>
    None,
    /// <summary>在原工作簿副本上标记错误。</summary>
    AnnotatedOriginal,
    /// <summary>只输出包含失败行的工作簿。</summary>
    ErrorRowsOnly
}

/// <summary>
/// 导入规则来源组合策略。
/// </summary>
public enum ExcelImportValidationMode
{
    /// <summary>禁用工作簿规则。</summary>
    Disabled,
    /// <summary>只执行配置和属性规则。</summary>
    ConfiguredRules,
    /// <summary>只执行工作簿原生规则。</summary>
    WorkbookRules,
    /// <summary>同时执行配置和工作簿规则。</summary>
    ConfiguredAndWorkbook
}

/// <summary>
/// 图片列多图片处理策略。
/// </summary>
public enum ExcelImageMultiplicityPolicy
{
    /// <summary>只绑定第一张图片。</summary>
    First,
    /// <summary>绑定全部图片。</summary>
    All,
    /// <summary>出现多张图片时报告错误。</summary>
    Fail
}

/// <summary>
/// 不支持的工作簿特性处理策略。
/// </summary>
public enum ExcelUnsupportedFeaturePolicy
{
    /// <summary>报告为导入错误。</summary>
    Report,
    /// <summary>直接拒绝导入。</summary>
    Fail
}

/// <summary>
/// 导入资源限制。
/// </summary>
public sealed class ExcelResourceLimits
{
    /// <summary>输入流最大字节数；null 表示不额外限制。</summary>
    public long? MaxInputBytes { get; init; }

    /// <summary>最大数据行数；null 表示不额外限制。</summary>
    public int? MaxRows { get; init; }

    /// <summary>最大错误数；null 表示不额外限制。</summary>
    public int? MaxErrors { get; init; }

    /// <summary>最大图片数量；null 表示不额外限制。</summary>
    public int? MaxPictures { get; init; }

    /// <summary>单张图片最大字节数；null 表示不额外限制。</summary>
    public long? MaxPictureBytes { get; init; }

    /// <summary>所有图片最大总字节数；null 表示不额外限制。</summary>
    public long? MaxTotalPictureBytes { get; init; }

    /// <summary>验证限制值。</summary>
    public void Validate()
    {
        if (MaxInputBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxInputBytes));
        if (MaxRows <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxRows));
        if (MaxErrors <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxErrors));
        if (MaxPictures <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxPictures));
        if (MaxPictureBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxPictureBytes));
        if (MaxTotalPictureBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxTotalPictureBytes));
    }
}

/// <summary>
/// 导入失败工作簿输出配置。
/// </summary>
public sealed class ExcelImportFailureOptions
{
    /// <summary>失败工作簿模式。</summary>
    public ExcelImportFailureWorkbookMode Mode { get; init; }

    /// <summary>失败工作簿目标流，由调用方拥有。</summary>
    public Stream Destination { get; init; }

    /// <summary>失败工作簿最大字节数。</summary>
    public long? MaxBytes { get; init; }

    /// <summary>验证失败输出配置。</summary>
    public void Validate()
    {
        if (MaxBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxBytes));
        if (Mode != ExcelImportFailureWorkbookMode.None && Destination == null)
            throw new ArgumentException("启用失败工作簿输出时必须提供目标流。", nameof(Destination));
        if (Destination != null && !Destination.CanWrite)
            throw new ArgumentException("失败工作簿目标流不可写入。", nameof(Destination));
    }
}

using System;
using System.IO;

namespace Bing.Offices.Imports;

/// <summary>
/// 导入资源限制。
/// </summary>
public sealed class ExcelResourceLimits
{
    /// <summary>输入流最大字节数，默认 128 MiB；显式设置 null 可关闭该限制。</summary>
    public long? MaxInputBytes { get; init; } = 128L * 1024 * 1024;

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

    /// <summary>单个工作表最大跟踪唯一值数量；null 表示不额外限制。</summary>
    public int? MaxTrackedUniqueValues { get; init; }

    /// <summary>唯一值比较策略。</summary>
    public StringComparison UniqueComparison { get; init; } = StringComparison.OrdinalIgnoreCase;

    /// <summary>XLSX ZIP 最大 entry 数量；为空表示不额外限制。</summary>
    public int? MaxZipEntries { get; init; } = 10000;

    /// <summary>XLSX ZIP 单个 entry 最大解压字节数；为空表示不额外限制。</summary>
    public long? MaxZipEntryUncompressedBytes { get; init; } = 64L * 1024 * 1024;

    /// <summary>XLSX ZIP 所有 entry 最大总解压字节数；为空表示不额外限制。</summary>
    public long? MaxZipTotalUncompressedBytes { get; init; } = 512L * 1024 * 1024;

    /// <summary>XLSX XML 节点和值的最大字符数量；为空表示不额外限制。</summary>
    public long? MaxXmlCharacters { get; init; } = 64L * 1024 * 1024;

    /// <summary>XLSX XML 最大嵌套深度；为空表示不额外限制。</summary>
    public int? MaxXmlDepth { get; init; } = 256;

    /// <summary>XLSX ZIP 允许的最大解压/压缩比；为空表示不额外限制。</summary>
    public double? MaxZipCompressionRatio { get; init; } = 1000;

    /// <summary>sharedStrings.xml 最大解压字节数；为空表示不额外限制。</summary>
    public long? MaxSharedStringsBytes { get; init; } = 64L * 1024 * 1024;

    /// <summary>styles.xml 最大解压字节数；为空表示不额外限制。</summary>
    public long? MaxStylesBytes { get; init; } = 16L * 1024 * 1024;

    /// <summary>单个 worksheet XML 最大解压字节数；为空表示不额外限制。</summary>
    public long? MaxWorksheetBytes { get; init; } = 64L * 1024 * 1024;

    /// <summary>所有 worksheet XML 最大解压字节数；为空表示不额外限制。</summary>
    public long? MaxTotalWorksheetBytes { get; init; } = 256L * 1024 * 1024;

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
        if (MaxTrackedUniqueValues <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxTrackedUniqueValues));
        if (MaxZipEntries <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxZipEntries));
        if (MaxZipEntryUncompressedBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxZipEntryUncompressedBytes));
        if (MaxZipTotalUncompressedBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxZipTotalUncompressedBytes));
        if (MaxXmlCharacters <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxXmlCharacters));
        if (MaxXmlDepth <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxXmlDepth));
        if (MaxZipCompressionRatio.HasValue && (MaxZipCompressionRatio.Value <= 0
            || double.IsNaN(MaxZipCompressionRatio.Value)
            || double.IsInfinity(MaxZipCompressionRatio.Value)))
            throw new ArgumentOutOfRangeException(nameof(MaxZipCompressionRatio));
        if (MaxSharedStringsBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxSharedStringsBytes));
        if (MaxStylesBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxStylesBytes));
        if (MaxWorksheetBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxWorksheetBytes));
        if (MaxTotalWorksheetBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxTotalWorksheetBytes));
        if (UniqueComparison != StringComparison.Ordinal && UniqueComparison != StringComparison.OrdinalIgnoreCase
            && UniqueComparison != StringComparison.InvariantCulture
            && UniqueComparison != StringComparison.InvariantCultureIgnoreCase
            && UniqueComparison != StringComparison.CurrentCulture
            && UniqueComparison != StringComparison.CurrentCultureIgnoreCase)
            throw new ArgumentOutOfRangeException(nameof(UniqueComparison));
    }
}


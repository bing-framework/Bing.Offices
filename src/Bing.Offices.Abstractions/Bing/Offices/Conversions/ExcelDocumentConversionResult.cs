namespace Bing.Offices.Conversions;

/// <summary>
/// 格式转换报告。
/// </summary>
public sealed class ExcelDocumentConversionResult
{
    /// <summary>
    /// 获取或初始化源格式。
    /// </summary>
    public ExcelFormat SourceFormat { get; init; }
    /// <summary>
    /// 获取或初始化目标格式。
    /// </summary>
    public ExcelFormat TargetFormat { get; init; }
    /// <summary>
    /// 获取或初始化由于格式差异产生的警告。
    /// </summary>
    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();
}

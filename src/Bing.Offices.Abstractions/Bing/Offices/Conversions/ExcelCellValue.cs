namespace Bing.Offices.Conversions;

/// <summary>
/// 不依赖具体 Excel 提供程序的单元格值描述。
/// </summary>
public sealed class ExcelCellValue
{
    /// <summary>
    /// 初始化一个<see cref="ExcelCellValue"/>类型的实例。
    /// </summary>
    /// <param name="value">供默认转换使用的原始值。</param>
    /// <param name="text">单元格显示文本。</param>
    /// <param name="kind">逻辑单元格类型。</param>
    /// <param name="cachedKind">公式缓存值的逻辑类型。</param>
    /// <param name="formula">公式文本。</param>
    /// <param name="errorCode">错误码。</param>
    /// <param name="formatIndex">提供程序格式索引。</param>
    public ExcelCellValue(object value, string text, ExcelCellKind kind, ExcelCellKind? cachedKind = null,
        string formula = null, int? errorCode = null, int? formatIndex = null)
    {
        Value = value;
        Text = text ?? string.Empty;
        Kind = kind;
        CachedKind = cachedKind;
        Formula = formula;
        ErrorCode = errorCode;
        FormatIndex = formatIndex;
    }

    /// <summary>
    /// 获取供默认转换使用的原始值。
    /// </summary>
    public object Value { get; }

    /// <summary>
    /// 获取单元格显示文本。
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// 获取逻辑单元格类型。
    /// </summary>
    public ExcelCellKind Kind { get; }

    /// <summary>
    /// 获取公式缓存值的逻辑类型。
    /// </summary>
    public ExcelCellKind? CachedKind { get; }

    /// <summary>
    /// 获取公式文本。
    /// </summary>
    public string Formula { get; }

    /// <summary>
    /// 获取错误码。
    /// </summary>
    public int? ErrorCode { get; }

    /// <summary>
    /// 获取提供程序格式索引。
    /// </summary>
    public int? FormatIndex { get; }
}

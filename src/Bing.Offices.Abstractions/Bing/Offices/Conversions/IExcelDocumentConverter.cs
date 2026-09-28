namespace Bing.Offices.Conversions;

/// <summary>
/// 独立的工作簿格式转换能力。
/// </summary>
public interface IExcelDocumentConverter
{
    /// <summary>
    /// 同步转换到调用方拥有的目标流。
    /// </summary>
    /// <param name="source">调用方拥有的输入工作簿流。</param>
    /// <param name="destination">调用方拥有的转换结果输出流。</param>
    /// <param name="sourceFormat">输入工作簿格式。</param>
    /// <param name="targetFormat">目标工作簿格式。</param>
    /// <param name="openOptions">打开工作簿的安全选项；为空时使用默认选项。</param>
    /// <param name="saveOptions">保存工作簿的安全选项；为空时使用默认选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含源格式、目标格式及格式损失警告的转换报告。</returns>
    ExcelDocumentConversionResult Convert(Stream source, Stream destination,
        ExcelFormat sourceFormat, ExcelFormat targetFormat,
        ExcelWorkbookOpenOptions openOptions = null,
        ExcelWorkbookSaveOptions saveOptions = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// 异步转换到调用方拥有的目标流。
    /// </summary>
    /// <param name="source">调用方拥有的输入工作簿流。</param>
    /// <param name="destination">调用方拥有的转换结果输出流。</param>
    /// <param name="sourceFormat">输入工作簿格式。</param>
    /// <param name="targetFormat">目标工作簿格式。</param>
    /// <param name="openOptions">打开工作簿的安全选项；为空时使用默认选项。</param>
    /// <param name="saveOptions">保存工作簿的安全选项；为空时使用默认选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>异步操作，结果为包含源格式、目标格式及格式损失警告的转换报告。</returns>
    Task<ExcelDocumentConversionResult> ConvertAsync(Stream source, Stream destination,
        ExcelFormat sourceFormat, ExcelFormat targetFormat,
        ExcelWorkbookOpenOptions openOptions = null,
        ExcelWorkbookSaveOptions saveOptions = null,
        CancellationToken cancellationToken = default);
}

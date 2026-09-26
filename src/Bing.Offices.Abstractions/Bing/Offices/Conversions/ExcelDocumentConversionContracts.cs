namespace Bing.Offices.Conversions;

/// <summary>
/// 宏项目处理策略。
/// </summary>
public enum ExcelMacroPolicy
{
    /// <summary>
    /// 遇到宏项目时拒绝操作。
    /// </summary>
    Reject,
    /// <summary>
    /// 保留宏二进制项目，但不解析、不执行或修改 VBA。
    /// </summary>
    Preserve,
    /// <summary>
    /// 显式删除宏项目。
    /// </summary>
    Strip
}

/// <summary>
/// 打开工作簿的安全选项。
/// </summary>
public sealed class ExcelWorkbookOpenOptions
{
    /// <summary>
    /// 获取或初始化解密密码。
    /// </summary>
    /// <remarks>不得将密码写入日志。</remarks>
    public string Password { get; init; }
    /// <summary>
    /// 获取或初始化宏处理策略，默认拒绝。
    /// </summary>
    public ExcelMacroPolicy MacroPolicy { get; init; } = ExcelMacroPolicy.Reject;
}

/// <summary>
/// 保存工作簿的安全选项。
/// </summary>
public sealed class ExcelWorkbookSaveOptions
{
    /// <summary>
    /// 获取或初始化加密密码。
    /// </summary>
    /// <remarks>不得将密码写入日志。</remarks>
    public string Password { get; init; }
    /// <summary>
    /// 获取或初始化宏处理策略，默认拒绝。
    /// </summary>
    public ExcelMacroPolicy MacroPolicy { get; init; } = ExcelMacroPolicy.Reject;
}

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

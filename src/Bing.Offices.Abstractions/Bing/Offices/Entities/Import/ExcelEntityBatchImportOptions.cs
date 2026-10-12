using System;

namespace Bing.Offices.Entities;

/// <summary>
/// 实体布局分批导入选项。
/// </summary>
public sealed class ExcelEntityBatchImportOptions
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityBatchImportOptions" /> 类型的实例。
    /// </summary>
    /// <param name="batchSize">每个列表区域的目标批次大小。</param>
    /// <param name="importOptions">完整实体导入选项；为空时使用默认选项。</param>
    public ExcelEntityBatchImportOptions(int batchSize = 1000,
        ExcelEntityImportOptions importOptions = null)
        : this(batchSize, importOptions, ExcelEntityBatchParsingMode.Buffered,
            ExcelEntityBatchResultMode.Complete)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityBatchImportOptions" /> 类型的实例。
    /// </summary>
    /// <param name="batchSize">每个列表区域的目标批次大小。</param>
    /// <param name="importOptions">完整实体导入选项；为空时使用默认选项。</param>
    /// <param name="parsingMode">实体列表的解析模式。</param>
    public ExcelEntityBatchImportOptions(int batchSize, ExcelEntityImportOptions importOptions,
        ExcelEntityBatchParsingMode parsingMode)
        : this(batchSize, importOptions, parsingMode, ExcelEntityBatchResultMode.Complete)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityBatchImportOptions" /> 类型的实例。
    /// </summary>
    /// <param name="batchSize">每个列表区域的目标批次大小。</param>
    /// <param name="importOptions">完整实体导入选项；为空时使用默认选项。</param>
    /// <param name="parsingMode">实体列表的解析模式。</param>
    /// <param name="resultMode">批次完成后保留的结果范围。</param>
    public ExcelEntityBatchImportOptions(int batchSize, ExcelEntityImportOptions importOptions,
        ExcelEntityBatchParsingMode parsingMode, ExcelEntityBatchResultMode resultMode)
    {
        if (batchSize <= 0)
            throw new ArgumentOutOfRangeException(nameof(batchSize));
        if (!Enum.IsDefined(typeof(ExcelEntityBatchParsingMode), parsingMode))
            throw new ArgumentOutOfRangeException(nameof(parsingMode));
        if (!Enum.IsDefined(typeof(ExcelEntityBatchResultMode), resultMode))
            throw new ArgumentOutOfRangeException(nameof(resultMode));
        BatchSize = batchSize;
        ImportOptions = importOptions ?? new ExcelEntityImportOptions();
        ParsingMode = parsingMode;
        ResultMode = resultMode;
    }

    /// <summary>
    /// 获取每个列表区域的目标批次大小。
    /// </summary>
    public int BatchSize { get; }

    /// <summary>
    /// 获取完整实体导入选项。
    /// </summary>
    public ExcelEntityImportOptions ImportOptions { get; }

    /// <summary>
    /// 获取实体列表的解析模式。
    /// </summary>
    public ExcelEntityBatchParsingMode ParsingMode { get; }

    /// <summary>
    /// 获取批次完成后保留的结果范围。
    /// </summary>
    /// <remarks>
    /// <see cref="ExcelEntityBatchResultMode.SummaryOnly" /> 使用受限的前向解析路径；
    /// 即使请求的 <see cref="ParsingMode" /> 为 Buffered，也必须满足 Provider 的 Forward 布局边界。
    /// </remarks>
    public ExcelEntityBatchResultMode ResultMode { get; }
}

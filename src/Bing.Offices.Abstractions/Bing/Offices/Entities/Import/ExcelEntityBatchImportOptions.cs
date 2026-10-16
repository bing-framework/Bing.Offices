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
        : this(batchSize, importOptions, parsingMode, resultMode, null)
    {
    }

    /// <summary>
    /// 初始化实体布局分批导入选项，并限制前向读取的临时存储。
    /// </summary>
    /// <param name="batchSize">每个列表区域的目标批次大小。</param>
    /// <param name="importOptions">完整实体导入选项；为空时使用默认选项。</param>
    /// <param name="parsingMode">实体列表的解析模式。</param>
    /// <param name="resultMode">批次完成后保留的结果范围。</param>
    /// <param name="maxTemporaryStorageBytes">同时存活的受管临时文件总字节数；为空时不限制。</param>
    public ExcelEntityBatchImportOptions(int batchSize, ExcelEntityImportOptions importOptions,
        ExcelEntityBatchParsingMode parsingMode, ExcelEntityBatchResultMode resultMode,
        long? maxTemporaryStorageBytes)
    {
        if (batchSize <= 0)
            throw new ArgumentOutOfRangeException(nameof(batchSize));
        if (!Enum.IsDefined(typeof(ExcelEntityBatchParsingMode), parsingMode))
            throw new ArgumentOutOfRangeException(nameof(parsingMode));
        if (!Enum.IsDefined(typeof(ExcelEntityBatchResultMode), resultMode))
            throw new ArgumentOutOfRangeException(nameof(resultMode));
        if (maxTemporaryStorageBytes.HasValue && maxTemporaryStorageBytes.Value <= 0)
            throw new ArgumentOutOfRangeException(nameof(maxTemporaryStorageBytes));
        BatchSize = batchSize;
        ImportOptions = importOptions ?? new ExcelEntityImportOptions();
        ParsingMode = parsingMode;
        ResultMode = resultMode;
        MaxTemporaryStorageBytes = maxTemporaryStorageBytes;
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

    /// <summary>
    /// 获取单次前向导入同时存活的受管临时文件逻辑长度总上限。
    /// </summary>
    /// <remarks>
    /// 覆盖输入暂存、共享字符串正文与索引及 XLS 转换中间文件，不含调用方流、
    /// 文件系统块开销或托管内存。仅用于 Forward 或 SummaryOnly；
    /// 启用时暂不支持模板预检及失败工作簿输出，Provider 必须在读取前明确拒绝这些组合。
    /// </remarks>
    public long? MaxTemporaryStorageBytes { get; }
}

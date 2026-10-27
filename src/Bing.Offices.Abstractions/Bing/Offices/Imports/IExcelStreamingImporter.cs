using System.ComponentModel;

namespace Bing.Offices.Imports;

/// <summary>逐行消费普通 XLSX，不累计实体；调用方输入流始终保持打开。</summary>
public interface IExcelStreamingImporter
{
    /// <summary>同步读取单 Sheet。消费者异常原样传播；Stop 返回不完整摘要。</summary>
    ExcelStreamingImportSummary Import<T>(ExcelStreamingImportRequest<T> request, Stream source,
        Func<ExcelImportRow<T>, ExcelRowReadDecision> consumer, ExcelStreamingImportOptions options = null,
        CancellationToken cancellationToken = default) where T : class, new();
    /// <summary>异步读取并等待每次消费者完成后再前进；取消原样传播。</summary>
    Task<ExcelStreamingImportSummary> ImportAsync<T>(ExcelStreamingImportRequest<T> request, Stream source,
        Func<ExcelImportRow<T>, CancellationToken, Task<ExcelRowReadDecision>> consumer,
        ExcelStreamingImportOptions options = null, CancellationToken cancellationToken = default)
        where T : class, new();
}

/// <summary>逐行消费者对后续读取的决定。</summary>
public enum ExcelRowReadDecision
{
    /// <summary>继续读取下一行。</summary>
    Continue,
    /// <summary>停止读取并返回不完整结果。</summary>
    Stop
}

/// <summary>拥有独立数据的单行结果；失败行的 Item 为 null。</summary>
public sealed class ExcelImportRow<T> where T : class, new()
{
    /// <summary>创建提供程序的独立行结果，复制错误集合。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelImportRow(int rowIndex, T item, IReadOnlyList<ExcelImportError> errors)
    {
        RowIndex = rowIndex;
        Errors = Array.AsReadOnly((errors ?? Array.Empty<ExcelImportError>()).ToArray());
        Item = Errors.Count == 0 ? item : null;
    }
    /// <summary>数据在工作表中的零基行索引。</summary>
    public int RowIndex { get; }
    /// <summary>成功实体；存在错误时为 null。</summary>
    public T Item { get; }
    /// <summary>当前行错误；错误内部行列号沿用一基约定。</summary>
    public IReadOnlyList<ExcelImportError> Errors { get; }
    /// <summary>是否具有可消费的成功实体。</summary>
    public bool IsValid => Item != null && Errors.Count == 0;
}

/// <summary>不包含完整数据集的导入统计结果。</summary>
public sealed class ExcelStreamingImportSummary
{
    /// <summary>创建提供程序统计结果；错误样本不代表总错误数量。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelStreamingImportSummary(long rowsRead, long successfulRows, long failedRows, long deliveredRows,
        long totalErrors, IReadOnlyList<ExcelImportError> errors, ImportStopReason stopReason)
    {
        RowsRead = rowsRead; SuccessfulRows = successfulRows; FailedRows = failedRows;
        DeliveredRows = deliveredRows; TotalErrors = totalErrors;
        Errors = Array.AsReadOnly((errors ?? Array.Empty<ExcelImportError>()).ToArray());
        StopReason = stopReason;
    }
    /// <summary>参与处理的数据行数，不含表头和按配置跳过的空行。</summary>
    public long RowsRead { get; }
    /// <summary>成功转换并通过校验的行数。</summary>
    public long SuccessfulRows { get; }
    /// <summary>转换或校验失败的行数。</summary>
    public long FailedRows { get; }
    /// <summary>消费者已正常返回的回调次数，包括失败行和 Stop 所在行。</summary>
    public long DeliveredRows { get; }
    /// <summary>已发现的错误总数。</summary>
    public long TotalErrors { get; }
    /// <summary>有界错误样本。</summary>
    public IReadOnlyList<ExcelImportError> Errors { get; }
    /// <summary>错误样本是否少于发现的错误总数。</summary>
    public bool ErrorsTruncated => TotalErrors > Errors.Count;
    /// <summary>结束或提前停止的原因。</summary>
    public ImportStopReason StopReason { get; }
    /// <summary>是否按配置读完；普通行错误不影响此标志。</summary>
    public bool IsComplete => StopReason == ImportStopReason.None;
    /// <summary>是否按配置读完且没有错误。</summary>
    public bool IsSuccess => IsComplete && TotalErrors == 0;
}

/// <summary>低内存导入额度；文件和 ZIP 资源限制与独立的驻留状态额度同时生效。</summary>
public sealed class ExcelStreamingImportOptions
{
    /// <summary>既有输入、ZIP、XML 资源限制；必须提供有限的包读取额度。</summary>
    public ExcelResourceLimits ResourceLimits { get; init; } = new();
    /// <summary>字符串驻留及缓存的总计量预算，默认 8 MiB（UTF-16 负载加每项管理开销）。</summary>
    public long SharedStringMemoryBytes { get; init; } = 8L * 1024 * 1024;
    /// <summary>单元格文本字符上限。</summary>
    public int MaxCellCharacters { get; init; } = 32767;
    /// <summary>单行文本字符上限。</summary>
    public int MaxRowCharacters { get; init; } = 262144;
    /// <summary>格式表中允许的样式条目数。</summary>
    public int MaxStyles { get; init; } = 4096;
    /// <summary>工作簿元数据中允许的工作表数量。</summary>
    public int MaxSheets { get; init; } = 32;
    /// <summary>样式表允许的字体数量；字体内容不保留。</summary>
    public int MaxFonts { get; init; } = 256;
    /// <summary>唯一值数量上限；与 ResourceLimits 中同名额度取较小值。</summary>
    public int MaxTrackedUniqueValues { get; init; } = 100000;
    /// <summary>唯一值累计字符上限，包含当前行未提交值。</summary>
    public long MaxTrackedUniqueCharacters { get; init; } = 8388608;
    /// <summary>累计错误停止阈值；与 ResourceLimits.MaxErrors 取较小值。</summary>
    public int MaxErrors { get; init; } = 1000;
    /// <summary>摘要保留的最大错误样本数；达到此数不停止读取。</summary>
    public int MaxErrorSamples { get; init; } = 100;
    /// <summary>输入、字符串数据及索引的临时文件总额度，默认 512 MiB。</summary>
    public long MaxTemporaryBytes { get; init; } = 512L * 1024 * 1024;
    /// <summary>库自有临时文件父目录；null 使用系统临时目录。</summary>
    public string TempDirectory { get; init; }
    /// <summary>验证有界读取需要的额度。</summary>
    public void Validate()
    {
        if (ResourceLimits == null) throw new ArgumentNullException(nameof(ResourceLimits));
        ResourceLimits.Validate();
        if (!ResourceLimits.MaxInputBytes.HasValue || !ResourceLimits.MaxZipEntries.HasValue
            || !ResourceLimits.MaxZipEntryUncompressedBytes.HasValue || !ResourceLimits.MaxZipTotalUncompressedBytes.HasValue
            || !ResourceLimits.MaxXmlCharacters.HasValue || !ResourceLimits.MaxXmlDepth.HasValue)
            throw new ArgumentException("流式导入不能关闭输入、ZIP 条目、解压和 XML 基础额度。", nameof(ResourceLimits));
        if (SharedStringMemoryBytes < 1 || MaxCellCharacters < 1 || MaxCellCharacters > 32767
            || MaxRowCharacters < 1 || MaxStyles < 1 || MaxSheets < 1 || MaxFonts < 1 || MaxTrackedUniqueValues < 1
            || MaxTrackedUniqueCharacters < 1 || MaxErrors < 1 || MaxErrorSamples < 0 || MaxTemporaryBytes < 1)
            throw new ArgumentOutOfRangeException(nameof(ExcelStreamingImportOptions));
    }
}

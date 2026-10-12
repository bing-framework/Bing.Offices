using Bing.Offices.ClosedXml.Internals;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Imports;

/// <summary>
/// 将 ClosedXML 完整工作簿导入结果交付为批次。
/// </summary>
/// <remarks>
/// ClosedXML 需要保留 DOM 以生成失败工作簿，因此批次回调在完整导入和失败产物提交后执行。
/// </remarks>
internal sealed class ClosedXmlWorkbookBatchImporter : IExcelWorkbookBatchImporter
{
    /// <summary>
    /// 执行完整工作簿导入的 ClosedXML 导入器。
    /// </summary>
    private readonly ClosedXmlExcelImporter _importer;

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlWorkbookBatchImporter" /> 类型的实例。
    /// </summary>
    /// <param name="importer">负责执行完整 Workbook 导入的 ClosedXML 导入器。</param>
    internal ClosedXmlWorkbookBatchImporter(ClosedXmlExcelImporter importer)
    {
        _importer = importer ?? throw new ArgumentNullException(nameof(importer));
    }

    /// <inheritdoc />
    public ExcelWorkbookBatchImportSummary ImportWorkbookBatches<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, Action<ExcelWorkbookImportBatch> onBatch,
        CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        cancellationToken.ThrowIfCancellationRequested();

        using var buffered = BufferSource(source, request.ResourceLimits?.MaxInputBytes,
            cancellationToken);
        var result = _importer.Import(buffered, request, cancellationToken);
        buffered.Position = 0;
        var physicalSheets = ReadPhysicalSheetNames(buffered, cancellationToken);
        return ExcelWorkbookBatchProjection.Deliver(result, request, physicalSheets, onBatch,
            cancellationToken);
    }

    /// <inheritdoc />
    public async Task<ExcelWorkbookBatchImportSummary> ImportWorkbookBatchesAsync<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request,
        Func<ExcelWorkbookImportBatch, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        cancellationToken.ThrowIfCancellationRequested();

        await using var buffered = await BufferSourceAsync(source, request.ResourceLimits?.MaxInputBytes,
            cancellationToken).ConfigureAwait(false);
        var result = await _importer.ImportAsync(buffered, request, cancellationToken)
            .ConfigureAwait(false);
        buffered.Position = 0;
        var physicalSheets = ReadPhysicalSheetNames(buffered, cancellationToken);
        return await ExcelWorkbookBatchProjection.DeliverAsync(result, request, physicalSheets,
            onBatch, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// 同步暂存输入工作簿。
    /// </summary>
    private static MemoryStream BufferSource(Stream source, long? maximum,
        CancellationToken cancellationToken)
    {
        var buffer = new MemoryStream();
        var bytes = new byte[81920];
        long total = 0;
        try
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var read = source.Read(bytes, 0, bytes.Length);
                if (read == 0)
                    break;
                total += read;
                if (maximum.HasValue && total > maximum.Value)
                    throw new BingOfficesResourceLimitException("输入工作簿超过最大字节数。",
                        provider: "ClosedXML", operation: BingOfficesOperation.Import,
                        stage: BingOfficesStage.Open);
                buffer.Write(bytes, 0, read);
            }
            buffer.Position = 0;
            return buffer;
        }
        catch
        {
            buffer.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 异步暂存输入工作簿。
    /// </summary>
    private static async Task<MemoryStream> BufferSourceAsync(Stream source, long? maximum,
        CancellationToken cancellationToken)
    {
        var buffer = new MemoryStream();
        var bytes = new byte[81920];
        long total = 0;
        try
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var read = await source.ReadAsync(bytes, 0, bytes.Length, cancellationToken)
                    .ConfigureAwait(false);
                if (read == 0)
                    break;
                total += read;
                if (maximum.HasValue && total > maximum.Value)
                    throw new BingOfficesResourceLimitException("输入工作簿超过最大字节数。",
                        provider: "ClosedXML", operation: BingOfficesOperation.Import,
                        stage: BingOfficesStage.Open);
                await buffer.WriteAsync(bytes, 0, read, cancellationToken).ConfigureAwait(false);
            }
            buffer.Position = 0;
            return buffer;
        }
        catch
        {
            buffer.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 读取源工作簿的物理工作表顺序。
    /// </summary>
    private static IReadOnlyList<string> ReadPhysicalSheetNames(Stream source,
        CancellationToken cancellationToken)
    {
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            source.Position = 0;
            using var workbook = new XLWorkbook(source);
            return workbook.Worksheets.Select(sheet => sheet.Name).ToArray();
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesImportException("ClosedXML 工作表顺序读取失败。", exception,
                "ClosedXML", BingOfficesStage.Read);
        }
    }
}

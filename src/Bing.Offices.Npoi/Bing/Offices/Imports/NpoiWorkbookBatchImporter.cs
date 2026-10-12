using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using NPOI.SS.UserModel;

namespace Bing.Offices.Imports;

/// <summary>
/// 将 NPOI 完整工作簿导入结果交付为批次。
/// </summary>
/// <remarks>
/// NPOI 需要保留 DOM 以生成失败工作簿，因此批次回调在完整导入和失败产物提交后执行。
/// </remarks>
internal sealed class NpoiWorkbookBatchImporter : IExcelWorkbookBatchImporter
{
    /// <summary>
    /// 执行完整工作簿导入的 NPOI 导入器。
    /// </summary>
    private readonly NpoiExcelImporter _importer;

    /// <summary>
    /// 初始化一个 <see cref="NpoiWorkbookBatchImporter" /> 类型的实例。
    /// </summary>
    /// <param name="importer">负责执行完整 Workbook 导入的 NPOI 导入器。</param>
    internal NpoiWorkbookBatchImporter(NpoiExcelImporter importer)
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

        using var buffered = new MemoryStream();
        NpoiStreamCopier.Copy(source, buffered, cancellationToken,
            request.ResourceLimits?.MaxInputBytes);
        buffered.Position = 0;
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

        using var buffered = new MemoryStream();
        await NpoiStreamCopier.CopyAsync(source, buffered, cancellationToken,
            request.ResourceLimits?.MaxInputBytes).ConfigureAwait(false);
        buffered.Position = 0;
        var result = await _importer.ImportAsync(buffered, request, cancellationToken)
            .ConfigureAwait(false);
        buffered.Position = 0;
        var physicalSheets = ReadPhysicalSheetNames(buffered, cancellationToken);
        return await ExcelWorkbookBatchProjection.DeliverAsync(result, request, physicalSheets,
            onBatch, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// 读取源工作簿的物理工作表顺序。
    /// </summary>
    /// <param name="source">可定位的工作簿流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>按物理顺序排列的工作表名称。</returns>
    private static IReadOnlyList<string> ReadPhysicalSheetNames(Stream source,
        CancellationToken cancellationToken)
    {
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            source.Position = 0;
            using var workbook = WorkbookFactory.Create(source);
            var names = new string[workbook.NumberOfSheets];
            for (var index = 0; index < names.Length; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                names[index] = workbook.GetSheetName(index);
            }
            return names;
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
            throw new BingOfficesImportException("NPOI 工作表顺序读取失败。", exception, "NPOI",
                BingOfficesStage.Read);
        }
    }
}

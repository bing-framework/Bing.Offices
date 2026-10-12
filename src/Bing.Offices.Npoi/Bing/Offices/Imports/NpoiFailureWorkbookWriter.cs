using System.Diagnostics;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.IO;
using NPOI.SS.UserModel;

namespace Bing.Offices.Imports;

/// <summary>
/// 生成并写出导入失败工作簿，隔离失败产物的复制、注释和摘要逻辑。
/// </summary>
internal static class NpoiFailureWorkbookWriter
{
    /// <summary>
    /// 异步写出失败工作簿，并保持文件提交和目标流复制的异步边界。
    /// </summary>
    /// <param name="workbook">原始导入工作簿。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">已收集的导入错误。</param>
    /// <param name="resolvedSheetRequests">实际解析 Sheet 名称到请求的映射。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="fileExportCommitter">文件路径输出使用的提交服务；为空时使用默认实现。</param>
    /// <param name="requiredRowsBySheet">实体布局要求保留的零基结构行；通用 Workbook 为 null。</param>
    /// <param name="workbookAlreadyFiltered">工作簿已由调用方按失败行筛选时跳过再次复制。</param>
    internal static async Task WriteAsync(IWorkbook workbook, ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests,
        CancellationToken cancellationToken, IFileExportCommitter fileExportCommitter = null,
        IReadOnlyDictionary<string, IReadOnlyList<int>> requiredRowsBySheet = null,
        bool workbookAlreadyFiltered = false)
    {
        if (options == null || options.Mode == ExcelImportFailureWorkbookMode.None || errors.Count == 0)
            return;
        cancellationToken.ThrowIfCancellationRequested();
        using var serialized = new MemoryStream();
        Write(workbook, options, errors, resolvedSheetRequests, cancellationToken,
            new SystemFailureWorkbookFileSystem(), serialized, fileExportCommitter: null,
            requiredRowsBySheet: requiredRowsBySheet,
            workbookAlreadyFiltered: workbookAlreadyFiltered);
        serialized.Position = 0;
        try
        {
            if (!string.IsNullOrWhiteSpace(options.DestinationPath))
            {
                await (fileExportCommitter ?? new DefaultFileExportCommitter()).CommitAsync(
                    options.DestinationPath,
                    (destination, token) => CopySerializedAsync(serialized, destination, token),
                    cancellationToken, "FailureWorkbook").ConfigureAwait(false);
            }
            else
            {
                await CopySerializedAsync(serialized, options.Destination, cancellationToken)
                    .ConfigureAwait(false);
            }
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
            throw new BingOfficesImportException("失败工作簿异步复制到目标流失败。", exception, "NPOI",
                BingOfficesStage.Write);
        }
    }

    /// <summary>
    /// 将已序列化的失败工作簿异步复制到目标流。
    /// </summary>
    /// <param name="source">已定位到开头的序列化流。</param>
    /// <param name="destination">接收内容的目标流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>表示复制操作的任务。</returns>
    private static async Task CopySerializedAsync(Stream source, Stream destination,
        CancellationToken cancellationToken)
    {
        if (destination == null || !destination.CanWrite)
            throw new ArgumentException("失败工作簿目标流不可写入。", nameof(destination));
        source.Position = 0;
        await source.CopyToAsync(destination, 81920, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 根据失败选项写出失败工作簿；没有配置失败产物或没有错误时不执行任何操作。
    /// </summary>
    /// <param name="workbook">原始导入工作簿。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">已收集的导入错误。</param>
    /// <param name="resolvedSheetRequests">实际解析 Sheet 名称到请求的映射。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    internal static void Write(IWorkbook workbook, ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests,
        CancellationToken cancellationToken)
        => Write(workbook, options, errors, resolvedSheetRequests, cancellationToken,
            new SystemFailureWorkbookFileSystem(), null);

    /// <summary>
    /// 使用指定文件系统适配器写出失败工作簿，便于测试临时文件操作。
    /// </summary>
    /// <param name="workbook">原始导入工作簿。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">已收集的导入错误。</param>
    /// <param name="resolvedSheetRequests">实际解析 Sheet 名称到请求的映射。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="fileSystem">临时文件系统适配器。</param>
    /// <param name="destinationOverride">可选的异步 staging 目标流。</param>
    /// <param name="fileExportCommitter">文件路径输出使用的提交服务；为空时使用默认实现。</param>
    /// <param name="requiredRowsBySheet">实体布局要求保留的零基结构行；通用 Workbook 为 null。</param>
    /// <param name="workbookAlreadyFiltered">工作簿已由调用方按失败行筛选时跳过再次复制。</param>
    internal static void Write(IWorkbook workbook, ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests,
        CancellationToken cancellationToken, IFailureWorkbookFileSystem fileSystem,
        Stream destinationOverride = null, IFileExportCommitter fileExportCommitter = null,
        IReadOnlyDictionary<string, IReadOnlyList<int>> requiredRowsBySheet = null,
        bool workbookAlreadyFiltered = false)
    {
        if (options == null || options.Mode == ExcelImportFailureWorkbookMode.None || errors.Count == 0)
            return;
        if (fileSystem == null)
            throw new ArgumentNullException(nameof(fileSystem));
        cancellationToken.ThrowIfCancellationRequested();
        NpoiFailureWorkbookPreflight.Validate(workbook, errors, resolvedSheetRequests, options,
            requiredRowsBySheet);
        IWorkbook outputWorkbook = workbook;
        IWorkbook independentWorkbook = null;
        try
        {
            try
            {
                if (options.Mode == ExcelImportFailureWorkbookMode.ErrorRowsOnly
                    && !workbookAlreadyFiltered)
                    outputWorkbook = independentWorkbook = NpoiFailureWorkbookCopier.CreateErrorRowsWorkbook(workbook, errors,
                        resolvedSheetRequests, options, cancellationToken, requiredRowsBySheet);
                else
                    NpoiFailureWorkbookAnnotationWriter.AnnotateErrors(outputWorkbook, errors,
                        options.CommentConflictPolicy);
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
                throw new BingOfficesImportException("失败工作簿内容生成失败。", exception, "NPOI",
                    BingOfficesStage.Write);
            }

            try
            {
                NpoiFailureWorkbookAnnotationWriter.WriteSummary(outputWorkbook, errors, cancellationToken);
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
                throw new BingOfficesImportException("失败工作簿错误摘要写入失败。", exception, "NPOI",
                    BingOfficesStage.Write);
            }
            if (destinationOverride != null)
            {
                SerializeToDestination(outputWorkbook, destinationOverride, options, cancellationToken);
                return;
            }
            var temporaryDirectory = options.TemporaryDirectory ?? Path.GetTempPath();
            try
            {
                fileSystem.CreateDirectory(temporaryDirectory);
            }
            catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException)
            {
                throw new BingOfficesImportException("失败工作簿临时目录创建失败。", exception, "NPOI",
                    BingOfficesStage.Open);
            }
            var temporaryPath = Path.Combine(temporaryDirectory, $"bing-offices-failure-{Guid.NewGuid():N}.tmp");
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                Stream output;
                try
                {
                    output = fileSystem.CreateFile(temporaryPath);
                }
                catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException)
                {
                    throw new BingOfficesImportException("失败工作簿临时文件创建失败。", exception, "NPOI",
                        BingOfficesStage.Open);
                }
                using (output)
                using (var limitedOutput = new NpoiFailureWorkbookSerialization.LimitedWriteStream(output,
                    options.MaxSerializedBytes))
                {
                    try
                    {
                        outputWorkbook.Write(limitedOutput, false);
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception exception) when (exception is not OutOfMemoryException
                        && exception is not StackOverflowException)
                    {
                        var fatalException = NpoiFailureWorkbookSerialization.FindFatalException(exception);
                        if (fatalException != null)
                        {
                            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(fatalException).Throw();
                            throw;
                        }
                        var limitException = NpoiFailureWorkbookSerialization.FindLimitException(exception);
                        if (limitException != null)
                            throw limitException;
                        if (exception is BingOfficesException)
                            throw;
                        throw new BingOfficesImportException("失败工作簿序列化失败。", exception, "NPOI",
                            BingOfficesStage.Serialize);
                    }
                    try
                    {
                        limitedOutput.Flush();
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
                        throw new BingOfficesImportException("失败工作簿序列化写入失败。", exception, "NPOI",
                            BingOfficesStage.Serialize);
                    }
                    if (options.MaxSerializedBytes.HasValue && output.Length > options.MaxSerializedBytes.Value)
                        throw new BingOfficesResourceLimitException(
                            $"失败工作簿超过最大序列化字节数: {options.MaxSerializedBytes.Value}",
                            provider: "NPOI", operation: BingOfficesOperation.Import,
                            stage: BingOfficesStage.Serialize);
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        output.Position = 0;
                    }
                    catch (Exception exception) when (exception is not OperationCanceledException
                        && exception is not OutOfMemoryException && exception is not StackOverflowException)
                    {
                        throw new BingOfficesImportException("失败工作簿临时文件读取准备失败。", exception, "NPOI",
                            BingOfficesStage.Write);
                    }
                    try
                    {
                        if (destinationOverride == null && options.DestinationPath != null)
                        {
                            (fileExportCommitter ?? new DefaultFileExportCommitter()).Commit(options.DestinationPath,
                                destination => NpoiFailureWorkbookSerialization.WriteStream(destination, output,
                                    cancellationToken), cancellationToken, "FailureWorkbook");
                        }
                        else
                        {
                            NpoiFailureWorkbookSerialization.WriteStream(destinationOverride ?? options.Destination,
                                output, cancellationToken);
                        }
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
                        throw new BingOfficesImportException("失败工作簿复制到目标流失败。", exception, "NPOI",
                            BingOfficesStage.Write);
                    }
                }
            }
            catch (OperationCanceledException exception)
            {
                DeleteTemporaryFile(options, temporaryPath, exception, fileSystem);
                throw;
            }
            catch (Exception exception) when (exception is not OutOfMemoryException
                && exception is not StackOverflowException)
            {
                DeleteTemporaryFile(options, temporaryPath, exception, fileSystem);
                throw;
            }
            DeleteTemporaryFile(options, temporaryPath, null, fileSystem);
        }
        finally
        {
            independentWorkbook?.Close();
        }
    }

    /// <summary>
    /// 将异步 staging 目标作为唯一序列化目标，避免再创建第二个失败工作簿临时文件。
    /// </summary>
    /// <param name="outputWorkbook">待序列化的失败工作簿。</param>
    /// <param name="destination">接收序列化内容的目标流。</param>
    /// <param name="options">失败工作簿输出和大小限制选项。</param>
    /// <param name="cancellationToken">序列化完成后检查的取消令牌。</param>
    private static void SerializeToDestination(IWorkbook outputWorkbook, Stream destination,
        ExcelImportFailureOptions options, CancellationToken cancellationToken)
    {
        if (!destination.CanWrite)
            throw new ArgumentException("失败工作簿目标流不可写入。", nameof(destination));
        using var limitedOutput = new NpoiFailureWorkbookSerialization.LimitedWriteStream(destination,
            options.MaxSerializedBytes);
        try
        {
            outputWorkbook.Write(limitedOutput, false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var fatalException = NpoiFailureWorkbookSerialization.FindFatalException(exception);
            if (fatalException != null)
            {
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(fatalException).Throw();
                throw;
            }
            var limitException = NpoiFailureWorkbookSerialization.FindLimitException(exception);
            if (limitException != null)
                throw limitException;
            if (exception is BingOfficesException)
                throw;
            throw new BingOfficesImportException("失败工作簿序列化失败。", exception, "NPOI",
                BingOfficesStage.Serialize);
        }
        try
        {
            limitedOutput.Flush();
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
            throw new BingOfficesImportException("失败工作簿序列化写入失败。", exception, "NPOI",
                BingOfficesStage.Serialize);
        }
        if (options.MaxSerializedBytes.HasValue && destination.Length > options.MaxSerializedBytes.Value)
            throw new BingOfficesResourceLimitException(
                $"失败工作簿超过最大序列化字节数: {options.MaxSerializedBytes.Value}",
                provider: "NPOI", operation: BingOfficesOperation.Import,
                stage: BingOfficesStage.Serialize);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 删除失败工作簿临时文件，并将清理失败写入诊断或主异常。
    /// </summary>
    /// <param name="options">失败工作簿输出选项和诊断接收器。</param>
    /// <param name="temporaryPath">待删除的临时文件路径。</param>
    /// <param name="primaryException">提交失败时的主异常。</param>
    /// <param name="fileSystem">临时文件系统适配器。</param>
    private static void DeleteTemporaryFile(ExcelImportFailureOptions options, string temporaryPath,
        Exception primaryException, IFailureWorkbookFileSystem fileSystem)
    {
        try
        {
            fileSystem.Delete(temporaryPath);
        }
        catch (Exception cleanupException) when (cleanupException is IOException
            || cleanupException is UnauthorizedAccessException)
        {
            var diagnostic = new ExcelImportFailureDiagnostic("FailureWorkbookTemporaryCleanupFailed",
                temporaryPath, cleanupException);
            if (primaryException != null)
                primaryException.Data["Bing.Offices.FailureWorkbook.TemporaryCleanupException"] = cleanupException;
            if (options.DiagnosticSink != null)
            {
                try
                {
                    options.DiagnosticSink(diagnostic);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception diagnosticException) when (diagnosticException is not OutOfMemoryException
                    && diagnosticException is not StackOverflowException)
                {
                    Trace.WriteLine($"失败工作簿诊断接收器执行失败: {diagnosticException.GetType().Name}");
                }
            }
            else if (primaryException == null)
            {
                throw new BingOfficesImportException("失败工作簿临时文件清理失败。", cleanupException, "NPOI",
                    BingOfficesStage.Cleanup);
            }
        }
    }


}

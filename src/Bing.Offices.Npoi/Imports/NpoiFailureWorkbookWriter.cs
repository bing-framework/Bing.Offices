using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Metadata;
using Bing.Offices.Npoi.Extensions;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// 生成并写出导入失败工作簿，隔离失败产物的复制、注释和摘要逻辑。
/// </summary>
internal static class NpoiFailureWorkbookWriter
{
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
            new SystemFailureWorkbookFileSystem());

    /// <summary>使用指定文件系统适配器写出失败工作簿，便于测试临时文件操作。</summary>
    /// <param name="workbook">原始导入工作簿。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">已收集的导入错误。</param>
    /// <param name="resolvedSheetRequests">实际解析 Sheet 名称到请求的映射。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="fileSystem">临时文件系统适配器。</param>
    internal static void Write(IWorkbook workbook, ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests,
        CancellationToken cancellationToken, IFailureWorkbookFileSystem fileSystem)
    {
        if (options == null || options.Mode == ExcelImportFailureWorkbookMode.None || errors.Count == 0)
            return;
        if (fileSystem == null)
            throw new ArgumentNullException(nameof(fileSystem));
        cancellationToken.ThrowIfCancellationRequested();
        NpoiFailureWorkbookPreflight.Validate(workbook, errors, resolvedSheetRequests, options);
        IWorkbook outputWorkbook = workbook;
        IWorkbook independentWorkbook = null;
        try
        {
            try
            {
                if (options.Mode == ExcelImportFailureWorkbookMode.ErrorRowsOnly)
                    outputWorkbook = independentWorkbook = NpoiFailureWorkbookCopier.CreateErrorRowsWorkbook(workbook, errors,
                        resolvedSheetRequests, options, cancellationToken);
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
                        NpoiFailureWorkbookSerialization.WriteStream(options.Destination, output,
                            cancellationToken);
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

    /// <summary>删除失败工作簿临时文件，并将清理失败写入诊断或主异常。</summary>
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

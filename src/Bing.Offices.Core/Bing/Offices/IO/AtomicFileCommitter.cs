using Bing.Offices.Exceptions;

namespace Bing.Offices.IO;

/// <summary>
/// 通过同目录临时文件和替换操作提交导出文件的工具。
/// </summary>
internal static class AtomicFileCommitter
{
    /// <summary>
    /// 生产环境使用的文件系统适配器；内部重载可替换以进行确定性测试。
    /// </summary>
    private static readonly IAtomicFileSystem DefaultFileSystem = new SystemAtomicFileSystem();

    /// <summary>
    /// 将写入结果原子提交到目标路径。
    /// </summary>
    /// <param name="path">最终输出文件路径。</param>
    /// <param name="write">向临时输出流写入内容的操作。</param>
    /// <param name="cancellationToken">提交过程检查的取消令牌。</param>
    /// <param name="format">用于临时文件清理错误上下文的格式名称。</param>
    public static void Commit(string path, Action<Stream> write, CancellationToken cancellationToken,
        string format)
        => Commit(path, write, cancellationToken, format, DefaultFileSystem);

    /// <summary>
    /// 将异步写入结果原子提交到目标路径。
    /// </summary>
    /// <param name="path">最终输出文件路径。</param>
    /// <param name="writeAsync">向临时输出流异步写入内容的操作。</param>
    /// <param name="cancellationToken">提交过程检查的取消令牌。</param>
    /// <param name="format">用于临时文件清理错误上下文的格式名称。</param>
    internal static Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
        CancellationToken cancellationToken, string format)
        => CommitAsync(path, writeAsync, cancellationToken, format, DefaultFileSystem);

    /// <summary>
    /// 将异步写入结果原子提交到目标路径。
    /// </summary>
    /// <param name="path">最终输出文件路径。</param>
    /// <param name="writeAsync">向临时输出流异步写入内容的操作。</param>
    /// <param name="cancellationToken">提交过程检查的取消令牌。</param>
    /// <param name="format">用于临时文件清理错误上下文的格式名称。</param>
    /// <param name="fileSystem">负责文件创建、替换、移动和清理的适配器。</param>
    /// <remarks>只有写入完成并刷新临时文件后才替换最终目标；取消或失败时会清理临时文件。</remarks>
    internal static async Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
        CancellationToken cancellationToken, string format, IAtomicFileSystem fileSystem)
    {
        if (fileSystem == null)
            throw new ArgumentNullException(nameof(fileSystem));
        if (writeAsync == null)
            throw new ArgumentNullException(nameof(writeAsync));
        var temporaryPath = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        var writingContent = false;
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            using (var destination = fileSystem.CreateFile(temporaryPath))
            {
                writingContent = true;
                await writeAsync(destination, cancellationToken).ConfigureAwait(false);
                writingContent = false;
                await fileSystem.FlushAsync(destination, cancellationToken).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (fileSystem.Exists(path))
                fileSystem.Replace(temporaryPath, path);
            else
                fileSystem.Move(temporaryPath, path);
            temporaryPath = null;
        }
        catch (OperationCanceledException exception)
        {
            Cleanup(temporaryPath, format, exception, fileSystem);
            if (cancellationToken.IsCancellationRequested
                && exception.GetType() != typeof(OperationCanceledException))
                throw new OperationCanceledException(cancellationToken);
            throw;
        }
        catch (Exception exception) when (writingContent)
        {
            Cleanup(temporaryPath, format, exception, fileSystem);
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesFileCommitException(
                $"{format} 文件提交失败。", exception, format, BingOfficesStage.Commit);
            Cleanup(temporaryPath, format, translated, fileSystem);
            throw translated;
        }
    }

    /// <summary>
    /// 将写入结果原子提交到目标路径。
    /// </summary>
    /// <param name="path">最终输出文件路径。</param>
    /// <param name="write">向临时输出流写入内容的操作。</param>
    /// <param name="cancellationToken">提交过程检查的取消令牌。</param>
    /// <param name="format">用于临时文件清理错误上下文的格式名称。</param>
    /// <param name="fileSystem">负责文件创建、替换、移动和清理的适配器。</param>
    internal static void Commit(string path, Action<Stream> write, CancellationToken cancellationToken,
        string format, IAtomicFileSystem fileSystem)
    {
        if (fileSystem == null)
            throw new ArgumentNullException(nameof(fileSystem));
        var temporaryPath = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        var writingContent = false;
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            using (var destination = fileSystem.CreateFile(temporaryPath))
            {
                writingContent = true;
                write(destination);
                writingContent = false;
                fileSystem.Flush(destination);
                cancellationToken.ThrowIfCancellationRequested();
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (fileSystem.Exists(path))
                fileSystem.Replace(temporaryPath, path);
            else
                fileSystem.Move(temporaryPath, path);
            temporaryPath = null;
        }
        catch (OperationCanceledException exception)
        {
            Cleanup(temporaryPath, format, exception, fileSystem);
            if (cancellationToken.IsCancellationRequested
                && exception.GetType() != typeof(OperationCanceledException))
                throw new OperationCanceledException(cancellationToken);
            throw;
        }
        catch (Exception exception) when (writingContent)
        {
            // 内容生成属于导出公共边界，不能伪装成文件系统提交失败。
            Cleanup(temporaryPath, format, exception, fileSystem);
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesFileCommitException(
                $"{format} 文件提交失败。", exception, format, BingOfficesStage.Commit);
            Cleanup(temporaryPath, format, translated, fileSystem);
            throw translated;
        }
    }

    /// <summary>
    /// 尝试清理失败提交遗留的临时文件，并保留原始异常作为主失败原因。
    /// </summary>
    /// <param name="temporaryPath">待删除的临时文件路径。</param>
    /// <param name="format">输出格式名称。</param>
    /// <param name="primaryException">导致提交失败的原始异常。</param>
    /// <param name="fileSystem">执行删除操作的文件系统适配器。</param>
    private static void Cleanup(string temporaryPath, string format, Exception primaryException,
        IAtomicFileSystem fileSystem)
    {
        if (temporaryPath == null)
            return;
        try
        {
            fileSystem.Delete(temporaryPath);
        }
        catch (Exception cleanupException) when (cleanupException is IOException
            || cleanupException is UnauthorizedAccessException)
        {
            if (primaryException != null)
            {
                primaryException.Data[$"Bing.Offices.{format}.TemporaryCleanupException"] = cleanupException;
                return;
            }
            throw new IOException($"{format} 导出临时文件清理失败。", cleanupException);
        }
    }
}

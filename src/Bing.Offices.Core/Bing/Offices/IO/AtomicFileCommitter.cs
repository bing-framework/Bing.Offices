using System;
using System.IO;
using System.Threading;
using Bing.Offices.Exceptions;

namespace Bing.Offices.IO;

/// <summary>通过同目录临时文件和替换操作提交导出文件的工具。</summary>
internal static class AtomicFileCommitter
{
    /// <summary>生产环境使用的文件系统适配器；内部重载可替换以进行确定性测试。</summary>
    private static readonly IAtomicFileSystem DefaultFileSystem = new SystemAtomicFileSystem();

    /// <summary>将写入结果原子提交到目标路径。</summary>
    /// <param name="path">最终输出文件路径。</param>
    /// <param name="write">向临时输出流写入内容的操作。</param>
    /// <param name="cancellationToken">提交过程检查的取消令牌。</param>
    /// <param name="format">用于临时文件清理错误上下文的格式名称。</param>
    public static void Commit(string path, Action<Stream> write, CancellationToken cancellationToken,
        string format)
        => Commit(path, write, cancellationToken, format, DefaultFileSystem);

    /// <summary>使用指定文件系统适配器将写入结果原子提交到目标路径。</summary>
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
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            using (var destination = fileSystem.CreateFile(temporaryPath))
            {
                write(destination);
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

    /// <summary>尝试清理失败提交遗留的临时文件，并保留原始异常作为主失败原因。</summary>
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

/// <summary>为原子文件提交抽象的最小文件系统操作集合。</summary>
internal interface IAtomicFileSystem
{
    /// <summary>以独占创建方式打开临时输出文件。</summary>
    Stream CreateFile(string path);
    /// <summary>将已写入流的内容持久化到存储介质。</summary>
    void Flush(Stream stream);
    /// <summary>确定目标文件是否存在。</summary>
    bool Exists(string path);
    /// <summary>以临时文件替换已存在的目标文件。</summary>
    void Replace(string sourcePath, string destinationPath);
    /// <summary>将临时文件移动为此前不存在的目标文件。</summary>
    void Move(string sourcePath, string destinationPath);
    /// <summary>删除提交失败遗留的临时文件。</summary>
    void Delete(string path);
}

/// <summary>基于 <see cref="File"/> 的生产文件系统适配器。</summary>
internal sealed class SystemAtomicFileSystem : IAtomicFileSystem
{
    /// <inheritdoc />
    public Stream CreateFile(string path) => new FileStream(path, FileMode.CreateNew, FileAccess.Write,
        FileShare.None, 4096, FileOptions.SequentialScan);

    /// <inheritdoc />
    public void Flush(Stream stream)
    {
        if (stream is FileStream fileStream)
            fileStream.Flush(true);
        else
            stream.Flush();
    }

    /// <inheritdoc />
    public bool Exists(string path) => File.Exists(path);

    /// <inheritdoc />
    public void Replace(string sourcePath, string destinationPath) => File.Replace(sourcePath, destinationPath, null);

    /// <inheritdoc />
    public void Move(string sourcePath, string destinationPath) => File.Move(sourcePath, destinationPath);

    /// <inheritdoc />
    public void Delete(string path) => File.Delete(path);
}

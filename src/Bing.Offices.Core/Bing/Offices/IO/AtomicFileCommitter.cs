using System;
using System.IO;
using System.Threading;

namespace Bing.Offices.IO;

internal static class AtomicFileCommitter
{
    private static readonly IAtomicFileSystem DefaultFileSystem = new SystemAtomicFileSystem();

    public static void Commit(string path, Action<Stream> write, CancellationToken cancellationToken,
        string format)
        => Commit(path, write, cancellationToken, format, DefaultFileSystem);

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
        catch (Exception exception)
        {
            Cleanup(temporaryPath, format, exception, fileSystem);
            throw;
        }
    }

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

internal interface IAtomicFileSystem
{
    Stream CreateFile(string path);
    void Flush(Stream stream);
    bool Exists(string path);
    void Replace(string sourcePath, string destinationPath);
    void Move(string sourcePath, string destinationPath);
    void Delete(string path);
}

internal sealed class SystemAtomicFileSystem : IAtomicFileSystem
{
    public Stream CreateFile(string path) => new FileStream(path, FileMode.CreateNew, FileAccess.Write,
        FileShare.None, 4096, FileOptions.SequentialScan);

    public void Flush(Stream stream)
    {
        if (stream is FileStream fileStream)
            fileStream.Flush(true);
        else
            stream.Flush();
    }

    public bool Exists(string path) => File.Exists(path);

    public void Replace(string sourcePath, string destinationPath) => File.Replace(sourcePath, destinationPath, null);

    public void Move(string sourcePath, string destinationPath) => File.Move(sourcePath, destinationPath);

    public void Delete(string path) => File.Delete(path);
}

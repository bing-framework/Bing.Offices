using Bing.Offices.Exceptions;

namespace Bing.Offices.IO;

/// <summary>
/// 基于 <see cref="File"/> 的生产文件系统适配器。
/// </summary>
internal sealed class SystemAtomicFileSystem : IAtomicFileSystem
{
    /// <inheritdoc />
    public Stream CreateFile(string path) => new FileStream(path, FileMode.CreateNew, FileAccess.Write,
        FileShare.None, 4096, FileOptions.SequentialScan | FileOptions.Asynchronous);

    /// <inheritdoc />
    public void Flush(Stream stream)
    {
        if (stream is FileStream fileStream)
            fileStream.Flush(true);
        else
            stream.Flush();
    }

    /// <inheritdoc />
    public async Task FlushAsync(Stream stream, CancellationToken cancellationToken)
    {
        if (stream is FileStream fileStream)
        {
            await fileStream.FlushAsync(cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            fileStream.Flush(true);
            return;
        }

        await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
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

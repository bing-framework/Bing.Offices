using System.ComponentModel;

namespace Bing.Offices.IO;

/// <summary>
/// 默认文件提交实现。
/// </summary>
/// <remarks>
/// 异常观察由具体导出器负责。
/// </remarks>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class DefaultFileExportCommitter : IFileExportCommitter
{
    /// <inheritdoc />
    public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken, string format)
        => AtomicFileCommitter.Commit(path, write, cancellationToken, format);

    /// <inheritdoc />
    public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
        CancellationToken cancellationToken, string format)
        => AtomicFileCommitter.CommitAsync(path, writeAsync, cancellationToken, format);
}

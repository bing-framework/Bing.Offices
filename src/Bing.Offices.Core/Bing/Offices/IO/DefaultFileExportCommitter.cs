using System;
using System.ComponentModel;
using System.IO;
using System.Threading;

namespace Bing.Offices.IO;

/// <summary>默认文件提交实现；具体 exporter 负责异常观察。</summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class DefaultFileExportCommitter : IFileExportCommitter
{
    /// <inheritdoc />
    public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken, string format)
        => AtomicFileCommitter.Commit(path, write, cancellationToken, format);
}

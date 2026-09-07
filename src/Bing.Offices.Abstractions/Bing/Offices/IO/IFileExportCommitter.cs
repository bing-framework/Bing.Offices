using System;
using System.ComponentModel;
using System.IO;
using System.Threading;

namespace Bing.Offices.IO;

/// <summary>
/// 文件导出提交 SPI；由 exporter 在自身异常观察边界内调用。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IFileExportCommitter
{
    /// <summary>将临时文件内容原子提交到目标路径。</summary>
    /// <param name="path">目标文件路径。</param>
    /// <param name="write">写入临时文件的操作。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="format">导出格式名称。</param>
    void Commit(string path, Action<Stream> write, CancellationToken cancellationToken, string format);
}

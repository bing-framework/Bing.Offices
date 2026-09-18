namespace Bing.Offices.Imports;

/// <summary>
/// 为失败工作簿临时文件创建和清理抽象的文件系统操作。
/// </summary>
internal interface IFailureWorkbookFileSystem
{
    /// <summary>
    /// 确保临时输出目录存在。
    /// </summary>
    /// <param name="path">需要创建的目录路径。</param>
    void CreateDirectory(string path);
    /// <summary>
    /// 以读写方式独占创建临时工作簿文件。
    /// </summary>
    /// <param name="path">需要创建的临时文件路径。</param>
    /// <returns>以读写方式打开且由调用方负责释放的临时文件流。</returns>
    Stream CreateFile(string path);
    /// <summary>
    /// 删除不再需要的临时工作簿文件。
    /// </summary>
    /// <param name="path">需要删除的临时文件路径。</param>
    void Delete(string path);
}

/// <summary>
/// 基于 <see cref="Directory"/> 和 <see cref="File"/> 的失败工作簿文件系统实现。
/// </summary>
internal sealed class SystemFailureWorkbookFileSystem : IFailureWorkbookFileSystem
{
    /// <inheritdoc />
    public void CreateDirectory(string path) => Directory.CreateDirectory(path);

    /// <inheritdoc />
    public Stream CreateFile(string path) => new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite,
        FileShare.None, 81920, FileOptions.SequentialScan);

    /// <inheritdoc />
    public void Delete(string path) => File.Delete(path);
}

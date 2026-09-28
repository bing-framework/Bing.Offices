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

using Bing.Offices.Exceptions;

namespace Bing.Offices.IO;

/// <summary>
/// 为原子文件提交抽象的最小文件系统操作集合。
/// </summary>
internal interface IAtomicFileSystem
{
    /// <summary>
    /// 以独占创建方式打开临时输出文件。
    /// </summary>
    /// <param name="path">要创建的临时文件路径。</param>
    /// <returns>以写入方式打开的临时文件流。</returns>
    Stream CreateFile(string path);
    /// <summary>
    /// 将已写入流的内容持久化到存储介质。
    /// </summary>
    /// <param name="stream">需要持久化的输出流。</param>
    void Flush(Stream stream);
    /// <summary>
    /// 异步写入并完成与同步提交一致的持久化边界。
    /// </summary>
    /// <param name="stream">需要持久化的输出流。</param>
    /// <param name="cancellationToken">刷新过程中检查的取消令牌。</param>
    Task FlushAsync(Stream stream, CancellationToken cancellationToken);
    /// <summary>
    /// 确定目标文件是否存在。
    /// </summary>
    /// <param name="path">要检查的文件路径。</param>
    /// <returns>文件存在时为 <see langword="true" />，否则为 <see langword="false" />。</returns>
    bool Exists(string path);
    /// <summary>
    /// 以临时文件替换已存在的目标文件。
    /// </summary>
    /// <param name="sourcePath">临时源文件路径。</param>
    /// <param name="destinationPath">要替换的目标文件路径。</param>
    void Replace(string sourcePath, string destinationPath);
    /// <summary>
    /// 将临时文件移动为此前不存在的目标文件。
    /// </summary>
    /// <param name="sourcePath">临时源文件路径。</param>
    /// <param name="destinationPath">要创建的目标文件路径。</param>
    void Move(string sourcePath, string destinationPath);
    /// <summary>
    /// 删除提交失败遗留的临时文件。
    /// </summary>
    /// <param name="path">要删除的临时文件路径。</param>
    void Delete(string path);
}

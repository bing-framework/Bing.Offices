namespace Bing.Offices.Imports;

/// <summary>
/// 导入器
/// </summary>
public interface IImporter
{
    /// <summary>
    /// 生成导入模板
    /// </summary>
    /// <typeparam name="T">类型</typeparam>
    /// <returns>二进制字节数组</returns>
    Task<byte[]> GenerateTemplateBytesAsync<T>() where T : class, new();
}

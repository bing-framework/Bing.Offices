namespace Bing.Offices.Exports;

/// <summary>
/// 根据模板导出文件
/// </summary>
public interface IExportFileByTemplate
{
    /// <summary>
    /// 根据模板导出到字节数组
    /// </summary>
    /// <typeparam name="T">类型</typeparam>
    /// <param name="data">数据</param>
    /// <param name="template">模板路径</param>
    /// <returns></returns>
    Task<byte[]> ExportBytesByTemplate<T>(T data, string template) where T : class;
}

namespace Bing.Offices.Exports;

/// <summary>
/// 导出器
/// </summary>
public interface IExporter
{
    /// <summary>
    /// 导出
    /// </summary>
    /// <typeparam name="TEntity">类型</typeparam>
    /// <param name="dataItems">数据</param>
    /// <returns>文件二进制数组</returns>
    Task<byte[]> ExportAsBytesAsync<TEntity>(ICollection<TEntity> dataItems) where TEntity : class, new();

    /// <summary>
    /// 导出
    /// </summary>
    /// <typeparam name="TEntity">类型</typeparam>
    /// <param name="dataItems">数据</param>
    /// <param name="excelFormat">Excel格式</param>
    /// <returns>文件二进制数组</returns>
    Task<byte[]> ExportAsBytesAsync<TEntity>(ICollection<TEntity> dataItems, ExcelFormat excelFormat) where TEntity : class, new();

    /// <summary>
    /// 导出
    /// </summary>
    /// <typeparam name="TEntity">类型</typeparam>
    /// <param name="dataItems">数据</param>
    /// <param name="excelFormat">Excel格式</param>
    /// <param name="sheetIndex">工作簿索引</param>
    /// <returns>文件二进制数组</returns>
    Task<byte[]> ExportAsBytesAsync<TEntity>(ICollection<TEntity> dataItems, ExcelFormat excelFormat, int sheetIndex) where TEntity : class, new();
}

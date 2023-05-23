using Bing.Offices.Internals;

namespace Bing.Offices.Exports;

/// <summary>
/// Excel 导出器基类
/// </summary>
public abstract class ExcelExporterBase : IExcelExporter
{
    /// <summary>
    /// 导出
    /// </summary>
    /// <typeparam name="TEntity">类型</typeparam>
    /// <param name="dataItems">数据</param>
    /// <returns>文件二进制数组</returns>
    public Task<byte[]> ExportAsBytesAsync<TEntity>(ICollection<TEntity> dataItems) where TEntity : class, new() =>
        ExportAsBytesAsync(dataItems, ExcelFormat.Xlsx);

    /// <summary>
    /// 导出
    /// </summary>
    /// <typeparam name="TEntity">类型</typeparam>
    /// <param name="dataItems">数据</param>
    /// <param name="excelFormat">Excel格式</param>
    /// <returns>文件二进制数组</returns>
    public Task<byte[]> ExportAsBytesAsync<TEntity>(ICollection<TEntity> dataItems, ExcelFormat excelFormat)
        where TEntity : class, new() => ExportAsBytesAsync(dataItems, excelFormat, 0);

    /// <summary>
    /// 导出
    /// </summary>
    /// <typeparam name="TEntity">类型</typeparam>
    /// <param name="dataItems">数据</param>
    /// <param name="excelFormat">Excel格式</param>
    /// <param name="sheetIndex">工作簿索引</param>
    /// <returns>文件二进制数组</returns>
    public async Task<byte[]> ExportAsBytesAsync<TEntity>(ICollection<TEntity> dataItems, ExcelFormat excelFormat, int sheetIndex) where TEntity : class, new()
    {
        if (dataItems is null)
            throw new ArgumentNullException(nameof(dataItems));
        var configuration = InternalHelper.GetExcelConfigurationMapping<TEntity>();
        var context = new ExportContext<TEntity>(excelFormat, sheetIndex, configuration);
        return await ExportAsBytesAsync(dataItems, context);
    }

    /// <summary>
    /// 导出
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <param name="dataItems">数据</param>
    /// <param name="context">导出上下文</param>
    protected abstract Task<byte[]> ExportAsBytesAsync<TEntity>(ICollection<TEntity> dataItems, ExportContext<TEntity> context);

    /// <summary>
    /// 根据模板导出到字节数组
    /// </summary>
    /// <typeparam name="T">类型</typeparam>
    /// <param name="data">数据</param>
    /// <param name="template">模板路径</param>
    /// <returns></returns>
    public async Task<byte[]> ExportBytesByTemplate<T>(T data, string template) where T : class
    {
        throw new NotImplementedException();
    }
}

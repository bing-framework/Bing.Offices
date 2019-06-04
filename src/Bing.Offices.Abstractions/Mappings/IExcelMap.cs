namespace Bing.Offices.Abstractions.Mappings
{
    /// <summary>
    /// Excel 映射
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    public interface IExcelMap<T> : IExcelImportMap<T>, IExcelExportMap<T> where T : class
    {
    }
}

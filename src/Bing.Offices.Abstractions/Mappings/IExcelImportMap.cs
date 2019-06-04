namespace Bing.Offices.Abstractions.Mappings
{
    /// <summary>
    /// Excel 导入映射
    /// </summary>
    public interface IExcelImportMap : IExcelCommonMap
    {
    }

    /// <summary>
    /// Excel 导入映射
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    public interface IExcelImportMap<T> : IExcelImportMap where T : class
    {

    }
}

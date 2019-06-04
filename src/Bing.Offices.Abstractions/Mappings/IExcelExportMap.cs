namespace Bing.Offices.Abstractions.Mappings
{
    /// <summary>
    /// Excel 导出映射
    /// </summary>
    public interface IExcelExportMap : IExcelCommonMap
    {
    }

    /// <summary>
    /// Excel 导出映射
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    public interface IExcelExportMap<T> : IExcelExportMap where T : class
    {

    }
}

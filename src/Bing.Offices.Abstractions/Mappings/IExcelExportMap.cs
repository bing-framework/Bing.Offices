using Bing.Offices.Abstractions.Configurations;

namespace Bing.Offices.Abstractions.Mappings
{
    /// <summary>
    /// Excel 导出映射
    /// </summary>
    public interface IExcelExportMap : IExcelCommonMap
    {
        /// <summary>
        /// 初始化
        /// </summary>
        void Init();
    }

    /// <summary>
    /// Excel 导出映射
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    public interface IExcelExportMap<T> : IExcelExportMap where T : class
    {
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="configuration">Excel配置</param>
        void Init(IExcelConfiguration<T> configuration);
    }
}

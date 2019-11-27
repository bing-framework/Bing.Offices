using Bing.Offices.Abstractions;

namespace Bing.Offices.Excel.Mappings
{
    /// <summary>
    /// Excel导出映射
    /// </summary>
    public interface IExcelExportMap : IExportMap
    {
        /// <summary>
        /// 映射配置
        /// </summary>
        /// <param name="context">上下文</param>
        void Map(ExcelContext context);
    }
}

using System.Collections.Generic;

namespace Bing.Offices.Abstractions.Mappings
{
    /// <summary>
    /// Excel 映射上下文
    /// </summary>
    public interface IExcelMapContext
    {
        /// <summary>
        /// 获取导入映射列表
        /// </summary>
        List<IExcelImportMap> GetImportMaps();

        /// <summary>
        /// 获取导出映射列表
        /// </summary>
        List<IExcelExportMap> GetExportMaps();
    }
}

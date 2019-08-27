using System.Collections.Generic;

namespace Bing.Offices.Abstractions.Exports
{
    /// <summary>
    /// 导出选项配置
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    public interface IExportOptions<T> where T : class, new()
    {
        /// <summary>
        /// 数据集合
        /// </summary>
        IList<T> Data { get; set; }

        /// <summary>
        /// 导出Excel类型
        /// </summary>
        ExportType ExportType { get; set; }

        /// <summary>
        /// 工作表名称
        /// </summary>
        string SheetName { get; set; }

        /// <summary>
        /// 表头行索引。默认为0
        /// </summary>
        int HeaderRowIndex { get; set; }

        /// <summary>
        /// 数据行起始索引。默认为1
        /// </summary>
        int DataRowStartIndex { get; set; }

        /// <summary>
        /// 自定义导出提供程序
        /// </summary>
        IExcelExportProvider CustomExportProvider { get; set; }
    }
}

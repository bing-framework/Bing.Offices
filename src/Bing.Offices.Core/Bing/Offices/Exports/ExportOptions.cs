using System.Collections.Generic;
using Bing.Offices.Metadata.Excels;

namespace Bing.Offices.Exports
{
    /// <summary>
    /// 导出选项配置
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    public class ExportOptions<T> : IExportOptions<T> where T : class, new()
    {
        /// <summary>
        /// 行头
        /// </summary>
        public IList<IRow> HeaderRow { get; set; }

        /// <summary>
        /// 数据集合
        /// </summary>
        public IList<T> Data { get; set; }

        /// <summary>
        /// 导出格式
        /// </summary>
        public ExcelFormat ExportFormat { get; set; } = ExcelFormat.Xlsx;

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName { get; set; } = "sheet1";

        /// <summary>
        /// 表头行索引。默认为0
        /// </summary>
        public int HeaderRowIndex { get; set; } = 0;

        /// <summary>
        /// 数据行起始索引。默认为1
        /// </summary>
        public int DataRowStartIndex { get; set; } = 1;

        /// <summary>
        /// 动态列
        /// </summary>
        public IList<string> DynamicColumns { get; set; } = new List<string>();

        /// <summary>
        /// 自定义导出提供程序
        /// </summary>
        public IExcelExportProvider CustomExportProvider { get; set; }

        /// <summary>
        /// 查询总数
        /// </summary>
        public int QueryCount { get; set; } = 10000;
    }
}

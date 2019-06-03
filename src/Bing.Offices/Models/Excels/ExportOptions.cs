using System.Collections.Generic;
using Bing.Offices.Abstractions;
using Bing.Offices.Enums;

namespace Bing.Offices.Models.Excels
{
    /// <summary>
    /// 导出选项
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    public class ExportOptions<T> where T : class, new()
    {
        /// <summary>
        /// 数据集合
        /// </summary>
        public List<T> Data { get; set; }

        /// <summary>
        /// 导出Excel文件类型
        /// </summary>
        public ExcelType ExcelType { get; set; } = ExcelType.Xlsx;

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName { get; set; } = "sheet1";

        /// <summary>
        /// 表头行索引。默认：0
        /// </summary>
        public int HeaderRowIndex { get; set; } = 0;

        /// <summary>
        /// 数据行起始索引。默认：1
        /// </summary>
        public int DataRowStartIndex { get; set; } = 1;

        /// <summary>
        /// 导出提供程序
        /// </summary>
        public IExcelExportProvider ExportProvider { get; set; }
    }
}

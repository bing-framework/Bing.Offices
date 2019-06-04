using System;

namespace Bing.Offices.Configs
{
    /// <summary>
    /// Excel 定义
    /// </summary>
    public class ExcelDefinition
    {
        /// <summary>
        /// 系统编号。必填项
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 全类名
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 类类型
        /// </summary>
        public Type ClassType { get; set; }

        /// <summary>
        /// 工作表名称。导出时，可以不设置
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 默认列宽。导出时，工作表所有的默认列宽，可以不设置
        /// </summary>
        public int DefaultColumnWidth { get; set; }

        /// <summary>
        /// 默认单元格对齐方式。导出时，单元格默认对齐方式，支持：center,left,right
        /// </summary>
        public short DefaultAlign { get; set; }


    }
}

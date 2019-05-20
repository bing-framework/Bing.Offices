using System.Collections.Generic;

namespace Bing.Offices.Models.Excels
{
    /// <summary>
    /// 数据单元行
    /// </summary>
    public class DataRow
    {
        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 数据单元格列表
        /// </summary>
        public List<DataCell> Cells { get; set; } = new List<DataCell>();

        /// <summary>
        /// 是否有效
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; } = string.Empty;
    }
}

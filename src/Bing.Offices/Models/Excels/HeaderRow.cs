using System.Collections.Generic;

namespace Bing.Offices.Models.Excels
{
    /// <summary>
    /// 表头行
    /// </summary>
    public class HeaderRow
    {
        /// <summary>
        /// 单元格
        /// </summary>
        public List<Cell> Cells { get; set; } = new List<Cell>();
    }
}

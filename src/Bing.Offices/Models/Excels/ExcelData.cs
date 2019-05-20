using System.Collections.Generic;

namespace Bing.Offices.Models.Excels
{
    /// <summary>
    /// Excel数据
    /// </summary>
    public class ExcelData
    {
        /// <summary>
        /// 表头
        /// </summary>
        public HeaderRow Header;

        /// <summary>
        /// 数据列表
        /// </summary>
        public List<DataRow> Data;
    }
}

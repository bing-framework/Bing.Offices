using Bing.Offices.Conversions;
using Bing.Offices.Npoi.Extensions;

namespace Bing.Offices.Npoi.Conversions
{
    /// <summary>
    /// 单元格值转换器
    /// </summary>
    internal class CellValueConverter : ICellValueConverter
    {
        /// <summary>
        /// 获取单元格字符串值
        /// </summary>
        /// <param name="cell">单元格</param>
        public string GetStringValue(object cell)
        {
            if (cell is NPOI.SS.UserModel.ICell npoiCell)
                return npoiCell.GetStringValue();
            return string.Empty;
        }
    }
}

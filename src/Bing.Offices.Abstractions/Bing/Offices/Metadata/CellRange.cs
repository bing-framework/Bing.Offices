namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 单元格范围
    /// </summary>
    public class CellRange
    {
        /// <summary>
        /// x1
        /// </summary>
        public int X1 { get; set; }

        /// <summary>
        /// x2
        /// </summary>
        public int X2 { get; set; }

        /// <summary>
        /// y1
        /// </summary>
        public int Y1 { get; set; }
        
        /// <summary>
        /// y2
        /// </summary>
        public int Y2 { get; set; }

        /// <summary>
        /// 列跨度
        /// </summary>
        public int ColumnSpan => X2 - X1 + 1;

        /// <summary>
        /// 行跨度
        /// </summary>
        public int RowSpan => Y2 - Y1 + 1;

        /// <summary>
        /// 是否合并单元格
        /// </summary>
        public bool IsMerged => ColumnSpan != 1 && RowSpan != 1;
    }
}

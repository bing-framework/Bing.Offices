namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 数据列元数据
    /// </summary>
    internal class ColumnMetadata
    {
        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 格式化
        /// </summary>
        public string Formatter { get; set; }

        /// <summary>
        /// 合并单元格
        /// </summary>
        public bool HasMergeCells { get; set; }

        /// <summary>
        /// 自动索引值
        /// </summary>
        public bool HasAutoIndex { get; set; }

        /// <summary>
        /// 单元格索引
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 范围
        /// </summary>
        public CellRange Range { get; set; } = new CellRange();

        /// <summary>
        /// 样式
        /// </summary>
        public ColumnStyleMetadata Style { get; set; } = new ColumnStyleMetadata();
    }
}
